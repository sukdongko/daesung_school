VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form STD010 
   Caption         =   "���л��� >> �л����"
   ClientHeight    =   11010
   ClientLeft      =   -240
   ClientTop       =   -2400
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   14850
   ScaleMode       =   0  '�����
   ScaleWidth      =   16242.14
   Begin VB.OptionButton optExmN 
      BackColor       =   &H00F7EFE7&
      Caption         =   "������"
      Height          =   285
      Left            =   2130
      TabIndex        =   8
      Top             =   2820
      Width           =   885
   End
   Begin VB.OptionButton optExmY 
      BackColor       =   &H00F7EFE7&
      Caption         =   "������"
      Height          =   285
      Left            =   1260
      TabIndex        =   7
      Top             =   2820
      Width           =   885
   End
   Begin VB.CheckBox chkSatam 
      BackColor       =   &H00F7EFE7&
      Caption         =   "�̼���"
      Height          =   345
      Index           =   12
      Left            =   7140
      TabIndex        =   201
      Top             =   5070
      Visible         =   0   'False
      Width           =   1188
   End
   Begin VB.Frame FraPay 
      BackColor       =   &H00808080&
      BorderStyle     =   0  '����
      Height          =   2295
      Left            =   15360
      TabIndex        =   173
      Top             =   540
      Width           =   5625
      Begin VB.Frame Frame7 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  '����
         Height          =   2235
         Left            =   30
         TabIndex        =   174
         Top             =   30
         Width           =   5565
         Begin VB.TextBox txtPayChg 
            Height          =   270
            IMEMode         =   10  '�ѱ� 
            Left            =   1470
            TabIndex        =   185
            Text            =   "txtPayChg"
            Top             =   810
            Width           =   1605
         End
         Begin VB.Frame Frame8 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  '����
            Height          =   435
            Left            =   1260
            TabIndex        =   183
            Top             =   1140
            Width           =   3405
            Begin VB.OptionButton OptPay1 
               BackColor       =   &H00FFFFFF&
               Caption         =   "ī��"
               Height          =   285
               Left            =   0
               TabIndex        =   176
               Top             =   90
               Width           =   885
            End
            Begin VB.OptionButton OptPay2 
               BackColor       =   &H00FFFFFF&
               Caption         =   "�ڵ���"
               Height          =   285
               Left            =   1830
               TabIndex        =   177
               Top             =   90
               Width           =   885
            End
         End
         Begin VB.TextBox txtPay 
            Height          =   270
            IMEMode         =   10  '�ѱ� 
            Left            =   1260
            TabIndex        =   175
            Text            =   "txtPay"
            Top             =   240
            Width           =   1605
         End
         Begin VB.ComboBox cboCard 
            Height          =   300
            Left            =   1230
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   178
            Top             =   1560
            Width           =   1725
         End
         Begin VB.CommandButton cmdPaySave 
            Caption         =   "����ϱ�"
            Height          =   450
            Left            =   3840
            TabIndex        =   179
            Top             =   1680
            Width           =   1395
         End
         Begin VB.Label Label43 
            Alignment       =   1  '������ ����
            BackStyle       =   0  '����
            Caption         =   "�ֹ���ȣ�� ����ÿ���"
            Height          =   210
            Left            =   900
            TabIndex        =   186
            Top             =   600
            Width           =   2175
         End
         Begin VB.Label Label57 
            Alignment       =   1  '������ ����
            BackStyle       =   0  '����
            Caption         =   "�ֹ���ȣ"
            Height          =   210
            Left            =   -150
            TabIndex        =   182
            Top             =   300
            Width           =   1185
         End
         Begin VB.Label Label55 
            Alignment       =   1  '������ ����
            BackStyle       =   0  '����
            Caption         =   "����"
            Height          =   210
            Left            =   -150
            TabIndex        =   181
            Top             =   1290
            Width           =   1185
         End
         Begin VB.Label Label58 
            Alignment       =   2  '��� ����
            BackStyle       =   0  '����
            Caption         =   "�ݱ�"
            BeginProperty Font 
               Name            =   "����"
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
            TabIndex        =   180
            Top             =   150
            Width           =   1035
         End
      End
   End
   Begin VB.Frame fraPoint 
      BackColor       =   &H00808080&
      BorderStyle     =   0  '����
      Height          =   5295
      Left            =   180
      TabIndex        =   166
      Top             =   11280
      Width           =   7185
      Begin VB.Frame Frame6 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  '����
         Height          =   5235
         Left            =   30
         TabIndex        =   167
         Top             =   30
         Width           =   7125
         Begin VB.CommandButton cmdAddPointRow 
            Caption         =   "�������� �߰�"
            Height          =   450
            Left            =   930
            TabIndex        =   171
            Top             =   4680
            Width           =   1635
         End
         Begin VB.CommandButton cmdSavePoint 
            Caption         =   "�������"
            Height          =   450
            Left            =   4380
            TabIndex        =   168
            Top             =   4680
            Width           =   2595
         End
         Begin FPSpread.vaSpread sprPoint 
            Height          =   4035
            Left            =   30
            TabIndex        =   170
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
            SpreadDesigner  =   "STD010.frx":0000
         End
         Begin VB.Label Label54 
            Alignment       =   1  '������ ����
            BackStyle       =   0  '����
            Caption         =   "������ �űԵ�ϳ��� ������ �������� ������ del Ű�� ��������."
            Height          =   210
            Left            =   -270
            TabIndex        =   172
            Top             =   270
            Width           =   5625
         End
         Begin VB.Label Label56 
            Alignment       =   2  '��� ����
            BackStyle       =   0  '����
            Caption         =   "�ݱ�"
            BeginProperty Font 
               Name            =   "����"
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
            TabIndex        =   169
            Top             =   120
            Width           =   1035
         End
      End
   End
   Begin VB.Frame fraAddr 
      BackColor       =   &H00808080&
      BorderStyle     =   0  '����
      Height          =   3015
      Left            =   15270
      TabIndex        =   155
      Top             =   6060
      Width           =   6315
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  '����
         Height          =   2955
         Left            =   30
         TabIndex        =   156
         Top             =   30
         Width           =   6255
         Begin VB.TextBox txtEmail 
            Height          =   345
            Left            =   1170
            TabIndex        =   84
            Text            =   "txtEmail"
            Top             =   1800
            Width           =   4605
         End
         Begin VB.CommandButton cmdSaveAddr 
            Caption         =   "Ȯ��"
            Height          =   450
            Left            =   4290
            TabIndex        =   85
            Top             =   2310
            Width           =   1395
         End
         Begin VB.TextBox txtAddr2 
            Height          =   345
            Left            =   1170
            TabIndex        =   83
            Text            =   "txtAddr2"
            Top             =   1380
            Width           =   4605
         End
         Begin VB.TextBox txtAddr1 
            Height          =   345
            Left            =   1170
            TabIndex        =   82
            Text            =   "txtAddr1"
            Top             =   990
            Width           =   4605
         End
         Begin EditLib.fpMask fpZip 
            Height          =   345
            Left            =   1170
            TabIndex        =   81
            Top             =   570
            Width           =   855
            _Version        =   196608
            _ExtentX        =   1508
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
            TabIndex        =   80
            Top             =   150
            Width           =   1155
            _Version        =   196608
            _ExtentX        =   2037
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
         Begin VB.Label Label51 
            Alignment       =   1  '������ ����
            BackStyle       =   0  '����
            Caption         =   "�̸���"
            Height          =   210
            Left            =   90
            TabIndex        =   162
            Top             =   1867
            Width           =   975
         End
         Begin VB.Label Label50 
            Alignment       =   1  '������ ����
            BackStyle       =   0  '����
            Caption         =   "�������"
            Height          =   210
            Left            =   90
            TabIndex        =   161
            Top             =   217
            Width           =   975
         End
         Begin VB.Label Label49 
            Alignment       =   2  '��� ����
            BackStyle       =   0  '����
            Caption         =   "�ݱ�"
            BeginProperty Font 
               Name            =   "����"
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
            TabIndex        =   160
            Top             =   180
            Width           =   1035
         End
         Begin VB.Label Label48 
            Alignment       =   1  '������ ����
            BackStyle       =   0  '����
            Caption         =   "�ּ�"
            Height          =   210
            Left            =   90
            TabIndex        =   159
            Top             =   1447
            Width           =   975
         End
         Begin VB.Label Label47 
            Alignment       =   1  '������ ����
            BackStyle       =   0  '����
            Caption         =   "�����ּ�"
            Height          =   210
            Left            =   90
            TabIndex        =   158
            Top             =   1057
            Width           =   975
         End
         Begin VB.Label Label46 
            Alignment       =   1  '������ ����
            BackStyle       =   0  '����
            Caption         =   "�����ȣ"
            Height          =   210
            Left            =   90
            TabIndex        =   157
            Top             =   637
            Width           =   975
         End
      End
   End
   Begin VB.Frame fraGwamok 
      BackColor       =   &H0082C8E8&
      BorderStyle     =   0  '����
      Caption         =   "����"
      Height          =   4275
      Left            =   2340
      TabIndex        =   136
      Top             =   11175
      Width           =   8865
      Begin VB.Frame Frame23 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  '����
         Height          =   4215
         Left            =   30
         TabIndex        =   137
         Top             =   30
         Width           =   8805
         Begin VB.CommandButton cmdClose 
            Caption         =   "�ݱ�"
            Height          =   330
            Left            =   8160
            TabIndex        =   138
            Top             =   3840
            Width           =   585
         End
         Begin VB.Image Image1 
            Height          =   4080
            Left            =   30
            Picture         =   "STD010.frx":1A0E
            Top             =   60
            Width           =   8730
         End
      End
   End
   Begin VB.Frame Frame20 
      BackColor       =   &H0082C8E8&
      BorderStyle     =   0  '����
      Caption         =   "Frame20"
      Height          =   4300
      Left            =   8460
      TabIndex        =   132
      Top             =   6150
      Width           =   6615
      Begin VB.Frame Frame21 
         BackColor       =   &H00D2EAF5&
         BorderStyle     =   0  '����
         Caption         =   "Frame21"
         Height          =   4240
         Left            =   30
         TabIndex        =   133
         Top             =   30
         Width           =   6555
         Begin VB.CommandButton cmdGwamokView 
            Caption         =   "���񺸱�"
            Enabled         =   0   'False
            Height          =   315
            Left            =   4260
            TabIndex        =   86
            Top             =   870
            Width           =   885
         End
         Begin VB.CommandButton cmdExcelSave 
            Caption         =   "�����ڷ� �����ϱ�"
            Enabled         =   0   'False
            Height          =   450
            Left            =   4590
            TabIndex        =   78
            Top             =   3765
            Width           =   1875
         End
         Begin VB.CommandButton cmdGetExcel 
            Caption         =   "�����ڷ� ��������"
            Enabled         =   0   'False
            Height          =   390
            Left            =   4410
            TabIndex        =   77
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
            TabIndex        =   87
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
            SpreadDesigner  =   "STD010.frx":90D8
         End
         Begin VB.Label Label30 
            BackStyle       =   0  '����
            Caption         =   $"STD010.frx":AD26
            Height          =   615
            Left            =   240
            TabIndex        =   135
            Top             =   630
            Width           =   5385
         End
         Begin VB.Label Label29 
            BackStyle       =   0  '����
            Caption         =   ">> ��ȸ�⺻�׸�"
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
            TabIndex        =   134
            Top             =   120
            Width           =   2625
         End
      End
   End
   Begin VB.Frame Frame18 
      BackColor       =   &H0082C8E8&
      BorderStyle     =   0  '����
      Caption         =   "Frame18"
      Height          =   6045
      Left            =   8460
      TabIndex        =   123
      Top             =   60
      Width           =   6615
      Begin VB.Frame Frame19 
         BackColor       =   &H00D2EAF5&
         BorderStyle     =   0  '����
         Caption         =   "Frame19"
         Height          =   5985
         Left            =   30
         TabIndex        =   124
         Top             =   30
         Width           =   6555
         Begin VB.TextBox Text1 
            Height          =   2655
            Left            =   930
            TabIndex        =   203
            Text            =   "Text1"
            Top             =   2730
            Visible         =   0   'False
            Width           =   3375
         End
         Begin VB.CommandButton cmdAllStdData 
            Caption         =   "������ ������ �ޱ�"
            Height          =   315
            Left            =   1740
            TabIndex        =   61
            Top             =   30
            Width           =   2955
         End
         Begin VB.ComboBox cboinGbn 
            Height          =   300
            Left            =   5220
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   65
            Top             =   450
            Width           =   885
         End
         Begin VB.ComboBox cboExmType 
            Height          =   300
            Left            =   810
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   66
            Top             =   780
            Width           =   855
         End
         Begin EditLib.fpLongInteger fpPayOK 
            Height          =   315
            Left            =   3480
            TabIndex        =   68
            Top             =   765
            Width           =   645
            _Version        =   196608
            _ExtentX        =   1138
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
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   70
            Top             =   765
            Width           =   855
         End
         Begin VB.ComboBox cboPassCN 
            Height          =   300
            Left            =   4710
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   73
            Top             =   1140
            Width           =   885
         End
         Begin VB.ComboBox cboKaeyol_F 
            Height          =   300
            Left            =   3210
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   64
            Top             =   420
            Width           =   915
         End
         Begin VB.CommandButton cmdFind 
            Caption         =   "��ȸ�ϱ�(&F)"
            Height          =   390
            Left            =   4530
            TabIndex        =   76
            Top             =   1470
            Width           =   1305
         End
         Begin VB.TextBox txtStdNM_F 
            Height          =   345
            IMEMode         =   10  '�ѱ� 
            Left            =   810
            TabIndex        =   71
            Text            =   "txtStdNM_F"
            Top             =   1125
            Width           =   825
         End
         Begin VB.ComboBox cboSel1_SCH_F 
            Height          =   300
            Left            =   810
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   74
            Top             =   1515
            Width           =   1005
         End
         Begin VB.ComboBox cboSel2_SCH_F 
            Height          =   300
            Left            =   2790
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   75
            Top             =   1515
            Width           =   1275
         End
         Begin EditLib.fpMask fpExmID_F 
            Height          =   345
            Left            =   810
            TabIndex        =   62
            Top             =   390
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
         Begin EditLib.fpMask fpBirth_ymd_F 
            Height          =   345
            Left            =   2430
            TabIndex        =   72
            Top             =   1110
            Width           =   1185
            _Version        =   196608
            _ExtentX        =   2090
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
         Begin EditLib.fpMask fpExmID_E 
            Height          =   345
            Left            =   1920
            TabIndex        =   63
            Top             =   390
            Width           =   795
            _Version        =   196608
            _ExtentX        =   1402
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
         Begin EditLib.fpLongInteger fpPayNot 
            Height          =   315
            Left            =   4710
            TabIndex        =   69
            Top             =   765
            Width           =   615
            _Version        =   196608
            _ExtentX        =   1085
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
            TabIndex        =   67
            Top             =   765
            Width           =   675
            _Version        =   196608
            _ExtentX        =   1191
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
            TabIndex        =   154
            Top             =   1890
            Width           =   6495
            _Version        =   393216
            _ExtentX        =   11456
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
            MaxCols         =   38
            ProcessTab      =   -1  'True
            SpreadDesigner  =   "STD010.frx":ADBD
         End
         Begin VB.Image imgExcel 
            Height          =   420
            Left            =   6120
            Picture         =   "STD010.frx":CFCD
            Stretch         =   -1  'True
            Top             =   1440
            Width           =   390
         End
         Begin VB.Label Label38 
            Alignment       =   1  '������ ����
            BackStyle       =   0  '����
            Caption         =   "��ü����"
            ForeColor       =   &H00C000C0&
            Height          =   210
            Left            =   1440
            TabIndex        =   146
            Top             =   810
            Width           =   975
         End
         Begin VB.Label Label37 
            Alignment       =   1  '������ ����
            BackStyle       =   0  '����
            Caption         =   "���ͳ�/�п�"
            Height          =   210
            Left            =   4110
            TabIndex        =   145
            Top             =   495
            Width           =   1095
         End
         Begin VB.Label Label36 
            Alignment       =   1  '������ ����
            BackStyle       =   0  '����
            Caption         =   "��/������"
            Height          =   210
            Left            =   -150
            TabIndex        =   144
            Top             =   825
            Width           =   975
         End
         Begin VB.Label Label35 
            Alignment       =   1  '������ ����
            BackStyle       =   0  '����
            Caption         =   "�̰���"
            ForeColor       =   &H000000C0&
            Height          =   210
            Left            =   3720
            TabIndex        =   143
            Top             =   810
            Width           =   975
         End
         Begin VB.Label Label34 
            Alignment       =   1  '������ ����
            BackStyle       =   0  '����
            Caption         =   "����"
            ForeColor       =   &H00FF0000&
            Height          =   210
            Left            =   3030
            TabIndex        =   142
            Top             =   810
            Width           =   435
         End
         Begin VB.Label Label33 
            Alignment       =   1  '������ ����
            BackStyle       =   0  '����
            Caption         =   "����"
            Height          =   210
            Left            =   5250
            TabIndex        =   141
            Top             =   810
            Width           =   465
         End
         Begin VB.Label Label32 
            Alignment       =   1  '������ ����
            BackStyle       =   0  '����
            Caption         =   "�հ�����"
            Height          =   210
            Left            =   3720
            TabIndex        =   140
            Top             =   1185
            Width           =   975
         End
         Begin VB.Label Label31 
            Alignment       =   1  '������ ����
            BackStyle       =   0  '����
            Caption         =   "��  ��"
            Height          =   210
            Left            =   2160
            TabIndex        =   139
            Top             =   465
            Width           =   1035
         End
         Begin VB.Label Label27 
            Alignment       =   1  '������ ����
            BackStyle       =   0  '����
            Caption         =   "�������"
            Height          =   210
            Left            =   1440
            TabIndex        =   130
            Top             =   1185
            Width           =   975
         End
         Begin VB.Label Label26 
            Alignment       =   1  '������ ����
            BackStyle       =   0  '����
            Caption         =   "�л���"
            Height          =   210
            Left            =   -240
            TabIndex        =   129
            Top             =   1185
            Width           =   975
         End
         Begin VB.Label Label25 
            BackStyle       =   0  '����
            Caption         =   "�����ȣ             ����"
            Height          =   210
            Left            =   30
            TabIndex        =   128
            Top             =   450
            Width           =   2025
         End
         Begin VB.Label Label24 
            BackStyle       =   0  '����
            Caption         =   ">> ��ȸ�⺻�׸�"
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
            TabIndex        =   127
            Top             =   90
            Width           =   2625
         End
         Begin VB.Label Label23 
            Alignment       =   1  '������ ����
            BackStyle       =   0  '����
            Caption         =   "1�����п�"
            Height          =   210
            Left            =   -180
            TabIndex        =   126
            Top             =   1560
            Width           =   975
         End
         Begin VB.Label Label22 
            Alignment       =   1  '������ ����
            BackStyle       =   0  '����
            Caption         =   "2�����п�"
            Height          =   210
            Left            =   1770
            TabIndex        =   125
            Top             =   1560
            Width           =   975
         End
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00C6AD84&
      BorderStyle     =   0  '����
      Caption         =   "Frame10"
      Height          =   11040
      Left            =   60
      TabIndex        =   88
      Top             =   30
      Width           =   8355
      Begin VB.Frame Frame9 
         BackColor       =   &H00F7EFE7&
         BorderStyle     =   0  '����
         Caption         =   "Frame9"
         Height          =   10965
         Left            =   30
         TabIndex        =   89
         Top             =   30
         Width           =   8295
         Begin VB.CommandButton cmdCancel 
            Caption         =   "�л� ����ϱ�"
            Height          =   450
            Left            =   4440
            TabIndex        =   55
            Top             =   10095
            Width           =   1815
         End
         Begin VB.CommandButton cmdStdDel 
            Caption         =   "�л������ϱ�"
            Height          =   450
            Left            =   6720
            TabIndex        =   56
            Top             =   10095
            Width           =   1365
         End
         Begin VB.CommandButton cmdStdin 
            Caption         =   "�л���� �� �����ϱ� (&S)"
            Height          =   450
            Left            =   900
            TabIndex        =   54
            Top             =   10095
            Width           =   2655
         End
         Begin VB.Frame Frame17 
            BackColor       =   &H0082C8E8&
            BorderStyle     =   0  '����
            Caption         =   "Frame17"
            Height          =   825
            Index           =   0
            Left            =   30
            TabIndex        =   114
            Top             =   9150
            Width           =   8235
            Begin VB.Frame fraSEL5 
               BackColor       =   &H00F7EFE7&
               BorderStyle     =   0  '����
               Caption         =   ">> ��� ���ð���"
               Height          =   765
               Index           =   0
               Left            =   30
               TabIndex        =   115
               Top             =   30
               Width           =   8175
               Begin VB.CheckBox chkNonsul 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "���"
                  Height          =   345
                  Index           =   1
                  Left            =   240
                  TabIndex        =   50
                  Top             =   360
                  Width           =   1245
               End
               Begin VB.CheckBox chkNonsul 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "����"
                  Height          =   345
                  Index           =   2
                  Left            =   1590
                  TabIndex        =   51
                  Top             =   360
                  Width           =   1245
               End
               Begin VB.CheckBox chkNonsul 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "�ܱ���"
                  Height          =   345
                  Index           =   3
                  Left            =   2970
                  TabIndex        =   52
                  Top             =   360
                  Width           =   1245
               End
               Begin VB.CheckBox chkNonsul 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "����Ž��"
                  Height          =   345
                  Index           =   4
                  Left            =   4290
                  TabIndex        =   53
                  Top             =   360
                  Visible         =   0   'False
                  Width           =   1245
               End
               Begin VB.Label Label15 
                  BackStyle       =   0  '����
                  Caption         =   ">> Ž������"
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
                  Index           =   0
                  Left            =   90
                  TabIndex        =   116
                  Top             =   90
                  Width           =   2625
               End
            End
         End
         Begin VB.Frame Frame16 
            BackColor       =   &H0082C8E8&
            BorderStyle     =   0  '����
            Caption         =   "Frame16"
            Height          =   825
            Left            =   30
            TabIndex        =   111
            Top             =   8310
            Width           =   8235
            Begin VB.Frame fraSEL4 
               BackColor       =   &H00F7EFE7&
               BorderStyle     =   0  '����
               Caption         =   ">> �������� ���ð���"
               Height          =   765
               Left            =   30
               TabIndex        =   112
               Top             =   30
               Width           =   8175
               Begin VB.CheckBox chkMath 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "������"
                  Height          =   345
                  Index           =   1
                  Left            =   240
                  TabIndex        =   46
                  Top             =   390
                  Width           =   1245
               End
               Begin VB.CheckBox chkMath 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "�̻����"
                  Height          =   345
                  Index           =   2
                  Left            =   1590
                  TabIndex        =   47
                  Top             =   390
                  Width           =   1245
               End
               Begin VB.CheckBox chkMath 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "Ȯ�����"
                  Height          =   345
                  Index           =   3
                  Left            =   2940
                  TabIndex        =   48
                  Top             =   390
                  Width           =   1245
               End
               Begin VB.CheckBox chkMath 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "��������"
                  Height          =   345
                  Index           =   4
                  Left            =   4290
                  TabIndex        =   49
                  Top             =   390
                  Width           =   1245
               End
               Begin VB.Label Label14 
                  BackStyle       =   0  '����
                  Caption         =   ">> �������� ���ð���"
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
                  Left            =   90
                  TabIndex        =   113
                  Top             =   90
                  Width           =   2625
               End
            End
         End
         Begin VB.Frame Frame15 
            BackColor       =   &H0082C8E8&
            BorderStyle     =   0  '����
            Caption         =   "Frame15"
            Height          =   1215
            Left            =   60
            TabIndex        =   108
            Top             =   7050
            Width           =   8235
            Begin VB.Frame fraSEL3 
               BackColor       =   &H00F7EFE7&
               BorderStyle     =   0  '����
               Caption         =   ">> ����Ž�� ���ð���"
               Height          =   1155
               Left            =   30
               TabIndex        =   109
               Top             =   30
               Width           =   8175
               Begin VB.CheckBox chkGwatam 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "Ư��"
                  Height          =   345
                  Index           =   9
                  Left            =   5760
                  TabIndex        =   209
                  Top             =   780
                  Width           =   1245
               End
               Begin VB.CheckBox chkGwatam 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "����1"
                  Height          =   345
                  Index           =   1
                  Left            =   240
                  TabIndex        =   38
                  Top             =   360
                  Width           =   1245
               End
               Begin VB.CheckBox chkGwatam 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "ȭ��1"
                  Height          =   345
                  Index           =   2
                  Left            =   1620
                  TabIndex        =   39
                  Top             =   360
                  Width           =   1245
               End
               Begin VB.CheckBox chkGwatam 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "�������1"
                  Height          =   345
                  Index           =   3
                  Left            =   2970
                  TabIndex        =   40
                  Top             =   360
                  Width           =   1245
               End
               Begin VB.CheckBox chkGwatam 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "��������1"
                  Height          =   345
                  Index           =   4
                  Left            =   4320
                  TabIndex        =   41
                  Top             =   360
                  Width           =   1245
               End
               Begin VB.CheckBox chkGwatam 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "����2"
                  Height          =   345
                  Index           =   5
                  Left            =   240
                  TabIndex        =   42
                  Top             =   780
                  Width           =   1245
               End
               Begin VB.CheckBox chkGwatam 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "ȭ��2"
                  Height          =   345
                  Index           =   6
                  Left            =   1620
                  TabIndex        =   43
                  Top             =   780
                  Width           =   1245
               End
               Begin VB.CheckBox chkGwatam 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "�������2"
                  Height          =   345
                  Index           =   7
                  Left            =   2970
                  TabIndex        =   44
                  Top             =   780
                  Width           =   1245
               End
               Begin VB.CheckBox chkGwatam 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "��������2"
                  Height          =   345
                  Index           =   8
                  Left            =   4320
                  TabIndex        =   45
                  Top             =   780
                  Width           =   1245
               End
               Begin VB.Label Label13 
                  BackStyle       =   0  '����
                  Caption         =   ">> ����Ž�� ���ð���"
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
                  Left            =   90
                  TabIndex        =   110
                  Top             =   90
                  Width           =   2625
               End
            End
         End
         Begin VB.Frame Frame14 
            BackColor       =   &H0082C8E8&
            BorderStyle     =   0  '����
            Caption         =   "Frame14"
            Height          =   855
            Left            =   0
            TabIndex        =   105
            Top             =   6180
            Width           =   8235
            Begin VB.Frame fraSEL2 
               BackColor       =   &H00F7EFE7&
               BorderStyle     =   0  '����
               Caption         =   ">> ��2 �ܱ��� ���ð���"
               Height          =   795
               Left            =   30
               TabIndex        =   106
               Top             =   30
               Width           =   8175
               Begin VB.CheckBox chkEng2 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "�ƶ���"
                  Height          =   345
                  Index           =   12
                  Left            =   7140
                  TabIndex        =   37
                  Top             =   510
                  Width           =   1245
               End
               Begin VB.CheckBox chkEng2 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "��������"
                  Height          =   345
                  Index           =   11
                  Left            =   5820
                  TabIndex        =   36
                  Top             =   510
                  Width           =   1245
               End
               Begin VB.CheckBox chkEng2 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "�����"
                  Height          =   345
                  Index           =   10
                  Left            =   4320
                  TabIndex        =   35
                  Top             =   510
                  Width           =   1245
               End
               Begin VB.CheckBox chkEng2 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "�ܱ���"
                  Height          =   345
                  Index           =   9
                  Left            =   2970
                  TabIndex        =   34
                  Top             =   510
                  Width           =   1245
               End
               Begin VB.CheckBox chkEng2 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "����"
                  Height          =   345
                  Index           =   8
                  Left            =   1620
                  TabIndex        =   33
                  Top             =   510
                  Width           =   1245
               End
               Begin VB.CheckBox chkEng2 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "���"
                  Height          =   345
                  Index           =   7
                  Left            =   240
                  TabIndex        =   32
                  Top             =   510
                  Width           =   1245
               End
               Begin VB.CheckBox chkEng2 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "����"
                  Height          =   345
                  Index           =   1
                  Left            =   240
                  TabIndex        =   26
                  Top             =   240
                  Width           =   1245
               End
               Begin VB.CheckBox chkEng2 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "�Ͼ�"
                  Height          =   345
                  Index           =   2
                  Left            =   1620
                  TabIndex        =   27
                  Top             =   240
                  Width           =   1245
               End
               Begin VB.CheckBox chkEng2 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "�����ĳľ�"
                  Height          =   345
                  Index           =   3
                  Left            =   2970
                  TabIndex        =   28
                  Top             =   240
                  Width           =   1245
               End
               Begin VB.CheckBox chkEng2 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "�Ҿ�"
                  Height          =   345
                  Index           =   4
                  Left            =   4320
                  TabIndex        =   29
                  Top             =   240
                  Width           =   1245
               End
               Begin VB.CheckBox chkEng2 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "�߱���"
                  Height          =   345
                  Index           =   5
                  Left            =   5820
                  TabIndex        =   30
                  Top             =   240
                  Width           =   1245
               End
               Begin VB.CheckBox chkEng2 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "�ѹ�"
                  Height          =   345
                  Index           =   6
                  Left            =   7140
                  TabIndex        =   31
                  Top             =   240
                  Width           =   1245
               End
               Begin VB.Label Label12 
                  BackStyle       =   0  '����
                  Caption         =   ">> ��2 �ܱ��� ���ð���"
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
                  Left            =   90
                  TabIndex        =   107
                  Top             =   60
                  Width           =   2625
               End
            End
         End
         Begin VB.Frame Frame13 
            BackColor       =   &H0082C8E8&
            BorderStyle     =   0  '����
            Caption         =   "Frame13"
            Height          =   1215
            Left            =   30
            TabIndex        =   104
            Top             =   4950
            Width           =   8235
            Begin VB.Frame fraSEL1 
               BackColor       =   &H00F7EFE7&
               BorderStyle     =   0  '����
               Caption         =   ">> ��ȸŽ�� ���ð���"
               Height          =   1155
               Left            =   30
               TabIndex        =   189
               Top             =   30
               Width           =   8175
               Begin VB.CheckBox chkSatam 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "Ư��"
                  Height          =   345
                  Index           =   11
                  Left            =   7080
                  TabIndex        =   208
                  Top             =   750
                  Width           =   1245
               End
               Begin VB.CheckBox chkSatam 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "�ѱ���"
                  Height          =   345
                  Index           =   1
                  Left            =   240
                  TabIndex        =   199
                  Top             =   330
                  Width           =   1245
               End
               Begin VB.CheckBox chkSatam 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "�����"
                  Height          =   345
                  Index           =   2
                  Left            =   1620
                  TabIndex        =   198
                  Top             =   330
                  Width           =   1245
               End
               Begin VB.CheckBox chkSatam 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "���ƽþƻ�"
                  Height          =   345
                  Index           =   3
                  Left            =   2970
                  TabIndex        =   197
                  Top             =   330
                  Width           =   1245
               End
               Begin VB.CheckBox chkSatam 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "�ѱ�����"
                  Height          =   345
                  Index           =   4
                  Left            =   4320
                  TabIndex        =   196
                  Top             =   330
                  Width           =   1245
               End
               Begin VB.CheckBox chkSatam 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "��������"
                  Height          =   345
                  Index           =   5
                  Left            =   5790
                  TabIndex        =   195
                  Top             =   330
                  Width           =   1245
               End
               Begin VB.CheckBox chkSatam 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "��Ȱ������"
                  Height          =   345
                  Index           =   6
                  Left            =   240
                  TabIndex        =   194
                  Top             =   750
                  Width           =   1245
               End
               Begin VB.CheckBox chkSatam 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "�����ͻ��"
                  Height          =   345
                  Index           =   7
                  Left            =   1620
                  TabIndex        =   193
                  Top             =   750
                  Width           =   1245
               End
               Begin VB.CheckBox chkSatam 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "������ġ"
                  Height          =   345
                  Index           =   8
                  Left            =   2970
                  TabIndex        =   192
                  Top             =   750
                  Width           =   1245
               End
               Begin VB.CheckBox chkSatam 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "����"
                  Height          =   345
                  Index           =   9
                  Left            =   4320
                  TabIndex        =   191
                  Top             =   750
                  Width           =   1245
               End
               Begin VB.CheckBox chkSatam 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "��ȸ��ȭ"
                  Height          =   345
                  Index           =   10
                  Left            =   5790
                  TabIndex        =   190
                  Top             =   750
                  Width           =   1245
               End
               Begin VB.Label Label11 
                  BackStyle       =   0  '����
                  Caption         =   ">> ��ȸŽ�� ���ð���"
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
                  Left            =   60
                  TabIndex        =   200
                  Top             =   90
                  Width           =   2625
               End
            End
         End
         Begin VB.Frame Frame12 
            BackColor       =   &H00C6AD84&
            BorderStyle     =   0  '����
            Caption         =   "Frame12"
            Height          =   1245
            Left            =   30
            TabIndex        =   98
            Top             =   3645
            Width           =   8235
            Begin VB.Frame Frame4 
               BackColor       =   &H00F7EFE7&
               BorderStyle     =   0  '����
               Caption         =   ">> ����"
               Height          =   1185
               Left            =   30
               TabIndex        =   99
               Top             =   30
               Width           =   8175
               Begin VB.CommandButton cmdAddPoint 
                  Caption         =   "�л�������(&P)"
                  BeginProperty Font 
                     Name            =   "����"
                     Size            =   9
                     Charset         =   129
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   345
                  Left            =   4260
                  TabIndex        =   165
                  Top             =   750
                  Width           =   1725
               End
               Begin VB.CommandButton cmdChgAddr 
                  Caption         =   "�л��󼼺���"
                  BeginProperty Font 
                     Name            =   "����"
                     Size            =   9
                     Charset         =   129
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   345
                  Left            =   6390
                  TabIndex        =   79
                  Top             =   750
                  Width           =   1665
               End
               Begin VB.TextBox txtCancel 
                  Enabled         =   0   'False
                  Height          =   270
                  IMEMode         =   10  '�ѱ� 
                  Left            =   6390
                  TabIndex        =   22
                  Text            =   "txtCancel"
                  Top             =   0
                  Width           =   1695
               End
               Begin EditLib.fpLongInteger fpK_Num 
                  Height          =   345
                  Left            =   600
                  TabIndex        =   23
                  Top             =   300
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
                  TabIndex        =   25
                  Top             =   300
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
                  TabIndex        =   24
                  Top             =   300
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
               Begin EditLib.fpLongInteger fpN_Num 
                  Height          =   345
                  Left            =   5220
                  TabIndex        =   206
                  Top             =   300
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
                  BackStyle       =   0  '����
                  Caption         =   "���ŵ��"
                  Height          =   195
                  Left            =   4410
                  TabIndex        =   207
                  Top             =   360
                  Width           =   885
               End
               Begin VB.Label Label44 
                  Alignment       =   1  '������ ����
                  BackStyle       =   0  '����
                  Caption         =   "��ҿ���"
                  ForeColor       =   &H00C000C0&
                  Height          =   180
                  Left            =   5010
                  TabIndex        =   152
                  Top             =   45
                  Width           =   1365
               End
               Begin VB.Label Label10 
                  BackStyle       =   0  '����
                  Caption         =   ">> ����"
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
                  Left            =   60
                  TabIndex        =   103
                  Top             =   30
                  Width           =   2625
               End
               Begin VB.Label Label6 
                  Alignment       =   1  '������ ����
                  BackStyle       =   0  '����
                  Caption         =   "���"
                  Height          =   210
                  Left            =   -390
                  TabIndex        =   102
                  Top             =   360
                  Width           =   975
               End
               Begin VB.Label Label7 
                  Alignment       =   1  '������ ����
                  BackStyle       =   0  '����
                  Caption         =   "�ܱ���"
                  Height          =   210
                  Left            =   2430
                  TabIndex        =   101
                  Top             =   360
                  Width           =   975
               End
               Begin VB.Label Label8 
                  Alignment       =   1  '������ ����
                  BackStyle       =   0  '����
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
                  Height          =   210
                  Left            =   930
                  TabIndex        =   100
                  Top             =   360
                  Width           =   975
               End
            End
         End
         Begin VB.Frame Frame11 
            BackColor       =   &H00C6AD84&
            BorderStyle     =   0  '����
            Caption         =   "Frame11"
            Height          =   3500
            Left            =   30
            TabIndex        =   90
            Top             =   90
            Width           =   8235
            Begin VB.Frame Frame3 
               BackColor       =   &H00F7EFE7&
               BorderStyle     =   0  '����
               Caption         =   ">> �⺻�׸�"
               Height          =   3450
               Left            =   30
               TabIndex        =   91
               Top             =   30
               Width           =   8175
               Begin VB.TextBox txt_P_Phone 
                  Height          =   270
                  Left            =   3780
                  TabIndex        =   15
                  Text            =   "Text1"
                  Top             =   2300
                  Width           =   1455
               End
               Begin VB.TextBox txt_MAJOR 
                  Height          =   285
                  Left            =   6345
                  TabIndex        =   17
                  Text            =   "Text1"
                  Top             =   2715
                  Width           =   1725
               End
               Begin VB.TextBox txt_UNI 
                  Height          =   285
                  Left            =   3765
                  TabIndex        =   16
                  Text            =   "Text1"
                  Top             =   2700
                  Width           =   1700
               End
               Begin VB.CommandButton cmdPayChg 
                  Caption         =   "������"
                  BeginProperty Font 
                     Name            =   "����"
                     Size            =   9
                     Charset         =   129
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   5445
                  TabIndex        =   184
                  Top             =   3090
                  Width           =   1245
               End
               Begin VB.ComboBox cboMu_type 
                  Height          =   300
                  Left            =   3780
                  Style           =   2  '��Ӵٿ� ���
                  TabIndex        =   9
                  Top             =   30
                  Width           =   1725
               End
               Begin VB.ComboBox cboPTS_Sel 
                  Height          =   300
                  Left            =   6810
                  Style           =   2  '��Ӵٿ� ���
                  TabIndex        =   19
                  Top             =   30
                  Width           =   1275
               End
               Begin VB.CommandButton cmdNew 
                  Caption         =   "�ű� (&S)"
                  Height          =   315
                  Left            =   1350
                  TabIndex        =   0
                  Top             =   -30
                  Width           =   1125
               End
               Begin VB.TextBox txtPayGbn 
                  Enabled         =   0   'False
                  Height          =   270
                  IMEMode         =   10  '�ѱ� 
                  Left            =   6795
                  TabIndex        =   21
                  Text            =   "txtPayGbn"
                  Top             =   3090
                  Width           =   1275
               End
               Begin VB.TextBox txtRegDate 
                  Enabled         =   0   'False
                  Height          =   270
                  IMEMode         =   10  '�ѱ� 
                  Left            =   3750
                  TabIndex        =   18
                  Text            =   "txtRegDate"
                  Top             =   3060
                  Width           =   1725
               End
               Begin VB.TextBox txtCel 
                  Height          =   270
                  IMEMode         =   10  '�ѱ� 
                  Left            =   3780
                  TabIndex        =   14
                  Text            =   "txtCel"
                  Top             =   1875
                  Width           =   1455
               End
               Begin VB.TextBox txtOrdNo 
                  Enabled         =   0   'False
                  Height          =   270
                  IMEMode         =   10  '�ѱ� 
                  Left            =   6810
                  TabIndex        =   20
                  Text            =   "txtOrdNo"
                  Top             =   2445
                  Width           =   1275
               End
               Begin VB.TextBox txtTel 
                  Height          =   270
                  IMEMode         =   10  '�ѱ� 
                  Left            =   3780
                  TabIndex        =   13
                  Text            =   "9999-9999-9999"
                  Top             =   1560
                  Width           =   1455
               End
               Begin VB.ComboBox cboKaeyol 
                  Height          =   300
                  Left            =   3780
                  Style           =   2  '��Ӵٿ� ���
                  TabIndex        =   10
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
               Begin VB.ComboBox cboPass4 
                  Height          =   300
                  Left            =   6810
                  Style           =   2  '��Ӵٿ� ���
                  TabIndex        =   60
                  Top             =   1612
                  Width           =   1275
               End
               Begin VB.ComboBox cboPass3 
                  Height          =   300
                  Left            =   6810
                  Style           =   2  '��Ӵٿ� ���
                  TabIndex        =   59
                  Top             =   1192
                  Width           =   1275
               End
               Begin VB.ComboBox cboPass2 
                  Height          =   300
                  Left            =   6810
                  Style           =   2  '��Ӵٿ� ���
                  TabIndex        =   58
                  Top             =   772
                  Width           =   1275
               End
               Begin VB.ComboBox cboPass1 
                  Height          =   300
                  Left            =   6810
                  Style           =   2  '��Ӵٿ� ���
                  TabIndex        =   57
                  Top             =   352
                  Width           =   1275
               End
               Begin VB.ComboBox cboSel2_Sch 
                  Height          =   300
                  Left            =   3780
                  Style           =   2  '��Ӵٿ� ���
                  TabIndex        =   12
                  Top             =   1192
                  Width           =   1725
               End
               Begin VB.ComboBox cboSel1_Sch 
                  Height          =   300
                  Left            =   3780
                  Style           =   2  '��Ӵٿ� ���
                  TabIndex        =   11
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
                  IMEMode         =   10  '�ѱ� 
                  Left            =   1140
                  TabIndex        =   3
                  Text            =   "txtStdNM"
                  Top             =   1170
                  Width           =   1605
               End
               Begin VB.Frame Frame1 
                  BackColor       =   &H00F7EFE7&
                  BorderStyle     =   0  '����
                  Height          =   435
                  Left            =   1140
                  TabIndex        =   92
                  Top             =   2025
                  Width           =   1800
                  Begin VB.OptionButton optSexFemale 
                     BackColor       =   &H00F7EFE7&
                     Caption         =   "����"
                     Height          =   285
                     Left            =   840
                     TabIndex        =   6
                     Top             =   90
                     Width           =   885
                  End
                  Begin VB.OptionButton optSexMale 
                     BackColor       =   &H00F7EFE7&
                     Caption         =   "����"
                     Height          =   285
                     Left            =   0
                     TabIndex        =   5
                     Top             =   90
                     Width           =   885
                  End
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
               Begin VB.Label Label5 
                  Alignment       =   1  '������ ����
                  BackStyle       =   0  '����
                  Caption         =   "����"
                  Height          =   210
                  Left            =   90
                  TabIndex        =   205
                  Top             =   2160
                  Width           =   975
               End
               Begin VB.Label Label62 
                  Alignment       =   1  '������ ����
                  BackStyle       =   0  '����
                  Caption         =   "��/������"
                  Height          =   210
                  Left            =   90
                  TabIndex        =   204
                  Top             =   2670
                  Width           =   975
               End
               Begin VB.Label Label61 
                  Caption         =   "�кθ�HP"
                  Height          =   255
                  Left            =   3000
                  TabIndex        =   202
                  Top             =   2280
                  Width           =   880
               End
               Begin VB.Label Label60 
                  BackStyle       =   0  '����
                  Caption         =   "�����ܴ�"
                  Height          =   300
                  Left            =   5520
                  TabIndex        =   188
                  Top             =   2760
                  Width           =   750
               End
               Begin VB.Label Label59 
                  Caption         =   "��������"
                  Height          =   225
                  Left            =   2970
                  TabIndex        =   187
                  Top             =   2760
                  Width           =   810
               End
               Begin VB.Label Label53 
                  Alignment       =   1  '������ ����
                  BackStyle       =   0  '����
                  Caption         =   "���"
                  Height          =   210
                  Left            =   2520
                  TabIndex        =   164
                  Top             =   105
                  Width           =   1185
               End
               Begin VB.Label Label52 
                  Alignment       =   1  '������ ����
                  BackStyle       =   0  '����
                  Caption         =   "������������"
                  Height          =   210
                  Left            =   5550
                  TabIndex        =   163
                  Top             =   105
                  Width           =   1185
               End
               Begin VB.Label Label42 
                  Alignment       =   1  '������ ����
                  BackStyle       =   0  '����
                  Caption         =   "�������"
                  ForeColor       =   &H00C000C0&
                  Height          =   180
                  Left            =   2310
                  TabIndex        =   151
                  Top             =   3090
                  Width           =   1365
               End
               Begin VB.Label Label41 
                  Alignment       =   1  '������ ����
                  BackStyle       =   0  '����
                  Caption         =   "�ڵ���"
                  ForeColor       =   &H00000000&
                  Height          =   210
                  Left            =   2730
                  TabIndex        =   149
                  Top             =   1890
                  Width           =   975
               End
               Begin VB.Label Label40 
                  Alignment       =   1  '������ ����
                  BackStyle       =   0  '����
                  Caption         =   "�ֹ���ȣ(��ȸ)"
                  ForeColor       =   &H00C000C0&
                  Height          =   180
                  Left            =   5370
                  TabIndex        =   148
                  Top             =   2490
                  Width           =   1365
               End
               Begin VB.Label Label39 
                  Alignment       =   1  '������ ����
                  BackStyle       =   0  '����
                  Caption         =   "TEL"
                  ForeColor       =   &H00000000&
                  Height          =   210
                  Left            =   2730
                  TabIndex        =   147
                  Top             =   1620
                  Width           =   975
               End
               Begin VB.Label Label28 
                  Alignment       =   1  '������ ����
                  BackStyle       =   0  '����
                  Caption         =   "��  ��"
                  Height          =   210
                  Left            =   2760
                  TabIndex        =   131
                  Top             =   390
                  Width           =   975
               End
               Begin VB.Label Label21 
                  Alignment       =   1  '������ ����
                  BackStyle       =   0  '����
                  Caption         =   "4���� �հ��п�"
                  Height          =   210
                  Left            =   5280
                  TabIndex        =   122
                  Top             =   1650
                  Width           =   1455
               End
               Begin VB.Label Label20 
                  Alignment       =   1  '������ ����
                  BackStyle       =   0  '����
                  Caption         =   "3���� �հ��п�"
                  Height          =   210
                  Left            =   5280
                  TabIndex        =   121
                  Top             =   1230
                  Width           =   1455
               End
               Begin VB.Label Label19 
                  Alignment       =   1  '������ ����
                  BackStyle       =   0  '����
                  Caption         =   "2���� �հ��п�"
                  Height          =   210
                  Left            =   5280
                  TabIndex        =   120
                  Top             =   810
                  Width           =   1455
               End
               Begin VB.Label Label18 
                  Alignment       =   1  '������ ����
                  BackStyle       =   0  '����
                  Caption         =   "1���� �հ��п�"
                  Height          =   210
                  Left            =   5280
                  TabIndex        =   119
                  Top             =   390
                  Width           =   1455
               End
               Begin VB.Label Label17 
                  Alignment       =   1  '������ ����
                  BackStyle       =   0  '����
                  Caption         =   "2���� �п�"
                  Height          =   210
                  Left            =   2760
                  TabIndex        =   118
                  Top             =   1230
                  Width           =   975
               End
               Begin VB.Label Label16 
                  Alignment       =   1  '������ ����
                  BackStyle       =   0  '����
                  Caption         =   "1���� �п�"
                  Height          =   210
                  Left            =   2760
                  TabIndex        =   117
                  Top             =   810
                  Width           =   975
               End
               Begin VB.Label Label9 
                  BackStyle       =   0  '����
                  Caption         =   ">> �⺻�׸�"
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
                  Left            =   90
                  TabIndex        =   97
                  Top             =   60
                  Width           =   2625
               End
               Begin VB.Label Label4 
                  Alignment       =   1  '������ ����
                  BackStyle       =   0  '����
                  Caption         =   "�ý��� �ڵ�"
                  Height          =   210
                  Left            =   60
                  TabIndex        =   96
                  Top             =   390
                  Width           =   975
               End
               Begin VB.Label Label1 
                  Alignment       =   1  '������ ����
                  BackStyle       =   0  '����
                  Caption         =   "�����ȣ"
                  Height          =   210
                  Left            =   60
                  TabIndex        =   95
                  Top             =   810
                  Width           =   975
               End
               Begin VB.Label Label2 
                  Alignment       =   1  '������ ����
                  BackStyle       =   0  '����
                  Caption         =   "�л���"
                  Height          =   210
                  Left            =   60
                  TabIndex        =   94
                  Top             =   1230
                  Width           =   975
               End
               Begin VB.Label Label3 
                  Alignment       =   1  '������ ����
                  BackStyle       =   0  '����
                  Caption         =   "�������"
                  Height          =   210
                  Left            =   60
                  TabIndex        =   93
                  Top             =   1650
                  Width           =   975
               End
            End
         End
         Begin VB.Label Label45 
            BackStyle       =   0  '����
            Caption         =   "�� �л������� �߸� �Է��� ��츸 ����Ͻʽÿ�."
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   3870
            TabIndex        =   153
            Top             =   10650
            Width           =   4365
         End
      End
   End
   Begin FPSpread.vaSpread sprStdData 
      Height          =   165
      Left            =   8430
      TabIndex        =   150
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
      SpreadDesigner  =   "STD010.frx":D40E
   End
End
Attribute VB_Name = "STD010"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################################################
'   �� �� ��  �� : �뼺�п� ���л���, �ݹ��� & �ð�ǥ ���α׷�
'   ����ý��۸� :
'   ��   ��   �� : STD010
'   �� ��  �� �� : �л� ��� �� ��ȸ
'
'   ��   ��   �� : 2007/08/22
'   ��   ��   �� : ���ϱ�
' --------------------------------------------------------------------------------------------------------------
' --------------------------------------------------------------------------------------------------------------
'                 ��     ��     ��     ��
' --------------------------------------------------------------------------------------------------------------
'   1. ������ :
'   2. ��  �� :
'################################################################################################################

Option Explicit


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




Private Sub Form_Terminate()
    
    Unload Me
End Sub


Private Sub Form_Load()
        
    Me.Move 0, 0, 15255, 11670
    
    fraGwamok.Visible = False '�� �Ʒ� �����ڵ� ������


    '>>>>>>>>>>>>��� �� �ʱ�ȭ <<<<<<<<<<<<<<
    txtSchNo.Text = ""
    fpExmID.Text = ""
    txtStdNM.Text = ""
    fpBirth_ymd.Text = ""   '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< ���� ����
    txt_P_Phone = ""    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< ���� ����
    txt_UNI.Text = ""   '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< ���� ����
    txt_MAJOR.Text = ""   '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< ���� ����
    txtTel.Text = ""
    txtCel.Text = ""
    txtRegDate.Text = ""
    txtOrdNo.Text = ""
    txtPayGbn.Text = ""

    txtCancel.Text = ""

    optExmY.value = True
    optSexMale.value = True
    
    fpK_Num.value = 0
    fpE_Num.value = 0
    fpM_Num.value = 0
    fpN_Num.value = 0

    '����â.
    fpBirth_ymdS.Text = ""
    fpZip.Text = ""
    txtAddr1.Text = ""
    txtAddr2.Text = ""
    txtEmail.Text = ""
    
    Call basCommonSTD.Init_CboKaeyolDefault(cboKaeyol)      '�迭
    Call basCommonSTD.Init_CboSch(cboSel1_Sch)   '1���� �п�
    Call basCommonSTD.Init_CboSch(cboSel2_Sch)   '2���� �п�
    Call basCommonSTD.Init_CboSch(cboPass1)      '1���� �հ� �п�
    Call basCommonSTD.Init_CboSch(cboPass2)      '2���� �հ� �п�
    Call basCommonSTD.Init_CboSch(cboPass3)      '3���� �հ� �п�
    Call basCommonSTD.Init_CboSch(cboPass4)      '4���� �հ� �п�
    Call basCommonSTD.Init_Mu_type(cboMu_type)       '���
    Call basCommonSTD.Init_PTS_Sel(cboPTS_Sel)       '������������
    Call basCommonSTD.Init_Card(cboCard)             'ī��
'    Call basCommonSTD.Init_Clinic(cbo_Clinic_L, cbo_Clinic_M, cbo_Clinic_E)     'Ŭ����
    
    
    '>>>>>>>>>>>> ��ȸ �� �ʱ�ȭ <<<<<<<<<<<<<<
    fpExmID_F.Text = ""
    fpExmID_E.Text = ""
    
    txtStdNM_F.Text = ""
    fpBirth_ymd_F.Text = ""
    sprSTD_F.MaxRows = 0
    
    sprExcel_STD_Data.MaxRows = 0
    
    fpPayOK.value = 0
    fpPayNot.value = 0
    fpPayTot.value = 0
    
    
    Call basCommonSTD.Init_CboKaeyolDefault(cboKaeyol_F)    '��ȸ �迭
    cboKaeyol_F.AddItem "��ü" & Space(30) & "ALL", 0
    cboKaeyol_F.ListIndex = 0
    
    Call basCommonSTD.Init_InGbn(cboinGbn)           '��ȸ ���ͳ�/�п�
    Call basCommonSTD.Init_ExmType(cboExmType)       '��ȸ ���������
    Call basCommonSTD.Init_Pay(cboPay)               '��ȸ ����
    Call basCommonSTD.Init_PassCN(cboPassCN)         '��ȸ �հ�����
    Call basCommonSTD.Init_CboSch(cboSel1_SCH_F)        '��ȸ 1���� �п�
    Call basCommonSTD.Init_CboSch(cboSel2_SCH_F)        '��ȸ 2���� �п�
    
    Call basCommonSTD.Set_Spread_Design1(sprSTD_F)              '�л���ȸ ��Ʈ
    Call basCommonSTD.Set_Spread_Design1(sprExcel_STD_Data)     '������������ ��Ʈ
    
    '>>>>>>>>>>>> ����â�� ��ġ����. <<<<<<<<<<<<<<
    With fraAddr        '< �л� �󼼳��� ��� : 2008.01.10
        .Top = 3420
        .Left = 6540
        
        .ZOrder 0
        .Visible = False
    End With
    With fraPoint       '< �л� ������ ��� : 2008.01.10
        .Top = 3420
        .Left = 4500

        .ZOrder 0
        .Visible = False
    End With
    
    With FraPay         '< �������� ��� : 2010.01.13
'        .Top = 3420
'        .Left = 4500

        .ZOrder 0
        .Visible = False
    End With
    
    
    '>>>>>>>>>>>> �п��� ���� �� ���� <<<<<<<<<<<<<<
    Dim ni As Integer
    
    '>> �����ϰ�� Ư�� ǥ��
    chkGwatam(9).Visible = False
    chkSatam(11).Visible = False
    If basModule.SchCD = "J" Then
        chkGwatam(9).Visible = True
        chkSatam(11).Visible = True
    End If
    
    '>> 1���� �п�
    Call basCommonSTD.Set_CboSch(cboSel1_Sch, basModule.SchCD)
    
    
    '>> �п�
    Select Case Trim(basModule.SchCD)
        Case "N"        '�뷮��
            For ni = 7 To 9 Step 1
                chkEng2(ni).Visible = True
            Next ni
            For ni = 10 To 11 Step 1
                chkEng2(ni).Visible = False
            Next ni
            
            chkEng2(12).Visible = True
            
        Case "M"        '���� ���̸�
            For ni = 7 To 9 Step 1
                chkEng2(ni).Visible = True
            Next ni
            For ni = 10 To 11 Step 1
                chkEng2(ni).Visible = False
            Next ni
            
            chkEng2(12).Visible = True
            
        Case "S"        '����
            
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
            chkEng2(9).Visible = True
            
            chkEng2(10).Visible = False     ' True
            chkEng2(11).Visible = False     ' True
            
            chkEng2(12).Visible = True
            
        Case "J"        '����
            For ni = 7 To 9 Step 1
                chkEng2(ni).Visible = True
            Next ni
            For ni = 10 To 11 Step 1
                chkEng2(ni).Visible = False
            Next ni
            
            chkEng2(12).Visible = True
            
        Case Else
            For ni = 7 To 11 Step 1
                chkEng2(ni).Visible = False
            Next ni
            
            chkEng2(12).Visible = True
    End Select
    
    
    
    
End Sub


'>> �ű�
Private Sub cmdNew_Click()

    Dim ni      As Integer

    txtSchNo.Text = ""
    fpExmID.Text = ""
    txtStdNM.Text = ""
    fpBirth_ymd.Text = ""   '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< ���� ����
    txt_P_Phone = ""    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< ���� ����
    txt_UNI.Text = ""   '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< ���� ����
    txt_MAJOR.Text = ""   '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< ���� ����
    txtTel.Text = ""
    txtCel.Text = ""
    txtRegDate.Text = ""

    cboPass1.ListIndex = 0
    cboPass2.ListIndex = 0
    cboPass3.ListIndex = 0
    cboPass4.ListIndex = 0

    txtOrdNo.Text = ""
    txtPayGbn.Text = ""

    txtCancel.Text = ""

    optExmY.value = True
    optSexMale.value = True
    
    fpK_Num.value = 0
    fpE_Num.value = 0
    fpM_Num.value = 0
    fpN_Num.value = 0

    '����â.
    fpBirth_ymdS.Text = ""
    fpZip.Text = ""
    txtAddr1.Text = ""
    txtAddr2.Text = ""
    txtEmail.Text = ""
    
    
    cboMu_type.ListIndex = cboMu_type.ListCount - 1
    cboKaeyol.ListIndex = 0
    
    
    For ni = 1 To SATAM_COUNT + 1 Step 1
        chkSatam(ni).value = 0
    Next ni

    For ni = 1 To 12 Step 1
        chkEng2(ni).value = 0
    Next ni

    For ni = 1 To 8 Step 1
        chkGwatam(ni).value = 0
    Next ni

    For ni = 1 To 4 Step 1
        chkMath(ni).value = 0
        chkNonsul(ni).value = 0
    Next ni

    '>> 1���� �п�
    Call basCommonSTD.Set_CboSch(cboSel1_Sch, basModule.SchCD)
    
    '>> 2���� �п�
    cboSel2_Sch.ListIndex = 0

    
    
    
End Sub


'���� �� ��ȸ �� ���뿡 ���� ȭ�� ����
Private Sub changeEnableGwamoks(bSatam As Boolean, bEng2 As Boolean, bGwatam As Boolean, bMath As Boolean, bNonsul As Boolean)

    Dim ni      As Integer
    
    ' �̼��Ë����� +1
    For ni = 1 To 12 Step 1    '< ��Ž
        If True = bSatam Then
            chkSatam(ni).Enabled = True
        Else
            chkSatam(ni).Enabled = False
        End If
    Next ni

    For ni = 1 To 12 Step 1                 '< ��2�ܱ���
        If True = bEng2 Then
            chkEng2(ni).Enabled = True
        Else
            chkEng2(ni).Enabled = False
        End If
    Next ni

    For ni = 1 To 9 Step 1                  '< ��Ž
        If True = bGwatam Then
            chkGwatam(ni).Enabled = True
        Else
            chkGwatam(ni).Enabled = False
        End If
    Next ni

    For ni = 1 To 4 Step 1                  '< ����
        If True = bMath Then
            chkMath(ni).Enabled = True
        Else
            chkMath(ni).Enabled = False
        End If
    Next ni

    For ni = 1 To 4 Step 1                  '< ���
        If True = bNonsul Then
            chkNonsul(ni).Enabled = True
        Else
            chkNonsul(ni).Enabled = False
        End If
    Next ni
    
    
End Sub


Private Sub cboKaeyol_Click()

    If Me.Tag = "LOAD" Then Exit Sub
    
    Select Case Trim(basModule.SchCD)
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
                Case "01", "03", "05", "11", "18", "21"                                                   '< 2008.02.15 : �迭 - ����, ����, ����      2009.06.02 : �迭�߰�
                    Call changeEnableGwamoks(True, True, False, False, True)
                    
                Case "02", "04", "06", "08", "12", "19", "22"                                             '< 2008.02.15 : �迭 - ����, ����, ����      2009.06.02 : �迭�߰�
                    Call changeEnableGwamoks(False, False, True, True, True)
            End Select
        Case Else
            Select Case Trim(Right(cboKaeyol.Text, 30))
                Case "01", "03", "04", "06", "11", "16", "19", "21"                        '< 2008.01.10 : �迭 - ����
                    Call changeEnableGwamoks(True, True, False, False, True)
                    
                Case "02", "05", "07", "12", "17", "20", "22"                               '< 2008.01.10 : �迭 - ����
                    Call changeEnableGwamoks(False, False, True, True, True)
            End Select
    End Select
    
End Sub


Private Sub cmdGwamokView_Click()
    fraGwamok.Left = 60
    fraGwamok.Top = 3390
    fraGwamok.ZOrder 0
    fraGwamok.Visible = True
End Sub

Private Sub cmdClose_Click()

    fraGwamok.Visible = False
    
End Sub


'>> ����� �ʱ�ȭ
Private Sub init_RegForm()

End Sub





'>> �л�����ϱ�
Private Sub cmdStdin_Click()
    Dim bRet        As Boolean
    
    '>> üũ����
    If Trim(fpExmID.UnFmtText) = "" Then
        MsgBox "�����ȣ�� �����ϴ�.", vbExclamation + vbOKOnly, "�л����"
        Exit Sub
    End If
    If Trim(fpBirth_ymd.UnFmtText) = "" Then
        MsgBox "��������� �����ϴ�.", vbExclamation + vbOKOnly, "�л����"
        Exit Sub
    End If
    
    
    On Error GoTo ErrStmt
    
    cmdStdin.Enabled = False
        bRet = Save_Stdin           '<< �л����
            
    cmdStdin.Enabled = True
    
    If bRet = True Then
        MsgBox "�л� ����Ͽ����ϴ�.", vbInformation + vbOKOnly, "�л����"
        Call cmdNew_Click
        
    Else
        
    End If
    
    Exit Sub
ErrStmt:
    MsgBox "�л���Ͻ� ������ �߻��Ͽ����ϴ�." & vbCrLf & _
        Trim(CStr(Err.Number)) & ":" & Trim(Err.Description), vbCritical + vbOKOnly, "�л����"
    On Error GoTo 0
    
End Sub

'>> �л����
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
    
    '>> ���� �Ķ���Ͱ� ���� ������ �޸𸮿��� ������.
    For ni = 0 To DBCmd.Parameters.count - 1 Step 1
        DBCmd.Parameters.Delete (0)
    Next ni
    
    '>> �л� ���/����
        sTmp = "INSERT":    If Trim(txtSchNo.Text) > "" Then sTmp = "UPDATE"
        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
            Set DBParam = DBCmd.CreateParameter("V_STYPE", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
            
    '>> �ý����ڵ�
        sTmp = Trim(txtSchNo.Text)
        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
            Set DBParam = DBCmd.CreateParameter("V_SCHNO", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
    '>> �п��ڵ�
        sTmp = basModule.SchCD
        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
            Set DBParam = DBCmd.CreateParameter("V_ACID", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
    '>> �����ȣ
        sTmp = Trim(fpExmID.UnFmtText)
        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
            Set DBParam = DBCmd.CreateParameter("V_EXMID", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
    '>> �л���
        sTmp = Trim(txtStdNM.Text)
        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
            Set DBParam = DBCmd.CreateParameter("V_STDNM", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
    '>> �������
        sTmp = Trim(fpBirth_ymd.UnFmtText)
        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
            Set DBParam = DBCmd.CreateParameter("V_birth_ymd", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
    
            
    '>> ��/������ ����
        If optExmY.value = True Then
            sTmp = "1"
        ElseIf optExmN.value = True Then
            sTmp = "0"
        End If
        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
            Set DBParam = DBCmd.CreateParameter("V_EXMTYPE", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
    
    '>> �迭
        sTmp = Trim(Right(cboKaeyol.Text, 30))
        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
            Set DBParam = DBCmd.CreateParameter("V_KAEYOL", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
    
    
    '## ���ð��� ###
        '>> ��Ž���� ����
        sTmp = ""
        For ni = 1 To SATAM_COUNT Step 1
            If chkSatam(ni).value = 1 Then
                sTmp = sTmp & Format(SATAM_CLASS + ni, "00") & "|"
            End If
        Next ni
        
        If basModule.SchCD = "J" Then
            If chkSatam(11).value = 1 Then: sTmp = sTmp & TGANG_CODE & "|"  'Ư�� ���縸
        End If
        
        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
            Set DBParam = DBCmd.CreateParameter("V_SEL1", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam

        '>> ��2�ܱ��� ����
        sTmp = ""
        For ni = 1 To 12 Step 1         '< 2008.01.14 : ���� �߰�����
            If chkEng2(ni).value = 1 Then
                sTmp = sTmp & Format(30 + ni, "00") & "|"
            End If
        Next ni
        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
            Set DBParam = DBCmd.CreateParameter("V_SEL2", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam

        '>> ��Ž���� ����
        sTmp = ""
        For ni = 1 To 8 Step 1
            If chkGwatam(ni).value = 1 Then
                sTmp = sTmp & Format(50 + ni, "00") & "|"
            End If
        Next ni
        
        If basModule.SchCD = "J" Then
            If chkGwatam(9).value = 1 Then: sTmp = sTmp & TGANG_CODE & "|"   'Ư�� ���縸
        End If
        
        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
            Set DBParam = DBCmd.CreateParameter("V_SEL3", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam

        '>> �������� ����
        sTmp = ""
        For ni = 1 To 4 Step 1
            If chkMath(ni).value = 1 Then
                sTmp = sTmp & Format(80 + ni, "00") & "|"
            End If
        Next ni
        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
            Set DBParam = DBCmd.CreateParameter("V_SEL4", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam

        '>> ������� ����
        sTmp = ""
        For ni = 1 To 4 Step 1
            If chkNonsul(ni).value = 1 Then
                sTmp = sTmp & Format(90 + ni, "00") & "|"
            End If
        Next ni
        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
            Set DBParam = DBCmd.CreateParameter("V_SEL5", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam


        '>> Ŭ���г뷮���������Ƿ� ����ó��
        Set DBParam = DBCmd.CreateParameter("V_SEL7", adVarChar, adParamInput, nLength, ""):   DBCmd.Parameters.Append DBParam
            
            
    '>> ��������
        nTmp = CDbl(fpK_Num.value)
            Set DBParam = DBCmd.CreateParameter("V_K_NUM", adDouble, adParamInput, , nTmp):   DBCmd.Parameters.Append DBParam
    '>> ��������
        nTmp = CDbl(fpE_Num.value)
            Set DBParam = DBCmd.CreateParameter("V_E_NUM", adDouble, adParamInput, , nTmp):   DBCmd.Parameters.Append DBParam
    '>> ��������
        nTmp = CDbl(fpM_Num.value)
            Set DBParam = DBCmd.CreateParameter("V_M_NUM", adDouble, adParamInput, , nTmp):   DBCmd.Parameters.Append DBParam
    '>> �հ�
        nTmp = CDbl(fpK_Num.value) + CDbl(fpM_Num.value) + CDbl(fpE_Num.value)
            Set DBParam = DBCmd.CreateParameter("V_TOT_NUM", adDouble, adParamInput, , nTmp):   DBCmd.Parameters.Append DBParam
    '>> ��������
        nTmp = CDbl(fpN_Num.value)
            Set DBParam = DBCmd.CreateParameter("V_N_NUM", adDouble, adParamInput, , nTmp):   DBCmd.Parameters.Append DBParam

    '>> 1���� �п�
        sTmp = Trim(Right(cboSel1_Sch.Text, 30))
        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
            Set DBParam = DBCmd.CreateParameter("V_SEL1_SCH", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
    '>> 2���� �п�
        sTmp = ""
        If Trim(Right(cboSel2_Sch.Text, 30)) <> "X" Then
            sTmp = Trim(Right(cboSel2_Sch.Text, 30))
        End If
        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
            Set DBParam = DBCmd.CreateParameter("V_SEL2_SCH", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam


    '>> 1���� �հ��п�
        sTmp = ""
        If Trim(Right(cboPass1.Text, 30)) <> "X" Then
            sTmp = Trim(Right(cboPass1.Text, 30))
        End If
        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
            Set DBParam = DBCmd.CreateParameter("V_PASS1", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
    '>> 2���� �հ��п�
        sTmp = ""
        If Trim(Right(cboPass2.Text, 30)) <> "X" Then
            sTmp = Trim(Right(cboPass2.Text, 30))
        End If
        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
            Set DBParam = DBCmd.CreateParameter("V_PASS2", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
    '>> 3���� �հ��п�
        sTmp = ""
        If Trim(Right(cboPass3.Text, 30)) <> "X" Then
            sTmp = Trim(Right(cboPass3.Text, 30))
        End If
        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
            Set DBParam = DBCmd.CreateParameter("V_PASS3", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
    '>> 4���� �հ��п�
        sTmp = ""
        If Trim(Right(cboPass4.Text, 30)) <> "X" Then
            sTmp = Trim(Right(cboPass4.Text, 30))
        End If
        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
            Set DBParam = DBCmd.CreateParameter("V_PASS4", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
    
    '>> ��ȭ��ȣ
        sTmp = Trim(txtTel.Text)
        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
            Set DBParam = DBCmd.CreateParameter("V_TEL", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
    '>> �ڵ���
        sTmp = Trim(txtCel.Text)
        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
            Set DBParam = DBCmd.CreateParameter("V_CEL", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
    '>> �������� ����
        sTmp = ""
        If Trim(Right(cboPTS_Sel.Text, 30)) <> "X" Then sTmp = Trim(Right(cboPTS_Sel.Text, 30))
        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
            Set DBParam = DBCmd.CreateParameter("V_PTS", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
    '>> ���
        sTmp = ""
        If Trim(Right(cboMu_type.Text, 30)) <> "X" Then sTmp = Trim(Right(cboMu_type.Text, 30))
        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
            Set DBParam = DBCmd.CreateParameter("V_PTS", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
        
    '>> �кθ�HP
        sTmp = Trim(txt_P_Phone.Text)
        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
            Set DBParam = DBCmd.CreateParameter("V_PRNT_TEL", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
            
    '>> ��������
        sTmp = Trim(txt_UNI.Text)
        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
            Set DBParam = DBCmd.CreateParameter("V_D_UNIVCD", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
            
    '>> �����ܴ�
        sTmp = Trim(txt_MAJOR.Text)
        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
            Set DBParam = DBCmd.CreateParameter("V_D_MAJORCD", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
            
    '>> Ž���� �뷮��,���ĸ� �����Ƿ� ����ó��
        'Ž��
        sTmp = ""
        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
            Set DBParam = DBCmd.CreateParameter("V_SEL6", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
            
        'Ž������
        nTmp = 0
            Set DBParam = DBCmd.CreateParameter("V_T_NUM", adDouble, adParamInput, , nTmp):   DBCmd.Parameters.Append DBParam
            
    '>> ���� ����
        If optSexMale.value = True Then
            sTmp = SexMaleValue
        ElseIf optSexFemale.value = True Then
            sTmp = SexFemaleValue
        End If
        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
            Set DBParam = DBCmd.CreateParameter("V_SEX", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
    
    
    
    '>> ������ ���
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
    
    MsgBox "�л� ��Ͻ� ������ �߻��Ͽ����ϴ�." & vbCrLf & _
        Trim(CStr(Err.Number)) & ":" & Trim(Err.Description), vbCritical + vbOKOnly, "�л����"
        
    Save_Stdin = False
    
    
    
End Function


'>> �л������ϱ�
Private Sub cmdStdDel_Click()
    Dim bRet        As Boolean
    Dim sTmp        As String
    
    '>> üũ����
    If Trim(txtSchNo.Text) = "" Then
        MsgBox "�ý����ڵ尡 �����ϴ�.", vbExclamation + vbOKOnly, "�л�����"
        Exit Sub
    End If
    
    sTmp = Trim(txtStdNM.Text) & "�� �л��� �����Ͻðڽ��ϱ�?"
    If MsgBox(sTmp, vbQuestion + vbYesNo, "�л�����") = vbNo Then
        Exit Sub
    End If
    
    On Error GoTo ErrStmt
    
    cmdStdDel.Enabled = False
        bRet = Delete_StdOut
        
    cmdStdDel.Enabled = True
    
    If bRet = True Then
        MsgBox "�л� �����Ͽ����ϴ�.", vbInformation + vbOKOnly, "�л�����"
    Else
        
    End If
    
    Exit Sub
ErrStmt:
    MsgBox "�л������� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "�л�����"
    On Error GoTo 0
    
End Sub

'>> �л�����
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
    
    '>> ���� �Ķ���Ͱ� ���� ������ �޸𸮿��� ������.
    For ni = 0 To DBCmd.Parameters.count - 1 Step 1
        DBCmd.Parameters.Delete (0)
    Next ni
            
    '>> �ý����ڵ�
        sTmp = Trim(txtSchNo.Text)
        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
            Set DBParam = DBCmd.CreateParameter("V_SCHNO", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
    '>> �п��ڵ�
        sTmp = basModule.SchCD
        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
            Set DBParam = DBCmd.CreateParameter("V_ACID", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
    
    '>> ������ ���
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
    
    MsgBox "�л� ������ ������ �߻��Ͽ����ϴ�." & vbCrLf & _
        Trim(CStr(Err.Number)) & ":" & Trim(Err.Description), vbCritical + vbOKOnly, "�л�����"
        
    Delete_StdOut = False
End Function


'>> �л� �հ���ҹ�ư ����
Private Sub cmdCancel_Click()
    
    Dim bRet        As Boolean
    
    '>> üũ����
    If Trim(txtSchNo.Text) = "" Then
        MsgBox "�ý��� �ڵ尡 �����ϴ�.", vbExclamation + vbOKOnly, "�л� ����ϱ�"
        Exit Sub
    End If
    
    On Error GoTo ErrStmt
    
    cmdCancel.Enabled = False
        bRet = Cancel_StdOut        ' ��� �Լ�
        
    cmdCancel.Enabled = True
    
    If bRet = True Then
        MsgBox "�л� �հ���� �Ͽ����ϴ�.", vbInformation + vbOKOnly, "�л� ����ϱ�"
    Else
        DMsgBox "�л� �հ���ҽ� ������ �߻��Ͽ����ϴ�.", "�л� ����ϱ�"
                
    End If
    
    Exit Sub
ErrStmt:
    MsgBox "�л� �հ���ҽ� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "�л� ����ϱ�"
    On Error GoTo 0
    
End Sub

'>> �л� �հ�����ϱ�
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
        Select Case Trim(basModule.SchCD)
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
            sStr = sStr & "         '" & Trim(basModule.SchCD) & "',"
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
                sStr = sStr & "    SET ACID  = '" & Trim(basModule.SchCD) & "',"
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
















'>> �л� ��ȸ�ϱ�
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
    
    '2011-01-11 ���ѿ� �뷮�� ��û�� ���� ���� ���� �� �κп� �������� �� ���� �ܴ� �Է�
    
    sStr = ""
    sStr = sStr & "  SELECT SCHNO, EXMID, STDNM, SEL1_SCH , SEL2_SCH, "
    
'    If basModule.SchCD = "N" Then
'        Select Case basModule.RegID
'            Case "10000", "00002", "10003", "00001" '�迵������
'                sStr = sStr & "         SUBSTR(REPLACE(birth_ymd,'-',''),1,6)||'-'||SUBSTR(REPLACE(birth_ymd,'-',''),7,7) AS birth_ymd,"
'            Case "10001"                            '������
'                sStr = sStr & "         SUBSTR(REPLACE(birth_ymd,'-',''),1,6)||'-*******' AS birth_ymd,"
'            Case "10002"                            '������
'                sStr = sStr & "         SUBSTR(REPLACE(birth_ymd,'-',''),1,6)||'-*******' AS birth_ymd,"
'        End Select
'    Else
'        sStr = sStr & "         SUBSTR(REPLACE(birth_ymd,'-',''),1,6)||'-'||SUBSTR(REPLACE(birth_ymd,'-',''),7,7) AS birth_ymd,"
'    End If
    
        '20121009 ���� : �ֹι�ȣ���� -> ������Ϸ� �����ϸ鼭 �ٷ��� �ּ� �ڵ�� �������ٷ� ��ü��.
        sStr = sStr & " birth_ymd, "
    
'<< �迭 >> : 2008.01.09

    '�迭 decode sql�� ����
    sStr = sStr & basCommonSTD.Get_SqlKaeyolDecode()

    
    
    sStr = sStr & "     /* ��Ž, ��Ž �и� */"
    sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'" & constSatamCodes(0) & "|') > 0 THEN          /* ��Ž-�ѱ��� */"
    sStr = sStr & "             '" & constSatamCodes(0) & "'"
    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'51|') > 0 THEN     /* ��Ž-����1 */"
    sStr = sStr & "             '51'"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             '00'"
    sStr = sStr & "         END END SEL1,"
    sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'" & constSatamCodes(1) & "|') > 0 THEN          /* ��Ž-����� */"
    sStr = sStr & "             '" & constSatamCodes(1) & "'"
    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'52|') > 0 THEN     /* ��Ž-ȭ��1 */"
    sStr = sStr & "             '52'"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             '00'"
    sStr = sStr & "         END END SEL2,"
    sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'" & constSatamCodes(2) & "|') > 0 THEN          /* ��Ž-���ƽþƻ� */"
    sStr = sStr & "             '" & constSatamCodes(2) & "'"
    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'53|') > 0 THEN     /* ��Ž-�������1 */"
    sStr = sStr & "             '53'"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             '00'"
    sStr = sStr & "         END END SEL3,"
    sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'" & constSatamCodes(3) & "|') > 0 THEN          /* ��Ž-�ѱ����� */"
    sStr = sStr & "             '" & constSatamCodes(3) & "'"
    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'54|') > 0 THEN     /* ��Ž-��������1 */"
    sStr = sStr & "             '54'"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             '00'"
    sStr = sStr & "         END END SEL4,"
    sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'" & constSatamCodes(4) & "|') > 0 THEN          /* ��Ž-�������� */"
    sStr = sStr & "             '" & constSatamCodes(4) & "'"
    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'55|') > 0 THEN     /* ��Ž-����2 */"
    sStr = sStr & "             '55'"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             '00'"
    sStr = sStr & "         END END SEL5,"
    sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'" & constSatamCodes(5) & "|') > 0 THEN          /* ��Ž-��Ȱ������ */"
    sStr = sStr & "             '" & constSatamCodes(5) & "'"
    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'56|') > 0 THEN     /* ��Ž-ȭ��2 */"
    sStr = sStr & "             '56'"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             '00'"
    sStr = sStr & "         END END SEL6,"
    sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'" & constSatamCodes(6) & "|') > 0 THEN          /* ��Ž-�����ͻ�� */"
    sStr = sStr & "             '" & constSatamCodes(6) & "'"
    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'57|') > 0 THEN     /* ��Ž-�������2 */"
    sStr = sStr & "             '57'"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             '00'"
    sStr = sStr & "         END END SEL7,"
    sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'" & constSatamCodes(7) & "|') > 0 THEN          /* ��Ž-������ġ */"
    sStr = sStr & "             '" & constSatamCodes(7) & "'"
    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'58|') > 0 THEN     /* ��Ž-��������2 */"
    sStr = sStr & "             '58'"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             '00'"
    sStr = sStr & "         END END SEL8,"
    sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'" & constSatamCodes(8) & "|') > 0 THEN          /* ��Ž-���� */"
    sStr = sStr & "             '" & constSatamCodes(8) & "'"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             '00'"
    sStr = sStr & "         END SEL9,"
    sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'" & constSatamCodes(9) & "|') > 0 THEN          /* ��Ž-��ȸ��ȭ */"
    sStr = sStr & "             '" & constSatamCodes(9) & "'"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             '00'"
    sStr = sStr & "         END SEL10,"
    sStr = sStr & "  "
    sStr = sStr & "      /* ��2�ܱ��� & ���� */"
    sStr = sStr & "              CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'31|') > 0 THEN '31'"
    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'32|') > 0 THEN '32'"
    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'33|') > 0 THEN '33'"
    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'34|') > 0 THEN '34'"
    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'35|') > 0 THEN '35'"
    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'36|') > 0 THEN '36'"
    '<< ���� >> : 2008.01.09
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
    sStr = sStr & "      /* ��� */"
    sStr = sStr & "         CASE WHEN INSTR(SEL5,'91|') > 0 THEN         /* ��� */"
    sStr = sStr & "             '91'"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             '00'"
    sStr = sStr & "         END SEL_N1,"
    sStr = sStr & "         CASE WHEN INSTR(SEL5,'92|') > 0 THEN         /* ���� */"
    sStr = sStr & "             '92'"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             '00'"
    sStr = sStr & "         END SEL_N2,"
    sStr = sStr & "         CASE WHEN INSTR(SEL5,'93|') > 0 THEN         /* ��Ž */"
    sStr = sStr & "             '93'"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             '00'"
    sStr = sStr & "         END SEL_N3,"
    sStr = sStr & "         CASE WHEN INSTR(SEL5,'94|') > 0 THEN         /* ��Ž */"
    sStr = sStr & "             '94'"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             '00'"
    sStr = sStr & "         END SEL_N4, "
    sStr = sStr & "         PAYOK, PAYNOT, "
    sStr = sStr & "         GET_INTERNET_TOT_STD_INWON('" & Trim(basModule.SchCD) & "') AS PAYTOT, "        '< ��ü���� �ϴ� �Լ�
    sStr = sStr & "         K_NUM, M_NUM, E_NUM, TOT_NUM, N_NUM, "
    sStr = sStr & "         ZIP, ADDR1, ADDR2, TEL, CEL, "
    sStr = sStr & "         REGDATE, PAYGBN, CASH_BILL_NUM, D_UNIVCD, D_MAJORCD, PRNT_TEL "
    
    Select Case Trim(Right(cboPassCN, 30))
        Case "ALL"      ' /* �հݻ��� ��� ��ȸ */
            sStr = sStr & " FROM (SELECT SCHNO, MAX(EXMID) AS EXMID, MAX(STDNM) AS STDNM, MAX(SEL1_SCH) AS SEL1_SCH, MAX(SEL2_SCH) AS SEL2_SCH, MAX(D_UNIVCD) AS D_UNIVCD, MAX(D_MAJORCD) AS D_MAJORCD, MAX(birth_ymd) AS birth_ymd,"
            sStr = sStr & "              MAX(KAEYOL) AS KAEYOL,"
            sStr = sStr & "              MAX(SEL1) AS SEL1, MAX(SEL2) AS SEL2, MAX(SEL3) AS SEL3, MAX(SEL4) AS SEL4, MAX(SEL5) SEL5, "
            sStr = sStr & "              MAX(CL_CLOSE) AS CL_CLOSE, "
            sStr = sStr & "              MAX(PAYOK) AS PAYOK, MAX(PAYNOT) AS PAYNOT, "
            sStr = sStr & "              MAX(K_NUM) AS K_NUM, MAX(M_NUM) AS M_NUM, MAX(E_NUM) AS E_NUM, MAX(TOT_NUM) AS TOT_NUM, MAX(N_NUM) AS N_NUM,"
            sStr = sStr & "              MAX(ZIP) AS ZIP, MAX(ADDR1) AS ADDR1, MAX(ADDR2) AS ADDR2, MAX(TEL) AS TEL, MAX(CEL) AS CEL, "
            sStr = sStr & "              MAX(REGDATE) AS REGDATE, MAX(PAYGBN) AS PAYGBN, MAX(CASH_BILL_NUM) AS CASH_BILL_NUM, MAX(PRNT_TEL) AS PRNT_TEL "
            sStr = sStr & "         FROM ("
            '==========================================================================================================
            
            sStr = sStr & "               SELECT SCHNO, EXMID, STDNM, SEL1_SCH, SEL2_SCH, D_UNIVCD, D_MAJORCD, birth_ymd,"
            sStr = sStr & "                      KAEYOL,"
            sStr = sStr & "                      SEL1 , SEL2, SEL3, SEL4, SEL5, CL_CLOSE, "
            sStr = sStr & "                      PAYOK, PAYNOT, "
            sStr = sStr & "                      NVL(K_NUM,0) AS K_NUM, NVL(M_NUM,0) AS M_NUM, NVL(E_NUM,0) AS E_NUM,"
            'sStr = sStr & "                      (NVL(K_NUM,0)+NVL(M_NUM,0)+NVL(E_NUM,0)) AS TOT_NUM ,"
            sStr = sStr & "                      TOT_NUM , NVL(N_NUM,0) AS N_NUM,"
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
            
            sStr = sStr & "                                WHERE ACID = '" & Trim(basModule.SchCD) & "'"
            '>> ��/������ üũ
            If Trim(Right(cboExmType.Text, 30)) = "0" Then
                sStr = sStr & "                              AND EXMTYPE = '0'"
            ElseIf Trim(Right(cboExmType.Text, 30)) = "1" Then
                sStr = sStr & "                              AND EXMTYPE = '1'"
            End If
            
            '>> ���ͳ�/�п�
            If Trim(Right(cboinGbn.Text, 30)) = "INT" Then          '< ���ͳ� ����
                sStr = sStr & "                              AND R_WAY = '2'"
            ElseIf Trim(Right(cboinGbn.Text, 30)) = "HAK" Then      '< �п� ����
                sStr = sStr & "                              AND R_WAY IN ('1','3') "
            End If
            
            '<< ���翩�� >>
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
            
            If Trim(Right(cboKaeyol_F.Text, 30)) <> "ALL" Then      ' �ι�
                sStr = sStr & "                              AND KAEYOL = '" & Trim(Right(cboKaeyol_F.Text, 30)) & "'"
            End If
    
            If Trim(txtStdNM_F.Text) <> "" Then
                sStr = sStr & "                              AND STDNM LIKE '%" & Trim(txtStdNM_F.Text) & "%'"
            End If
            If Trim(fpBirth_ymd_F.UnFmtText) <> "" Then
                sStr = sStr & "                              AND birth_ymd LIKE '" & Trim(fpBirth_ymd_F.UnFmtText) & "%'"
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
            sStr = sStr & "                  AND A.ACID = '" & Trim(basModule.SchCD) & "'"
            
            '>> ��/������ üũ
            If Trim(Right(cboExmType.Text, 30)) = "0" Then
                sStr = sStr & "              AND EXMTYPE = '0'"
            ElseIf Trim(Right(cboExmType.Text, 30)) = "1" Then
                sStr = sStr & "              AND EXMTYPE = '1'"
            End If
            
            '>> ���ͳ�/�п�
            If Trim(Right(cboinGbn.Text, 30)) = "INT" Then          '< ���ͳ� ����
                sStr = sStr & "              AND R_WAY = '2'"
            ElseIf Trim(Right(cboinGbn.Text, 30)) = "HAK" Then      '< �п� ����
                sStr = sStr & "              AND R_WAY IN ('1','3') "
            End If
            '<< ���翩�� >>
            Select Case Trim(Right(cboPay.Text, 30))
                Case "OK"
                    sStr = sStr & "          AND EXMID > ' ' "
                Case "NOT"
                    sStr = sStr & "          AND EXMID IS NULL "
            End Select
            sStr = sStr & "                  AND CL_CLOSE IS NULL "
            
            sStr = sStr & "                  AND BIGO2 IS NULL"                     '< 2008.12. ���ɺ� �л��� �⵵�� ���� �ƴϸ� NULL
                        
            sStr = sStr & "               Union All"
            sStr = sStr & "               SELECT SCHNO, EXMID, STDNM, SEL1_SCH, SEL2_SCH, D_UNIVCD, D_MAJORCD, birth_ymd,"
            sStr = sStr & "                      KAEYOL,"
            sStr = sStr & "                      SEL1 , SEL2, SEL3, SEL4, SEL5, CL_CLOSE, "
            sStr = sStr & "                      0 AS PAYOK, 0 AS PAYNOT, "
            sStr = sStr & "                      0 AS K_NUM, 0 AS M_NUM, 0 AS E_NUM, 0 AS TOT_NUM, 0 AS N_NUM,"
            sStr = sStr & "                      SUBSTR(ZIP,1,3)||'-'||SUBSTR(ZIP,4,3) AS ZIP, ADDR1, ADDR2, TEL, CEL, "
            sStr = sStr & "                      TO_CHAR(REGDATE,'YYYY-MM-DD HH24:MI') AS REGDATE, GET_PAYGUBN(ORD_NO) AS PAYGBN, CASH_BILL_NUM, PRNT_TEL AS PRNT_TEL "
            sStr = sStr & "                 From CLSTD01TB"
            sStr = sStr & "                WHERE (PASS1 = '" & Trim(basModule.SchCD) & "'" & " OR"
            sStr = sStr & "                       PASS2 = '" & Trim(basModule.SchCD) & "'" & " OR"
            sStr = sStr & "                       PASS3 = '" & Trim(basModule.SchCD) & "'" & " OR"
            sStr = sStr & "                       PASS4 = '" & Trim(basModule.SchCD) & "'" & " )"
            
            '>> ��/������ üũ
            If Trim(Right(cboExmType.Text, 30)) = "0" Then
                sStr = sStr & "              AND EXMTYPE = '0'"
            ElseIf Trim(Right(cboExmType.Text, 30)) = "1" Then
                sStr = sStr & "              AND EXMTYPE = '1'"
            End If
            
            '>> ���ͳ�/�п�
            If Trim(Right(cboinGbn.Text, 30)) = "INT" Then          '< ���ͳ� ����
                sStr = sStr & "              AND R_WAY = '2'"
            ElseIf Trim(Right(cboinGbn.Text, 30)) = "HAK" Then      '< �п� ����
                sStr = sStr & "              AND R_WAY IN ('1','3') "
            End If
            
            '<< ���翩�� >>
            Select Case Trim(Right(cboPay.Text, 30))
                Case "OK"
                    sStr = sStr & "          AND EXMID > ' ' "
                Case "NOT"
                    sStr = sStr & "          AND EXMID IS NULL "
            End Select
            sStr = sStr & "                  AND CL_CLOSE IS NULL "
            
            sStr = sStr & "                  AND BIGO2 IS NULL"                     '< 2008.12. ���ɺ� �л��� �⵵�� ���� �ƴϸ� NULL
            
            '==========================================================================================================
            sStr = sStr & "               )"
            sStr = sStr & "        GROUP BY SCHNO"
            sStr = sStr & "       )"
            
            
        Case Else       ' /* Ư�� �հ������� �հ��ڸ� ��ȸ�� */
            sStr = sStr & " FROM (SELECT SCHNO, EXMID, STDNM, SEL1_SCH, SEL2_SCH, D_UNIVCD, D_MAJORCD, birth_ymd,"
            sStr = sStr & "              KAEYOL,"
            sStr = sStr & "              SEL1 , SEL2, SEL3, SEL4, SEL5, CL_CLOSE, "
            sStr = sStr & "              0 AS PAYOK , 0 AS PAYNOT, "
            sStr = sStr & "              GET_INTERNET_TOT_STD_INWON('" & Trim(basModule.SchCD) & "') AS PAYTOT"     '< ��ü���� �ϴ� �Լ�
            sStr = sStr & "         From CLSTD01TB"
            sStr = sStr & "        WHERE PASS" & Trim(Right(cboPassCN, 30)) & " = '" & Trim(basModule.SchCD) & "'"
            
            '>> ��/������ üũ
            If Trim(Right(cboExmType.Text, 30)) = "0" Then
                sStr = sStr & "      AND EXMTYPE = '0'"
            ElseIf Trim(Right(cboExmType.Text, 30)) = "1" Then
                sStr = sStr & "      AND EXMTYPE = '1'"
            End If
            
            '>> ���ͳ�/�п�
            If Trim(Right(cboinGbn.Text, 30)) = "INT" Then          '< ���ͳ� ����
                sStr = sStr & "      AND R_WAY = '2'"
            ElseIf Trim(Right(cboinGbn.Text, 30)) = "HAK" Then      '< �п� ����
                sStr = sStr & "      AND R_WAY IN ('1','3') "
            End If
            
            '<< ���翩�� >>
            Select Case Trim(Right(cboPay.Text, 30))
                Case "OK"
                    sStr = sStr & "  AND EXMID > ' ' "
                Case "NOT"
                    sStr = sStr & "  AND EXMID IS NULL "
            End Select
            sStr = sStr & "          AND CL_CLOSE IS NULL "
            
            sStr = sStr & "          AND BIGO2 IS NULL"                     '< 2008.12. ���ɺ� �л��� �⵵�� ���� �ƴϸ� NULL
            
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
    
    '<< ���翩�� >>
    Select Case Trim(Right(cboPay.Text, 30))
        Case "OK"
            sStr = sStr & " AND EXMID > ' ' "
        Case "NOT"
            sStr = sStr & " AND EXMID IS NULL "
    End Select
    
    If Trim(Right(cboKaeyol_F.Text, 30)) <> "ALL" Then      ' �ι�
        sStr = sStr & " AND KAEYOL = '" & Trim(Right(cboKaeyol_F.Text, 30)) & "'"
    End If
    
    If Trim(txtStdNM_F.Text) <> "" Then
        sStr = sStr & " AND STDNM LIKE '%" & Trim(txtStdNM_F.Text) & "%'"
    End If
    If Trim(fpBirth_ymd_F.UnFmtText) <> "" Then
        sStr = sStr & " AND birth_ymd LIKE '" & Trim(fpBirth_ymd_F.UnFmtText) & "%'"
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


    
    '>> �п�
'        sTmp = Trim(basModule.SchCD)
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
        
    '>> �����ȣ
'        If Trim(fpExmID_F.UnFmtText) > "" Then
'            sTmp = Trim(fpExmID_F.UnFmtText)
'            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'                Set DBParam = DBCmd.CreateParameter("EXMID", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
'        End If
        
'    '>> �л���
'        If Trim(txtStdNM_F.Text) > "" Then
'            sTmp = "%" & Trim(txtStdNM_F.Text) & "%"
'            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'                Set DBParam = DBCmd.CreateParameter("STDNM", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
'        End If
'
'    '>> �ֹι�ȣ
'        If Trim(fpbirth_ymd_F.UnFmtText) > "" Then
'            sTmp = "%" & Trim(fpbirth_ymd_F.UnFmtText) & "%"
'            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'                Set DBParam = DBCmd.CreateParameter("birth_ymd", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
'        End If
'
'    '>> �����п�
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
    
    
    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic���·� ���ԵǸ� record count�� �� �� ����.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop
    '2011-01-11 ���ѿ� ����� �� �� �� ���� double ó��(�뷮������)
    With DBRec
        If .RecordCount > 0 Then
            .MoveFirst
            
            For nRec = 1 To .RecordCount Step 1
            
                If nRec = 1 Then        '< �ο����� ���� �κ��� �ѹ��� ����ϸ� �˴ϴ�.
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
                
                
                'MsgBox " ":   If IsNull(.Fields("birth_ymd")) = False Then sTmp = Trim(.Fields("birth_ymd"))
                sprSTD_F.Col = 6
                    sTmp = " ":  If IsNull(.Fields("birth_ymd")) = False Then sTmp = Trim(.Fields("birth_ymd"))
                    Call basFunction.Set_SprType_Text(sprSTD_F, "CENTER", "LEFT", LenB(sTmp), sTmp)
                
                sprSTD_F.SetCellBorder sprSTD_F.Col, sprSTD_F.Row, sprSTD_F.Col, sprSTD_F.Row, 2, basModule.SectionColor2, CellBorderStyleSolid
                
                
                sprSTD_F.Col = 7
                    nTmp = 0:   If IsNumeric(.Fields("K_NUM")) = True Then nTmp = Trim(.Fields("K_NUM"))
                    Call basFunction.Set_SprType_Numeric(sprSTD_F, 1, 1, 999999, "", nTmp)
                sprSTD_F.Col = 8
                    nTmp = 0:   If IsNumeric(.Fields("M_NUM")) = True Then nTmp = Trim(.Fields("M_NUM"))
                    Call basFunction.Set_SprType_Numeric(sprSTD_F, 1, 1, 999999, "", nTmp)
                sprSTD_F.Col = 9
                    nTmp = 0:   If IsNumeric(.Fields("E_NUM")) = True Then nTmp = Trim(.Fields("E_NUM"))
                    Call basFunction.Set_SprType_Numeric(sprSTD_F, 1, 1, 999999, "", nTmp)
                sprSTD_F.Col = 10
                    nTmp = 0:   If IsNumeric(.Fields("TOT_NUM")) = True Then nTmp = Trim(.Fields("TOT_NUM"))
                    Call basFunction.Set_SprType_Numeric(sprSTD_F, 1, 1, 999999, "", nTmp)
                
                sprSTD_F.Col = sprSTD_F.Col + 1
                    sTmp = " ":  If IsNull(.Fields("N_NUM")) = False Then sTmp = Trim(.Fields("N_NUM")) ' ���ŵ��
                    Call basFunction.Set_SprType_Text(sprSTD_F, "CENTER", "CENTER", LenB(sTmp), sTmp)
                    
                sprSTD_F.Col = sprSTD_F.Col + 1
                    sTmp = " ":  If IsNull(.Fields("GAEYUL")) = False Then sTmp = Trim(.Fields("GAEYUL"))
                    Call basFunction.Set_SprType_Text(sprSTD_F, "CENTER", "LEFT", LenB(sTmp), sTmp)
                
                
                sprSTD_F.SetCellBorder sprSTD_F.Col, sprSTD_F.Row, sprSTD_F.Col, sprSTD_F.Row, 2, basModule.SectionColor2, CellBorderStyleSolid
                
                '>> ���ð��� (��Ž/ ��Ž)
                For ni = 1 To 10 Step 1

                    '�Ķ��� ���� ��Լ� �߱�
                    If ni Mod 4 = 1 Then: sprSTD_F.SetCellBorder sprSTD_F.Col, sprSTD_F.Row, sprSTD_F.Col, sprSTD_F.Row, 2, basModule.SectionColor2, CellBorderStyleSolid

                    sprSTD_F.Col = sprSTD_F.Col + 1
                    
                    sGbn = "SEL" & Trim(CStr(ni))
                    sTmp = IIf(Trim(.Fields(sGbn)) = "00", "", Trim(.Fields(sGbn)))
                    If sTmp <> "" Then: sTmp = basGwamok.Get_StrGwaMokByCode(sTmp)   ' sTmp(�ڵ�)�� ���� �����̸�������

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
                        
                            Case "31":  sTmp = "����"
                            Case "32":  sTmp = "�Ͼ�"
                            Case "33":  sTmp = "�����ĳľ�"
                            Case "34":  sTmp = "�Ҿ�"
                            Case "35":  sTmp = "�߱���"
                            Case "36":  sTmp = "�ѹ�"
                            
                            '<< ���� >> : 2008.01.09
                            Case "37":  sTmp = "���"
                            Case "38":  sTmp = "����"
                            Case "39":  sTmp = "����"
                            Case "40":  sTmp = "�����"
                            Case "41":  sTmp = "��������"
                            
                            Case "81":  sTmp = "������"
                            Case "82":  sTmp = "�̻����"
                            Case "83":  sTmp = "Ȯ�����"
                            Case "84":  sTmp = "��������"
                            
                        End Select
                        Call basFunction.Set_SprType_Text(sprSTD_F, "CENTER", "LEFT", LenB(sTmp), sTmp)
                    End If
                End If
                
                sprSTD_F.SetCellBorder sprSTD_F.Col, sprSTD_F.Row, sprSTD_F.Col, sprSTD_F.Row, 2, basModule.SectionColor1, CellBorderStyleSolid
                
            '>> ���
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
                                    Case "91":  sTmp = "���"
                                    Case "92":  sTmp = "����"
                                    Case "93":  sTmp = "�ܱ���"         '< ����
                                    Case "94":  sTmp = ""               '< ����
                                    
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
    
    MsgBox "�л��� ��ȸ�Ǿ����ϴ�.", vbInformation + vbOKOnly, "�л���ȸ"
    
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    cmdFind.Enabled = True
    
    Exit Sub
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    cmdFind.Enabled = True
    
    MsgBox "�л���ȸ�� ������ �߻��Ͽ����ϴ�." & vbCrLf & _
           Trim(CStr(Err.Number)) & ":" & Err.Description, vbCritical + vbOKOnly, "�л���ȸ"
           
    On Error GoTo 0
End Sub

'>> �л�����
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

'>> �����л� �����ֱ�
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
    
'    If basModule.SchCD = "N" Then
'        Select Case basModule.RegID
'            Case "10000", "00002", "10003", "00001" '�迵������
'                sStr = sStr & "         SUBSTR(REPLACE(birth_ymd,'-',''),1,6)||SUBSTR(REPLACE(birth_ymd,'-',''),7,7) AS birth_ymd,"
'            Case "10001"                            '������
'                sStr = sStr & "         SUBSTR(REPLACE(birth_ymd,'-',''),1,6)||'*******' AS birth_ymd,"
'            Case "10002"                            '������
'                sStr = sStr & "         SUBSTR(REPLACE(birth_ymd,'-',''),1,6)||'*******' AS birth_ymd,"
'        End Select
'    Else
'        sStr = sStr & "         SUBSTR(REPLACE(birth_ymd,'-',''),1,6)||SUBSTR(REPLACE(birth_ymd,'-',''),7,7) AS birth_ymd,"
'    End If
        
    '20121009 ���� : �ֹι�ȣ���� -> ������Ϸ� �����ϸ鼭 �ٷ��� �ּ� �ڵ�� �������ٷ� ��ü��.
    sStr = sStr & " birth_ymd, "
    
    sStr = sStr & "         EXMTYPE , KAEYOL  ,"
    sStr = sStr & "         SEL1    , SEL2    , SEL3    , SEL4    , SEL5    ,"
    sStr = sStr & "         K_NUM   , M_NUM   , E_NUM   , TOT_NUM , N_NUM   ,"
    sStr = sStr & "         SEL1_SCH, SEL2_SCH,"
    sStr = sStr & "         PASS1 , PASS2, PASS3, PASS4 , TEL     , PRNT_TEL , CEL     , ORD_NO , "
    sStr = sStr & "         TO_CHAR(REGDATE,'YYYY-MM-DD HH24:MI') AS REGDATE, "
    sStr = sStr & "         GET_CANCELYN(SCHNO) AS WORKDAY, "
    sStr = sStr & "         GET_PAYGUBN(ORD_NO) AS PAYGBN, "
    sStr = sStr & "         ZIP, ADDR1, ADDR2 , EMAIL   , PTS_SEL , MU_TYPE, SEX "
    sStr = sStr & "    From CLSTD01TB "
    sStr = sStr & "   WHERE SCHNO = '" & Trim(aSchNO) & "'"
    
    
    'W,Q�� ������ ���ο� �迭 (���ε� �迭�� �߰�)
    Select Case Trim(basModule.SchCD)
        Case "W", "Q"
            sStr = sStr & "     AND (ACID  = '" & Trim(basModule.SchCD) & "'"
            sStr = sStr & "          OR ACID = 'K')"
        Case Else
            sStr = sStr & "     AND ACID  = '" & Trim(basModule.SchCD) & "'"
    End Select
    
    Set DBCmd = New ADODB.Command
    Set DBRec = New ADODB.Recordset
    Set DBParam = New ADODB.Parameter
    

    Text1.Text = sStr
    
    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
    DBCmd.CommandText = sStr
    DBCmd.CommandType = adCmdText
    DBCmd.CommandTimeout = 30


    
'    '>> �л��ڵ�
'        sTmp = Trim(aSchNO)
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
'    '>> �п�
'        sTmp = Trim(basModule.SchCD)
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
        
        
    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic���·� ���ԵǸ� record count�� �� �� ����.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop
    
    With DBRec
        If .RecordCount <> 1 Then
            MsgBox "��ȸ�� �л��� �����ϴ�.", vbExclamation + vbOKOnly, "�л���ȸ"
        Else
            .MoveFirst
            
            txtSchNo.Text = "":     If IsNull(.Fields("SCHNO")) = False Then txtSchNo.Text = Trim(.Fields("SCHNO"))
            fpExmID.Text = "":      If IsNull(.Fields("EXMID")) = False Then fpExmID.Text = Trim(.Fields("EXMID"))
            txtStdNM.Text = "":     If IsNull(.Fields("STDNM")) = False Then txtStdNM.Text = Trim(.Fields("STDNM"))
            
            txt_UNI.Text = "":     If IsNull(.Fields("D_UNIVCD")) = False Then txt_UNI.Text = Trim(.Fields("D_UNIVCD"))
            txt_MAJOR.Text = "":     If IsNull(.Fields("D_MAJORCD")) = False Then txt_MAJOR.Text = Trim(.Fields("D_MAJORCD"))
            
            fpBirth_ymd.Text = ""
            fpBirth_ymdS.Text = ""
                If IsNull(.Fields("birth_ymd")) = False Then
                    fpBirth_ymd.Text = Trim(.Fields("birth_ymd"))
                    fpBirth_ymdS.Text = Trim(.Fields("birth_ymd"))
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
            
            '>> ���� ����
            If IsNull(.Fields("SEX")) = False Then
                If Trim(.Fields("SEX")) = "M" Then
                    optSexMale.value = True
                Else
                    optSexFemale.value = True
                End If
            End If
            
            '��/�� ����
            optExmY.value = True
            If IsNull(.Fields("EXMTYPE")) = False Then
                If Trim(.Fields("EXMTYPE")) = "0" Then
                    optExmN.value = True
                Else
                    optExmY.value = True
                End If
            End If
            
            '>> �迭
            cboKaeyol.ListIndex = 0
            Call basCommonSTD.Set_CboKaeyol(cboKaeyol, basModule.SchCD, Trim(.Fields("KAEYOL")))
            
            
            
            '>> 1�����п�
            If IsNull(.Fields("SEL1_SCH")) = False Then
                Call basCommonSTD.Set_CboSch(cboSel1_Sch, Trim(.Fields("SEL1_SCH")))
            End If
            
            
            '>> 2�����п�
            If IsNull(.Fields("SEL2_SCH")) = False Then
                Call basCommonSTD.Set_CboSch(cboSel2_Sch, Trim(.Fields("SEL2_SCH")))
            End If
            
           
            '>> 1�� �հ��п�
            If IsNull(.Fields("PASS1")) = False Then
                Call basCommonSTD.Set_CboSch(cboPass1, Trim(.Fields("PASS1")))
            End If
            
            '>> 2�� �հ��п�
            If IsNull(.Fields("PASS2")) = False Then
                Call basCommonSTD.Set_CboSch(cboPass2, Trim(.Fields("PASS2")))
            End If
            
            '>> 3�� �հ��п�
            If IsNull(.Fields("PASS3")) = False Then
                Call basCommonSTD.Set_CboSch(cboPass3, Trim(.Fields("PASS3")))
            End If
                
            
            '>> 4�� �հ��п�
            If IsNull(.Fields("PASS4")) = False Then
                Call basCommonSTD.Set_CboSch(cboPass4, Trim(.Fields("PASS4")))
            End If
            
            
            
            '2011-01-11 ���ѿ� ����� �� �� �� ���� double ó��(�뷮������)
            fpK_Num.value = 0:  If IsNull(.Fields("K_NUM")) = False Then fpK_Num.value = CDbl(.Fields("K_NUM"))
            fpE_Num.value = 0:  If IsNull(.Fields("E_NUM")) = False Then fpE_Num.value = CDbl(.Fields("E_NUM"))
            fpM_Num.value = 0:  If IsNull(.Fields("M_NUM")) = False Then fpM_Num.value = CDbl(.Fields("M_NUM"))
            fpN_Num.value = 0:  If IsNull(.Fields("N_NUM")) = False Then fpN_Num.value = CDbl(.Fields("N_NUM"))
            
            
        '## ���ð���
            '>> ��Ž
            '2011-05-17 ���ѿ� ��ȸ �κ� �Ѱ��� �� �߰��� ���� ���� �߰�����
            If (Trim(basModule.SchCD) = "Q") Then
                For ni = 1 To 11 + 1 Step 1
                    chkSatam(ni).value = 0
                Next ni
            Else
                For ni = 1 To 11 Step 1
                    chkSatam(ni).value = 0
                Next ni
            End If
            
            If IsNull(.Fields("SEL1")) = False Then
                sTmp = Trim(.Fields("SEL1"))
                sDiv = Split(sTmp, "|", -1, vbTextCompare)
                
                '2011-05-17 ���ѿ�  ��ȸ �κ� �Ѱ��� �� �߰��� ���� ���� �߰�����
                Dim arrIdx As Integer
                For ni = 0 To UBound(sDiv) - 1 Step 1
                    
                    '�뷮�� ��û
                    If sDiv(ni) = TGANG_CODE Then
                        chkSatam(11).value = 1  '���狚���� 95�� Ư��.
                    Else
                        '���� ��Ž�� �ڵ� ������ 21~30����. arrIdx = CInt(21) - (21-1)   , arrIdx = 1
                        arrIdx = CInt(sDiv(ni)) - SATAM_CLASS
                        If arrIdx > 0 And arrIdx <= chkSatam.UBound Then
                            chkSatam(arrIdx).value = 1
                        Else
                            MsgBox "DB�� ��Ž�����ڵ� ���� �ùٸ��� �ʽ��ϴ�. �ٽ� �������ּ���."
                        End If
                    End If
                        
                Next ni
            End If
            
            '>> ��2�ܱ���
            For ni = 1 To 12 Step 1
                chkEng2(ni).value = 0
            Next ni
            If IsNull(.Fields("SEL2")) = False Then
                sTmp = Trim(.Fields("SEL2"))
                sDiv = Split(sTmp, "|", -1, vbTextCompare)
                
                For ni = 0 To UBound(sDiv) - 1 Step 1
                    chkEng2(CInt(sDiv(ni)) - 30).value = 1
                Next ni
            End If
            
            '>> ��Ž
            For ni = 1 To 9 Step 1
                chkGwatam(ni).value = 0
            Next ni
            If IsNull(.Fields("SEL3")) = False Then
                sTmp = Trim(.Fields("SEL3"))
                sDiv = Split(sTmp, "|", -1, vbTextCompare)
                
                For ni = 0 To UBound(sDiv) - 1 Step 1
                    If sDiv(ni) = TGANG_CODE Then
                        chkGwatam(9).value = 1  '���狚���� 95�� Ư��.
                    Else
                        chkGwatam(CInt(sDiv(ni)) - 50).value = 1
                    End If
                Next ni
            End If
            '>> ����
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
            '>> ���
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
            
            
            '������/�� : ����/ ���ĸ��̸�
            Select Case Trim(basModule.SchCD)
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
            
            
            
            '���ɵ��
            If IsNull(.Fields("MU_TYPE")) = True Then
                cboMu_type.ListIndex = 0
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
    MsgBox "�����л� ��ȸ�� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "�л���ȸ"
    
End Sub
















'## EXCEL �ڷ���ȸ
Private Sub cmdGetExcel_Click()
    
    On Error GoTo ErrStmt
    
    cmdGetExcel.Enabled = False
        Call Get_Excel_Data
        
    cmdGetExcel.Enabled = True
    
    Exit Sub
ErrStmt:
    MsgBox "�����ڷ� �������� �� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "�л� �����ڷ� ��������"
    On Error GoTo 0
7
End Sub

'##############################################################################
'################ 2012.11�� ����� ���� �ʴ°� ���Ƽ� ������ enable = false�� �ص�.
'################            ������. ���߿� �ʿ��Ұ�� ���� ���� ���� �۾��ؾ���.
'##############################################################################

Private Sub Get_Excel_Data()

'    Dim sPath       As String
'
'    ' Excel Data ó��
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
'   On Error GoTo ErrStmt1
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
'            MsgBox "������ ������ �����ϴ�.", vbExclamation + vbOKOnly, Me.Caption
'            Exit Sub
'        End If
'
'        sPath = .fileName
'
'    End With
'
'    On Error GoTo 0
'
'    On Error GoTo ErrStmt2                          '>> error ó��
'
'    Set xlsDBConn = New ADODB.Connection
'    sConn = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
'            "Data Source=" & sPath & ";" & _
'            "Extended Properties=""Excel 8.0;HDR=no;"";"
'
'    With xlsDBConn
'        .ConnectionString = sConn                   ' �����ͺ��̽��� ������ �õ��մϴ�.
'        .ConnectionTimeout = 30                     ' ���� �ð����� ������ ���� ������ �ڵ����� �����ϴ�.
'        .Properties("Prompt") = adPromptNever       ' �̰��� ADO���� �⺻ ������Ʈ ����Դϴ�.
'        .CursorLocation = adUseClient               ' Ŀ����ġ�� Client �ʿ� �ֽ��ϴ�.
'
'        .Open                                       ' �����ͺ��̽��� ���ϴ�.
'
'        Do While .State And adStateConnecting
'            DoEvents
'        Loop
'    End With
'
''>> ���� DB Open
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
'        MsgBox "Excel Data�� �����ϴ�.", vbExclamation + vbOKOnly, "IT2008"
'        Exit Sub
'    End If
'
'
'    sprExcel_STD_Data.MaxRows = 0       ' �ʱ�ȭ
'
'
'    DBExRec.MoveFirst
'
'    '## header 1 line skip
'    DBExRec.MoveNext
'
'
'    For nRow = 2 To DBExRec.RecordCount Step 1
'    '�п��ڵ�
'        sTmp = "":  If IsNull(DBExRec.Fields(0)) = False Then sTmp = UCase(Trim(DBExRec.Fields(0)))
'        uExcel_StdData.ACID = sTmp
'    '�����ȣ
'        sTmp = "":  If IsNull(DBExRec.Fields(1)) = False Then sTmp = Trim(DBExRec.Fields(1))
'        uExcel_StdData.EXMID = sTmp
'    '�л���
'        sTmp = "":  If IsNull(DBExRec.Fields(2)) = False Then sTmp = Trim(DBExRec.Fields(2))
'        uExcel_StdData.STDNM = sTmp
'    '�������
'        sTmp = "":  If IsNull(DBExRec.Fields(3)) = False Then sTmp = Trim(DBExRec.Fields(3))
'        sTmp = Replace(sTmp, "-", "", 1, -1, vbTextCompare)
'        If basFunction.LenKor(sTmp) > 6 Then
'            sTmp = Left(sTmp, 4) & "-" & Mid(sTmp, 5, 2) & "-" & Mid(sTmp, 7, 2)
'        End If
'        uExcel_StdData.Birth_ymd = sTmp
'    '��.������
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
'    '�迭
'        sTmp = "01"
'        If Trim(basModule.schcd) = "N" Then             '< �迭 : 2008.01.09 - �뷮��
'            If IsNull(DBExRec.Fields(5)) = False Then
'                sTmp = UCase(Trim(DBExRec.Fields(5)))
'                Select Case sTmp
'                    Case "1" To "9"
'                        sTmp = Format(sTmp, "00")
'                    Case "�ι�", "��"
'                        sTmp = "01"
'                    Case "�ڿ�", "��"
'                        sTmp = "02"
'                    Case "��ü", "��"
'                        sTmp = "03"
'
'                    Case "����(��)", "������"
'                        sTmp = "04"
'                    Case "�ι�����", "�����ι�"
'                        sTmp = "05"
'                    Case "�ڿ�����", "�����ڿ�"
'                        sTmp = "06"
'
'                    Case "�ż��ι�"
'                        sTmp = "07"
'                    Case "�ż��ڿ�"
'                        sTmp = "08"
'                    'Case "�ż������ι�"
'                    '    sTmp = "09"
'                    'Case "�ż������ڿ�"
'                    '    sTmp = "10"
'
'                    Case "�����ι�", "����"
'                        sTmp = "11"
'                    Case "�����ڿ�", "����"
'                        sTmp = "12"
'                    Case "��ü", "��"
'                        sTmp = "13"
'
'                    Case "�����(��)", "�������"
'                        sTmp = "14"
'                    Case "���ι�����", "������ι�"
'                        sTmp = "15"
'                    Case "���ڿ�����", "������ڿ�"
'                        sTmp = "16"
'                    Case "������ι�"
'                        sTmp = "21"
'                    Case "������ڿ�"
'                        sTmp = "22"
'                    Case Else
'                        sTmp = "01"
'                End Select
'            End If
'        ElseIf Trim(basModule.schcd) = "K" Or Trim(basModule.schcd) = "W" Or Trim(basModule.schcd) = "Q" Then           '< �迭 : 2008.01.09 - �뷮��, 2008.03.24 - ����
'            If IsNull(DBExRec.Fields(5)) = False Then
'                sTmp = UCase(Trim(DBExRec.Fields(5)))
'                Select Case sTmp
'                    Case "1" To "9"
'                        sTmp = Format(sTmp, "00")
'                    Case "�ι�", "��"
'                        sTmp = "01"
'                    Case "�ڿ�", "��"
'                        sTmp = "02"
'
'                    Case "�ְ�����", "�ֹ�"
'                        sTmp = "04"
'                    Case "�ְ��Ǵ�", "����"
'                        sTmp = "05"
'
'                    Case "�߰�����", "�߹�"
'                        sTmp = "06"
'                    Case "�߰��Ǵ�", "����"
'                        sTmp = "07"
'
'                    Case "�������ι�"
'                        sTmp = "11"
'                    Case "�������ڿ�"
'                        sTmp = "12"
'
'                    Case "�������ι�16"
'                        sTmp = "16"
'                    Case "�������ڿ�17"
'                        sTmp = "17"
'
'                    Case Else
'                        sTmp = "01"
'                End Select
'            End If
'        ElseIf Trim(basModule.schcd) = "S" Then             '< �迭 : 2008.02.15 - ����
'            If IsNull(DBExRec.Fields(5)) = False Then
'                sTmp = UCase(Trim(DBExRec.Fields(5)))
'                Select Case sTmp
'                    Case "1" To "9"
'                        sTmp = Format(sTmp, "00")
'                    Case "�ι�", "��"
'                        sTmp = "01"
'                    Case "�ڿ�", "��"
'                        sTmp = "02"
'
'                    Case "Ư��", "Ư���ι�"
'                        sTmp = "03"
'                    Case "Ư��", "Ư���ڿ�"
'                        sTmp = "04"
'
'                    Case "�ż������ι�"
'                        sTmp = "05"
'                    Case "�ż������ڿ�"
'                        sTmp = "06"
'
'                    Case "�ż��ι�"
'                        sTmp = "11"
'                    Case "�ż��ڿ�"
'                        sTmp = "12"
'
'                    Case "�ż���������"
'                        sTmp = "08"
'
'                    Case "�ι������̾�"
'                        sTmp = "18"
'                    Case "�ڿ������̾�"
'                        sTmp = "19"
'
'                    Case Else
'                        sTmp = "01"
'                End Select
'            End If
'        ElseIf Trim(basModule.schcd) = "P" Then             '< �迭 : 2008.02.15 - ����
'            If IsNull(DBExRec.Fields(5)) = False Then
'                sTmp = UCase(Trim(DBExRec.Fields(5)))
'                Select Case sTmp
'                    Case "1" To "9"
'                        sTmp = Format(sTmp, "00")
'                    Case "�ι�", "��"
'                        sTmp = "01"
'                    Case "�ڿ�", "��"
'                        sTmp = "02"
'
'                    Case "Ư��", "Ư���ι�"
'                        sTmp = "03"
'                    Case "Ư��", "Ư���ڿ�"
'                        sTmp = "04"
'
'                    Case Else
'                        sTmp = "01"
'                End Select
'            End If
'
'        ElseIf Trim(basModule.schcd) = "J" Then             '< ����
'            If IsNull(DBExRec.Fields(5)) = False Then
'                sTmp = UCase(Trim(DBExRec.Fields(5)))
'                Select Case sTmp
'                    Case "1" To "9"
'                        sTmp = Format(sTmp, "00")
'                    Case "�ι�", "��"
'                        sTmp = "01"
'                    Case "�ڿ�", "��"
'                        sTmp = "02"
'
'                    Case "�ż��ι�"
'                        sTmp = "11"
'                    Case "�ż��ڿ�"
'                        sTmp = "12"
'
'                    Case "�ι������̾�"
'                        sTmp = "18"
'                    Case "�ڿ������̾�"
'                        sTmp = "19"
'
'                    Case Else
'                        sTmp = "01"
'                End Select
'            End If
'        ElseIf Trim(basModule.schcd) = "B" Then             '< ����
'            If IsNull(DBExRec.Fields(5)) = False Then
'                sTmp = UCase(Trim(DBExRec.Fields(5)))
'                Select Case sTmp
'                    Case "1" To "9"
'                        sTmp = Format(sTmp, "00")
'                    Case "�ι�", "��"
'                        sTmp = "01"
'                    Case "�ڿ�", "��"
'                        sTmp = "02"
'                    Case "��ü", "��"
'                        sTmp = "03"
'                    Case "�ι�PS��"
'                        sTmp = "23"
'                    Case "�ڿ�PM��"
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
'                    Case "�ι�", "��"
'                        sTmp = "01"
'                    Case "�ڿ�", "��"
'                        sTmp = "02"
'                    Case "��ü", "��"
'                        sTmp = "03"
'                    Case Else
'                        sTmp = "01"
'                End Select
'            End If
'        End If
'        uExcel_StdData.kaeyol = sTmp
'
'    '1 �����п�
'        sTmp = Trim(basModule.schcd)
'        If IsNull(DBExRec.Fields(6)) = False Then
'            sTmp = UCase(Trim(DBExRec.Fields(6)))
'            Select Case sTmp
'                Case "N", "K", "S", "P", "M", "W", "Q", "J", "B"
'                    ' NEXT
'                Case "�뷮��"
'                    sTmp = "N"
'                Case "����"
'                    sTmp = "K"
'                Case "����"
'                    sTmp = "S"
'                Case "����M", "���ĸ��̸�", "���� MIMAC", "����MIMAC", "����"
'                    sTmp = "P"
'                Case "����M", "�������̸�", "���� MIMAC", "����MIMAC", "����"
'                    sTmp = "M"
'
'                Case "�ָ����Ǵ�", "�ָ���", "�ֹ�"
'                    sTmp = "W"
'                Case "�߰����Ǵ�", "�߰���", "�߹�"
'                    sTmp = "Q"
'
'                Case "����"
'                    sTmp = "J"
'
'                Case "�λ�"
'                    sTmp = "B"
'
'                Case Else
'                    sTmp = Trim(basModule.schcd)
'            End Select
'        End If
'        uExcel_StdData.WANT_ACID1 = sTmp
'
'    '2 �����п�
'        sTmp = Trim(basModule.schcd)
'        If IsNull(DBExRec.Fields(7)) = False Then
'            sTmp = UCase(Trim(DBExRec.Fields(7)))
'            Select Case sTmp
'                Case "N", "K", "S", "P", "M", "W", "Q", "J", "B"
'                    ' NEXT
'                Case "�뷮��"
'                    sTmp = "N"
'                Case "����"
'                    sTmp = "K"
'                Case "����"
'                    sTmp = "S"
'                Case "����M", "���ĸ��̸�", "���� MIMAC", "����MIMAC", "����"
'                    sTmp = "P"
'                Case "����M", "�������̸�", "���� MIMAC", "����MIMAC", "����"
'                    sTmp = "M"
'
'                Case "�ָ����Ǵ�", "�ָ���", "�ֹ�"
'                    sTmp = "W"
'                Case "�߰����Ǵ�", "�߰���", "�߹�"
'                    sTmp = "Q"
'
'                Case "����"
'                    sTmp = "J"
'
'                Case "�λ�"
'                    sTmp = "B"
'
'                Case Else
'                    sTmp = Trim(basModule.schcd)
'            End Select
'        End If
'        uExcel_StdData.WANT_ACID2 = sTmp
'
'    '����
'        nTmp = 0:  If IsNumeric(DBExRec.Fields(8)) = True Then nTmp = CLng(Trim(DBExRec.Fields(8)))
'        uExcel_StdData.KOR = nTmp
'    '����
'        nTmp = 0:  If IsNumeric(DBExRec.Fields(9)) = True Then nTmp = CLng(Trim(DBExRec.Fields(9)))
'        uExcel_StdData.ENG = nTmp
'    '����
'        nTmp = 0:  If IsNumeric(DBExRec.Fields(10)) = True Then nTmp = CLng(Trim(DBExRec.Fields(10)))
'        uExcel_StdData.MAT = nTmp
'
'    '��Ž
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
'        'uExcel_StdData.SATAM11 = ""
'
'        For ni = 1 To SATAM_COUNT Step 1
'            sTmp = ""
'            nC = 10 + ni
'            If IsNull(DBExRec.Fields(nC)) = False Then sTmp = Trim(DBExRec.Fields(nC))
'
'            Dim gwaMoks() As String
'            Dim gwaMokCodes() As String
'            gwaMoks = basGwamok.Get_GwaMokNames("SEL1")
'            gwaMokCodes = basGwamok.Get_GwaMokCodes("SEL1")
'
'            Select Case sTmp
'                Case ""
'                    'no action
'                Case gwaMoks(0)
'                    uExcel_StdData.SATAM1 = gwaMokCodes(0) + "|"
'                Case gwaMoks(1)
'                    uExcel_StdData.SATAM2 = gwaMokCodes(1) + "|"
'                Case gwaMoks(2)
'                    uExcel_StdData.SATAM3 = gwaMokCodes(2) + "|"
'                Case gwaMoks(3)
'                    uExcel_StdData.SATAM4 = gwaMokCodes(3) + "|"
'                Case gwaMoks(4)
'                    uExcel_StdData.SATAM5 = gwaMokCodes(4) + "|"
'                Case gwaMoks(5)
'                    uExcel_StdData.SATAM6 = gwaMokCodes(5) + "|"
'                Case gwaMoks(6)
'                    uExcel_StdData.SATAM7 = gwaMokCodes(6) + "|"
'                Case gwaMoks(7)
'                    uExcel_StdData.SATAM8 = gwaMokCodes(7) + "|"
'                Case gwaMoks(8)
'                    uExcel_StdData.SATAM9 = gwaMokCodes(8) + "|"
'                Case gwaMoks(9)
'                    uExcel_StdData.SATAM10 = gwaMokCodes(9) + "|"
'            End Select
'        Next ni
'
'    '��2�ܱ���
'        uExcel_StdData.ENG2 = ""
'
'        sTmp = ""
'            nC = 10 + 11 + 1
'            If IsNull(DBExRec.Fields(nC)) = False Then sTmp = Trim(DBExRec.Fields(nC))
'
'            Select Case sTmp
'                Case ""
'                    'no action
'                Case "����"
'                    uExcel_StdData.ENG2 = "31|"
'                Case "�Ͼ�"
'                    uExcel_StdData.ENG2 = "32|"
'                Case "����", "�����ĳ�"
'                    uExcel_StdData.ENG2 = "33|"
'                Case "�Ҿ�"
'                    uExcel_StdData.ENG2 = "34|"
'                Case "�߱�", "�߾�"
'                    uExcel_StdData.ENG2 = "35|"
'                Case "�ѹ�"
'                    uExcel_StdData.ENG2 = "36|"
'
'                '<< ���� >> : 2008.01.09
'                Case "���"
'                    uExcel_StdData.ENG2 = "37|"
'                Case "����"
'                    uExcel_StdData.ENG2 = "38|"
'                Case "����"
'                    uExcel_StdData.ENG2 = "39|"
'                Case "�����"
'                    uExcel_StdData.ENG2 = "40|"
'                Case "��������"
'                    uExcel_StdData.ENG2 = "41|"
'                Case "�ƶ���"
'                    uExcel_StdData.ENG2 = "42|"
'
'            End Select
'    '��Ž
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
'                Case "��1"
'                    uExcel_StdData.GWATAM1 = "51|"
'                Case "ȭ1"
'                    uExcel_StdData.GWATAM2 = "52|"
'                Case "��1"
'                    uExcel_StdData.GWATAM3 = "53|"
'                Case "��1"
'                    uExcel_StdData.GWATAM4 = "54|"
'                Case "��2"
'                    uExcel_StdData.GWATAM5 = "55|"
'                Case "ȭ2"
'                    uExcel_StdData.GWATAM6 = "56|"
'                Case "��2"
'                    uExcel_StdData.GWATAM7 = "57|"
'                Case "��2"
'                    uExcel_StdData.GWATAM8 = "58|"
'            End Select
'        Next ni
'    '����
'        uExcel_StdData.SURI = ""
'
'        sTmp = ""
'            nC = 10 + 11 + 1
'            If IsNull(DBExRec.Fields(nC)) = False Then sTmp = Trim(DBExRec.Fields(nC))
'
'            Select Case sTmp
'                Case ""
'                    'no action
'                Case "����"
'                    uExcel_StdData.SURI = "81|"
'                Case "�̻�"
'                    uExcel_StdData.SURI = "82|"
'                Case "Ȯ��"
'                    uExcel_StdData.SURI = "83|"
'                Case "����"
'                    uExcel_StdData.SURI = "84|"
'            End Select
'    '���
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
'                Case "���"
'                    uExcel_StdData.NONSUL1 = "91|"
'                Case "����"
'                    uExcel_StdData.NONSUL2 = "92|"
'                Case "�ܱ���"
'                    uExcel_StdData.NONSUL3 = "93|"
'                Case ""
'                    uExcel_StdData.NONSUL4 = "94|"
'            End Select
'        Next ni
'
'
'    '## �������忡 ������ �ֱ� --------------------------------------------------------------------
'        With sprExcel_STD_Data
'            .MaxRows = .MaxRows + 1
'            .Row = .MaxRows:            .RowHeight(.Row) = 13
'
'            '>> �п�
'                .Col = 1
'                    sTmp = uExcel_StdData.ACID
'                    Call basFunction.Set_SprType_Text(sprExcel_STD_Data, "center", "left", basFunction.LenKor(sTmp), sTmp)
'
'            '>> �����ȣ
'                .Col = .Col + 1
'                    sTmp = uExcel_StdData.EXMID
'                    Call basFunction.Set_SprType_Text(sprExcel_STD_Data, "center", "left", basFunction.LenKor(sTmp), sTmp)
'            '>> �л���
'                .Col = .Col + 1
'                    sTmp = uExcel_StdData.STDNM
'                    Call basFunction.Set_SprType_Text(sprExcel_STD_Data, "center", "left", basFunction.LenKor(sTmp), sTmp)
'            '>> �������
'                .Col = .Col + 1
'                    sTmp = Replace(uExcel_StdData.Birth_ymd, "-", "", 1, -1, vbTextCompare)
'                    Call basFunction.Set_SprType_Text(sprExcel_STD_Data, "center", "left", basFunction.LenKor(sTmp), sTmp)
'            '>> ��.������
'                .Col = .Col + 1
'                    sTmp = uExcel_StdData.EXMTYPE
'                    Call basFunction.Set_SprType_Text(sprExcel_STD_Data, "center", "left", basFunction.LenKor(sTmp), sTmp)
'            '>> �迭
'                .Col = .Col + 1
'                    sTmp = uExcel_StdData.kaeyol
'                    Call basFunction.Set_SprType_Text(sprExcel_STD_Data, "center", "left", basFunction.LenKor(sTmp), sTmp)
'
'            '>> 1 �����п�
'                .Col = .Col + 1
'                    sTmp = uExcel_StdData.WANT_ACID1
'                    Call basFunction.Set_SprType_Text(sprExcel_STD_Data, "center", "left", basFunction.LenKor(sTmp), sTmp)
'            '>> 2 �����п�
'                .Col = .Col + 1
'                    sTmp = uExcel_StdData.WANT_ACID2
'                    Call basFunction.Set_SprType_Text(sprExcel_STD_Data, "center", "left", basFunction.LenKor(sTmp), sTmp)
'
'            '>> ����
'                .Col = .Col + 1
'                    nTmp = uExcel_StdData.KOR
'                    Call basFunction.Set_SprType_Numeric(sprExcel_STD_Data, 0, 0, 9999, "", nTmp)
'            '>> ����
'                .Col = .Col + 1
'                    nTmp = uExcel_StdData.ENG
'                    Call basFunction.Set_SprType_Numeric(sprExcel_STD_Data, 0, 0, 9999, "", nTmp)
'            '>> ����
'                .Col = .Col + 1
'                    nTmp = uExcel_StdData.MAT
'                    Call basFunction.Set_SprType_Numeric(sprExcel_STD_Data, 0, 0, 9999, "", nTmp)
'
'            '>> ��Ž
'                .Col = .Col + 1
'                    sTmp = ""
'
''                    Dim n_i As Integer
''                    Dim subject As String
''                    For n_i = 1 To SATAM_COUNT
''                        nC = 10 + n_i
''                        subject = Trim(DBExRec.Fields(nC))
''
''                        If IsNull(subject) = False Then
''                            sTmp = basGwamok.Get_GwaMokCodeByName(subject) & "|"  '�ڵ� ���Ѵ�. 21|24
''                        End If
''
''                    Next n_i
'
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
'
'
'                    sTmp = Replace(sTmp, " ", "", 1, -1, vbTextCompare)
'                    Call basFunction.Set_SprType_Text(sprExcel_STD_Data, "center", "left", basFunction.LenKor(sTmp), sTmp)
'
'            '>> ��2�ܱ���
'                .Col = .Col + 1
'                    sTmp = Trim(uExcel_StdData.ENG2)
'                    Call basFunction.Set_SprType_Text(sprExcel_STD_Data, "center", "left", basFunction.LenKor(sTmp), sTmp)
'
'            '>> ��Ž
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
'            '>> ����
'                .Col = .Col + 1
'                    sTmp = Trim(uExcel_StdData.SURI)
'                    Call basFunction.Set_SprType_Text(sprExcel_STD_Data, "center", "left", basFunction.LenKor(sTmp), sTmp)
'
'            '>> ���
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
'    MsgBox "�л� �����ڷḦ ������ �Խ��ϴ�.", vbInformation + vbOKOnly, Me.Caption
'
'    On Error GoTo 0
'    Exit Sub
'ErrStmt1:
'    MsgBox "���� ���ϼ����� �Ͻʽÿ�.", vbExclamation + vbOKOnly, Me.Caption
'    Exit Sub
'ErrStmt2:
'    Set DBExRec = Nothing
'    Set DBExCmd = Nothing
'    xlsDBConn.Close
'    Set xlsDBConn = Nothing
'
'    MsgBox "EXCEL �ڷ� Open�� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, Me.Caption
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





'>> �л����
Private Sub cmdExcelSave_Click()
    Dim bRet        As Boolean
    
    '>> üũ����
    If sprExcel_STD_Data.MaxRows = 0 Then
        MsgBox "����� �л��� �����ϴ�.", vbExclamation + vbOKOnly, "������ �л����"
        Exit Sub
    End If
    
    On Error GoTo ErrStmt
    
    cmdExcelSave.Enabled = False
        bRet = Save_Excel_Stdin             '<< �л����
            
    cmdExcelSave.Enabled = True
            
    If bRet = True Then
        MsgBox "�л� �����ڷ�� ����Ͽ����ϴ�.", vbInformation + vbOKOnly, "������ �л����"
    Else
        MsgBox "�л� �����ڷ� ��Ͻ� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "������ �л����"
    End If
    
    Exit Sub
ErrStmt:
    MsgBox "�л� �����ڷ� ��Ͻ� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "������ �л����"
    On Error GoTo 0
    
End Sub

'>> �л����  ��������
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
    
        '>> ���� �Ķ���Ͱ� ���� ������ �޸𸮿��� ������.
        For ni = 0 To DBCmd.Parameters.count - 1 Step 1
            DBCmd.Parameters.Delete (0)
        Next ni
    
        '>> ��� �ű�
            sTmp = "INSERT"
            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
                Set DBParam = DBCmd.CreateParameter("V_STYPE", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
        '>> �ý����ڵ�
            sTmp = ""
            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
                Set DBParam = DBCmd.CreateParameter("V_SCHNO", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
        '>> �п��ڵ�
            sprExcel_STD_Data.Col = 1
                sTmp = UCase(Trim(sprExcel_STD_Data.Text))
            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
                Set DBParam = DBCmd.CreateParameter("V_ACID", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
        
        '>> �����ȣ
            sprExcel_STD_Data.Col = sprExcel_STD_Data.Col + 1
                sTmp = UCase(Trim(sprExcel_STD_Data.Text))
            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
                Set DBParam = DBCmd.CreateParameter("V_EXMID", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
        '>> �л���
            sprExcel_STD_Data.Col = sprExcel_STD_Data.Col + 1
                sTmp = UCase(Trim(sprExcel_STD_Data.Text))
            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
                Set DBParam = DBCmd.CreateParameter("V_STDNM", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
        '>> �������
            sprExcel_STD_Data.Col = sprExcel_STD_Data.Col + 1
                sTmp = UCase(Trim(sprExcel_STD_Data.Text)):     sTmp = Replace(sTmp, "-", "", 1, -1, vbTextCompare)
            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
                Set DBParam = DBCmd.CreateParameter("V_birth_ymd", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
        
        '>> ��/������ ����
            sprExcel_STD_Data.Col = sprExcel_STD_Data.Col + 1
                sTmp = UCase(Trim(sprExcel_STD_Data.Text))
            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
                Set DBParam = DBCmd.CreateParameter("V_EXMTYPE", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam

        '>> �迭
            sprExcel_STD_Data.Col = sprExcel_STD_Data.Col + 1
                sTmp = UCase(Trim(sprExcel_STD_Data.Text))
            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
                Set DBParam = DBCmd.CreateParameter("V_KAEYOL", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam


        '## ���ð��� ###
            '>> ��Ž���� ����
            sprExcel_STD_Data.Col = 12
                sTmp = UCase(Trim(sprExcel_STD_Data.Text))
            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
                Set DBParam = DBCmd.CreateParameter("V_SEL1", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam

            '>> ��2�ܱ��� ����
            sprExcel_STD_Data.Col = sprExcel_STD_Data.Col + 1
                sTmp = UCase(Trim(sprExcel_STD_Data.Text))
            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
                Set DBParam = DBCmd.CreateParameter("V_SEL2", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam

            '>> ��Ž���� ����
            sprExcel_STD_Data.Col = sprExcel_STD_Data.Col + 1
                sTmp = UCase(Trim(sprExcel_STD_Data.Text))
            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
                Set DBParam = DBCmd.CreateParameter("V_SEL3", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam

            '>> �������� ����
            sprExcel_STD_Data.Col = sprExcel_STD_Data.Col + 1
                sTmp = UCase(Trim(sprExcel_STD_Data.Text))
            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
                Set DBParam = DBCmd.CreateParameter("V_SEL4", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam

            '>> ������� ����
            sprExcel_STD_Data.Col = sprExcel_STD_Data.Col + 1
                sTmp = UCase(Trim(sprExcel_STD_Data.Text))
            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
                Set DBParam = DBCmd.CreateParameter("V_SEL5", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam


        nTotJumsu = 0

        '>> ��������
            sprExcel_STD_Data.Col = 9
                If Trim(sprExcel_STD_Data.Text) > " " Then
                    nTmp = CLng(Trim(sprExcel_STD_Data.Text))
                Else
                    nTmp = 0
                End If
                nTotJumsu = nTotJumsu + nTmp
                Set DBParam = DBCmd.CreateParameter("V_K_NUM", adDouble, adParamInput, , nTmp):   DBCmd.Parameters.Append DBParam
        '>> ��������
            sprExcel_STD_Data.Col = sprExcel_STD_Data.Col + 1
                If Trim(sprExcel_STD_Data.Text) > " " Then
                    nTmp = CLng(Trim(sprExcel_STD_Data.Text))
                Else
                    nTmp = 0
                End If
                nTotJumsu = nTotJumsu + nTmp
                Set DBParam = DBCmd.CreateParameter("V_E_NUM", adDouble, adParamInput, , nTmp):   DBCmd.Parameters.Append DBParam
        '>> ��������
            sprExcel_STD_Data.Col = sprExcel_STD_Data.Col + 1
                If Trim(sprExcel_STD_Data.Text) > " " Then
                    nTmp = CLng(Trim(sprExcel_STD_Data.Text))
                Else
                    nTmp = 0
                End If
                nTotJumsu = nTotJumsu + nTmp
                Set DBParam = DBCmd.CreateParameter("V_M_NUM", adDouble, adParamInput, , nTmp):   DBCmd.Parameters.Append DBParam
        '>> �հ�
            nTmp = nTotJumsu
                Set DBParam = DBCmd.CreateParameter("V_TOT_NUM", adDouble, adParamInput, , nTmp):   DBCmd.Parameters.Append DBParam
        '>> ���ŵ��
            sprExcel_STD_Data.Col = sprExcel_STD_Data.Col + 1
            If Trim(sprExcel_STD_Data.Text) > " " Then
                nTmp = CLng(Trim(sprExcel_STD_Data.Text))
            Else
                nTmp = 0
            End If
            
            Set DBParam = DBCmd.CreateParameter("V_N_NUM", adDouble, adParamInput, , nTmp):   DBCmd.Parameters.Append DBParam

        '>> 1���� �п�
            sprExcel_STD_Data.Col = 7
                sTmp = UCase(Trim(sprExcel_STD_Data.Text))
            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
                Set DBParam = DBCmd.CreateParameter("V_SEL1_SCH", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
        '>> 2���� �п�
            sprExcel_STD_Data.Col = sprExcel_STD_Data.Col + 1
                sTmp = UCase(Trim(sprExcel_STD_Data.Text))
            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
                Set DBParam = DBCmd.CreateParameter("V_SEL2_SCH", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam


        '>> 1���� �հ��п�
            sTmp = ""
            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
                Set DBParam = DBCmd.CreateParameter("V_PASS1", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
        '>> 2���� �հ��п�
            sTmp = ""

            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
                Set DBParam = DBCmd.CreateParameter("V_PASS2", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
        '>> 3���� �հ��п�
            sTmp = ""
            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
                Set DBParam = DBCmd.CreateParameter("V_PASS3", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
        '>> 4���� �հ��п�
            sTmp = ""
            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
                Set DBParam = DBCmd.CreateParameter("V_PASS4", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
                
        '>> ��ȭ��ȣ
            sTmp = ""
            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
                Set DBParam = DBCmd.CreateParameter("V_TEL", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
        '>> �ּ�
            sTmp = ""
            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
                Set DBParam = DBCmd.CreateParameter("V_CEL", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
            
        '>> ������ ���
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










'## ��ü�л� ������ �ޱ�
Private Sub cmdAllStdData_Click()

    Call imgExcel_Click
    
    Exit Sub
        
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
'    '> �ʱ�ȭ
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
'            MsgBox "������ ������ �����ϴ�.", vbExclamation + vbOKOnly, Me.Caption
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
'    sStr = sStr & "  SELECT SCHNO AS �ý����ڵ�   , "
'    sStr = sStr & "         ACID  AS �п�   , "
'    sStr = sStr & "         EXMID AS �����ȣ, STDNM AS �л�, D_UNIVCD AS ��������, D_MAJORCD AS �����ܴ�,"
'
''    If basModule.SchCD = "N" Then
''        Select Case basModule.RegID
''            Case "10000", "00002", "10003", "00001" '�迵������
''                sStr = sStr & "         SUBSTR(REPLACE(birth_ymd,'-',''),1,6)||'-'||SUBSTR(REPLACE(birth_ymd,'-',''),7,7) AS �ֹι�ȣ,"
''            Case "10001"                            '������
''                sStr = sStr & "         SUBSTR(REPLACE(birth_ymd,'-',''),1,6)||'-*******' AS �ֹι�ȣ,"
''            Case "10002"                            '������
''                sStr = sStr & "         SUBSTR(REPLACE(birth_ymd,'-',''),1,6)||'-*******' AS �ֹι�ȣ,"
''        End Select
''    Else
''        sStr = sStr & "         SUBSTR(REPLACE(birth_ymd,'-',''),1,6)||'-'||SUBSTR(REPLACE(birth_ymd,'-',''),7,7) AS �ֹι�ȣ,"
''    End If
'
'    '20121009 ���� : �ֹι�ȣ���� -> ������Ϸ� �����ϸ鼭 �ٷ��� �ּ� �ڵ�� �������ٷ� ��ü��.
'    sStr = sStr & " birth_ymd, "
'
'    sStr = sStr & "         DECODE(EXMTYPE,'0','������','1','������') AS ��������, "
'    sStr = sStr & "         DECODE(KAEYOL,'01','�ι�',"
'    sStr = sStr & "                       '02','�ڿ�',"
''<< �迭 >> : 2008.01.09
'    If Trim(basModule.SchCD) = "N" Then
'        sStr = sStr & "                   '03','��ü',"
'        sStr = sStr & "                   '04','����(��)',"
'        sStr = sStr & "                   '05','�ι�����',"
'        sStr = sStr & "                   '06','�ڿ�����',"
'
'        sStr = sStr & "                   '07','�ż��ι�',"
'        sStr = sStr & "                   '08','�ż��ڿ�',"
'        sStr = sStr & "                   '09','�ż������ι�',"
'        sStr = sStr & "                   '10','�ż������ڿ�',"
'
'        sStr = sStr & "                   '11','��)�ι�',"
'        sStr = sStr & "                   '12','��)�ڿ�',"
'        sStr = sStr & "                   '13','��)��ü',"
'        sStr = sStr & "                   '14','��)����(��)',"
'        sStr = sStr & "                   '15','��)�ι�����',"
'        sStr = sStr & "                   '16','��)�ڿ�����',"
'        sStr = sStr & "                   '21','������ι�',"
'        sStr = sStr & "                   '22','������ڿ�',"
'    End If
''<< �迭 >> : 2008.01.09
'    If Trim(basModule.SchCD) = "K" Or Trim(basModule.SchCD) = "W" Or Trim(basModule.SchCD) = "Q" Then
'        sStr = sStr & "                   '04','�ָ�����',"
'        sStr = sStr & "                   '05','�ָ��Ǵ�',"
'        sStr = sStr & "                   '06','�߰�����',"
'        sStr = sStr & "                   '07','�߰��Ǵ�',"
'
'        sStr = sStr & "                   '11','�������ι�',"
'        sStr = sStr & "                   '12','�������ڿ�',"
'
'        sStr = sStr & "                   '16','�������ι�16',"
'        sStr = sStr & "                   '17','�������ڿ�17',"
'    End If
''<< �迭 >> : 2008.02.15
'    If Trim(basModule.SchCD) = "S" Then
'        sStr = sStr & "                   '03','��ü��',"
'        'sStr = sStr & "                   '04','Ư���ڿ�',"
'
'        sStr = sStr & "                   '05','�����ι�',"
'        sStr = sStr & "                   '06','�����ڿ�',"
'
'        sStr = sStr & "                   '11','�ż��ι�',"
'        sStr = sStr & "                   '12','�ż��ڿ�',"
'
'        sStr = sStr & "                   '18','�ι������̾�',"
'        sStr = sStr & "                   '19','�ڿ������̾�',"
'
'    End If
''<< �迭 >> : 2008.02.15
'    If Trim(basModule.SchCD) = "P" Then         '< ����
'        sStr = sStr & "                   '03','Ư���ι�',"
'        sStr = sStr & "                   '04','Ư���ڿ�',"
'    End If
'
'    If Trim(basModule.SchCD) = "J" Then         '< ����
'        sStr = sStr & "                   '11','�ż��ι�',"
'        sStr = sStr & "                   '12','�ż��ڿ�',"
'
'        sStr = sStr & "                   '18','�ι������̾�',"
'        sStr = sStr & "                   '19','�ڿ������̾�',"
'
'    End If
'
''<< �迭 >> : 2009.01.09
'    If Trim(basModule.SchCD) = "B" Then         '< �λ�
'        sStr = sStr & "                   '05','Ư���ι�',"
'        sStr = sStr & "                   '06','Ư���ڿ�',"
'        sStr = sStr & "                   '07','������ι�',"
'        sStr = sStr & "                   '08','������ڿ�',"
'        sStr = sStr & "                   '09','��ȭ�ι�',"
'        sStr = sStr & "                   '10','��ȭ�ڿ�',"
'        sStr = sStr & "                   '23','�ι�PS��',"
'        sStr = sStr & "                   '24','�ڿ�PM��',"
'    End If
'
'    sStr = sStr & "                       '','��Ÿ') AS �迭,"
'
'    sStr = sStr & "     /* ��Ž, ��Ž �и� */"
'    sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'" & constSatamCodes(0) & "|') > 0 THEN          /* ��Ž-�ѱ��� */"
'    sStr = sStr & "             '" & constSatamCodes(0) & "'"
'    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'51|') > 0 THEN     /* ��Ž-����1 */"
'    sStr = sStr & "             '51'"
'    sStr = sStr & "         ELSE"
'    sStr = sStr & "             '00'"
'    sStr = sStr & "         END END SEL1,"
'    sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'" & constSatamCodes(1) & "|') > 0 THEN          /* ��Ž-����� */"
'    sStr = sStr & "             '" & constSatamCodes(1) & "'"
'    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'52|') > 0 THEN     /* ��Ž-ȭ��1 */"
'    sStr = sStr & "             '52'"
'    sStr = sStr & "         ELSE"
'    sStr = sStr & "             '00'"
'    sStr = sStr & "         END END SEL2,"
'    sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'" & constSatamCodes(2) & "|') > 0 THEN          /* ��Ž-���ƽþƻ� */"
'    sStr = sStr & "             '" & constSatamCodes(2) & "'"
'    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'53|') > 0 THEN     /* ��Ž-�������1 */"
'    sStr = sStr & "             '53'"
'    sStr = sStr & "         ELSE"
'    sStr = sStr & "             '00'"
'    sStr = sStr & "         END END SEL3,"
'    sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'" & constSatamCodes(3) & "|') > 0 THEN          /* ��Ž-�ѱ����� */"
'    sStr = sStr & "             '" & constSatamCodes(3) & "'"
'    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'54|') > 0 THEN     /* ��Ž-��������1 */"
'    sStr = sStr & "             '54'"
'    sStr = sStr & "         ELSE"
'    sStr = sStr & "             '00'"
'    sStr = sStr & "         END END SEL4,"
'    sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'" & constSatamCodes(4) & "|') > 0 THEN          /* ��Ž-�������� */"
'    sStr = sStr & "             '" & constSatamCodes(4) & "'"
'    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'55|') > 0 THEN     /* ��Ž-����2 */"
'    sStr = sStr & "             '55'"
'    sStr = sStr & "         ELSE"
'    sStr = sStr & "             '00'"
'    sStr = sStr & "         END END SEL5,"
'    sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'" & constSatamCodes(5) & "|') > 0 THEN          /* ��Ž-��Ȱ������ */"
'    sStr = sStr & "             '" & constSatamCodes(5) & "'"
'    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'56|') > 0 THEN     /* ��Ž-ȭ��2 */"
'    sStr = sStr & "             '56'"
'    sStr = sStr & "         ELSE"
'    sStr = sStr & "             '00'"
'    sStr = sStr & "         END END SEL6,"
'    sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'" & constSatamCodes(6) & "|') > 0 THEN          /* ��Ž-�����ͻ�� */"
'    sStr = sStr & "             '" & constSatamCodes(6) & "'"
'    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'57|') > 0 THEN     /* ��Ž-�������2 */"
'    sStr = sStr & "             '57'"
'    sStr = sStr & "         ELSE"
'    sStr = sStr & "             '00'"
'    sStr = sStr & "         END END SEL7,"
'    sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'" & constSatamCodes(7) & "|') > 0 THEN          /* ��Ž-������ġ */"
'    sStr = sStr & "             '" & constSatamCodes(7) & "'"
'    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'58|') > 0 THEN     /* ��Ž-��������2 */"
'    sStr = sStr & "             '58'"
'    sStr = sStr & "         ELSE"
'    sStr = sStr & "             '00'"
'    sStr = sStr & "         END END SEL8,"
'    sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'" & constSatamCodes(8) & "|') > 0 THEN          /* ��Ž-���� */"
'    sStr = sStr & "             '" & constSatamCodes(8) & "'"
'    sStr = sStr & "         ELSE"
'    sStr = sStr & "             '00'"
'    sStr = sStr & "         END SEL9,"
'    sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'" & constSatamCodes(9) & "|') > 0 THEN          /* ��Ž-��ȸ��ȭ */"
'    sStr = sStr & "             '" & constSatamCodes(9) & "'"
'    sStr = sStr & "         ELSE"
'    sStr = sStr & "             '00'"
'    sStr = sStr & "         END SEL10,"
'    sStr = sStr & "  "
'    sStr = sStr & "      /* ��2�ܱ��� & ���� */"
'    sStr = sStr & "              CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'31|') > 0 THEN '����'"
'    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'32|') > 0 THEN '�Ͼ�'"
'    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'33|') > 0 THEN '����'"
'    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'34|') > 0 THEN '�Ҿ�'"
'    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'35|') > 0 THEN '�߾�'"
'    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'36|') > 0 THEN '�ѹ�'"
'    '<< ���� >> : 2008.01.09
'    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'37|') > 0 THEN '���'"
'    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'38|') > 0 THEN '����'"
'    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'39|') > 0 THEN '����'"
'    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'40|') > 0 THEN '�����'"
'    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'41|') > 0 THEN '����'"
'    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'42|') > 0 THEN '�ƶ���'"
'
'    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL4,'81|') > 0 THEN '����'"
'    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL4,'82|') > 0 THEN '�̻�'"
'    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL4,'83|') > 0 THEN 'Ȯ��'"
'    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL4,'84|') > 0 THEN '����'"
'    sStr = sStr & "         ELSE"
'    sStr = sStr & "             ' '"
'    sStr = sStr & "         END END END END END END END END END END END END END END END END ��2����,"
'    sStr = sStr & "  "
'    sStr = sStr & "      /* ��� */"
'    sStr = sStr & "         CASE WHEN INSTR(SEL5,'91|') > 0 THEN         /* ��� */"
'    sStr = sStr & "             '���'"
'    sStr = sStr & "         ELSE"
'    sStr = sStr & "             ' '"
'    sStr = sStr & "         END �����,"
'    sStr = sStr & "         CASE WHEN INSTR(SEL5,'92|') > 0 THEN         /* ���� */"
'    sStr = sStr & "             '����'"
'    sStr = sStr & "         ELSE"
'    sStr = sStr & "             ' '"
'    sStr = sStr & "         END �������,"
'    sStr = sStr & "         CASE WHEN INSTR(SEL5,'93|') > 0 THEN         /* �ܱ��� */"      '< ����
'    sStr = sStr & "             '�ܱ���'"
'    sStr = sStr & "         ELSE"
'    sStr = sStr & "             ' '"
'    sStr = sStr & "         END ��Ž���,"
'    sStr = sStr & "         CASE WHEN INSTR(SEL5,'94|') > 0 THEN         /*  */"            '< ����
'    sStr = sStr & "             ' '"
'    sStr = sStr & "         ELSE"
'    sStr = sStr & "             ' '"
'    sStr = sStr & "         END ��Ž���,"
'    sStr = sStr & "  "
'    sStr = sStr & "         CY_ACNT AS �������, TOT_AMT AS ��ü�ݾ�    ,"
'    sStr = sStr & "         NVL(BASE_AMT1    ,0) AS �⺻�ݾ�1  ,"
'    sStr = sStr & "         NVL(BASE_AMT2    ,0) AS �⺻�ݾ�2  ,"
'    sStr = sStr & "         NVL(BASE_AMT3    ,0) AS �⺻�ݾ�3  ,"
'    sStr = sStr & "         NVL(BASE_AMT4    ,0) AS �⺻�ݾ�4  ,"
'    sStr = sStr & "         NVL(BASE_AMT5    ,0) AS �⺻�ݾ�5  ,"
'    sStr = sStr & "         NVL(BASE_AMT6    ,0) AS �⺻�ݾ�6  ,"
'    sStr = sStr & "         NVL(BASE_AMT7    ,0) AS �⺻�ݾ�7  ,"
'    sStr = sStr & "         NVL(BASE_AMT8    ,0) AS �⺻�ݾ�8  ,"
'    sStr = sStr & "         NVL(BASE_AMT9    ,0) AS �⺻�ݾ�9  ,"
'    sStr = sStr & "         NVL(BASE_AMT10   ,0) AS �⺻�ݾ�10 ,"
'    sStr = sStr & "         NVL(TAMGU_AMT1   ,0) AS Ž�������ݾ�1 ,"
'    sStr = sStr & "         NVL(TAMGU_AMT2   ,0) AS Ž�������ݾ�2 ,"
'    sStr = sStr & "         NVL(TAMGU_AMT3   ,0) AS Ž�������ݾ�3 ,"
'    sStr = sStr & "         NVL(TAMGU_AMT4   ,0) AS Ž�������ݾ�4 ,"
'    sStr = sStr & "         NVL(TAMGU_AMT5   ,0) AS Ž�������ݾ�5 ,"
'    sStr = sStr & "         NVL(TAMGU_AMT6   ,0) AS Ž�������ݾ�6 ,"
'    sStr = sStr & "         NVL(TAMGU_AMT7   ,0) AS Ž�������ݾ�7 ,"
'    sStr = sStr & "         NVL(TAMGU_AMT8   ,0) AS Ž�������ݾ�8 ,"
'    sStr = sStr & "         NVL(TAMGU_AMT9   ,0) AS Ž�������ݾ�9 ,"
'    sStr = sStr & "         NVL(TAMGU_AMT10  ,0) AS Ž�������ݾ�10,"
'    sStr = sStr & "         NVL(TAMGU_AMT11  ,0) AS Ž�������ݾ�11,"
'    sStr = sStr & "         NVL(TAMGU_AMT12  ,0) AS Ž�������ݾ�12,"
'
'    sStr = sStr & "         K_NUM AS �������, M_NUM AS ��������, E_NUM AS ��������, "
'    sStr = sStr & "         (NVL(K_NUM,0)+NVL(M_NUM,0)+NVL(E_NUM,0)) AS ��ü����, N_NUM AS ���ŵ��, "
'
'
'    sStr = sStr & "         DECODE(SEL1_SCH,'N','�뷮��','K','����','S','����','P','���ĸ��̸�','M','�������̸�', 'W', '�ָ����Ǵ�','Q','�߰����Ǵ�','H','����', 'B','�λ�') AS ��1����,"
'    sStr = sStr & "         DECODE(SEL2_SCH,'N','�뷮��','K','����','S','����','P','���ĸ��̸�','M','�������̸�', 'W', '�ָ����Ǵ�','Q','�߰����Ǵ�','H','����', 'B','�λ�') AS ��2����,"
'
'    sStr = sStr & "         DECODE(PASS1,'N','�뷮��','K','����','S','����','P','���ĸ��̸�','M','�������̸�', 'W', '�ָ����Ǵ�','Q','�߰����Ǵ�','H','����', 'B','�λ�') AS �հ�1   ,"
'    sStr = sStr & "         DECODE(PASS2,'N','�뷮��','K','����','S','����','P','���ĸ��̸�','M','�������̸�', 'W', '�ָ����Ǵ�','Q','�߰����Ǵ�','H','����', 'B','�λ�') AS �հ�2   ,"
'    sStr = sStr & "         DECODE(PASS3,'N','�뷮��','K','����','S','����','P','���ĸ��̸�','M','�������̸�', 'W', '�ָ����Ǵ�','Q','�߰����Ǵ�','H','����', 'B','�λ�') AS �հ�3   ,"
'    sStr = sStr & "         DECODE(PASS4,'N','�뷮��','K','����','S','����','P','���ĸ��̸�','M','�������̸�', 'W', '�ָ����Ǵ�','Q','�߰����Ǵ�','H','����', 'B','�λ�') AS �հ�4   ,"
'
'
'    sStr = sStr & "         DECODE(SEX,'M','��','F','��') AS ����        , "
'    sStr = sStr & "         SUBSTR(ZIP,1,3)||'-'||SUBSTR(ZIP,4,3) AS �����ȣ, ADDR1 AS �����ּ�      , ADDR2 AS ���ּ�     ,"
'    sStr = sStr & "         TEL AS ��ȭ��ȣ, CEL AS �ڵ���        , EMAIL AS �̸���     ,"
'    sStr = sStr & "         HIGH_SCH AS ����б� , GRADE_YEAR AS �����⵵ ,"
'    sStr = sStr & "         PRNT_NM AS �кθ�� , DECODE(PRNT_RLTN,'1','��','2','��','3','��Ÿ') AS �кθ����, "
'    sStr = sStr & "         SUBSTR(PRNT_ZIP,1,3)||'-'||SUBSTR(PRNT_ZIP,4,3) AS �кθ�_�����ȣ, PRNT_ADDR1 AS �кθ�_�����ּ� , PRNT_ADDR2 AS �кθ�_���ּ�,"
'    sStr = sStr & "         PRNT_TEL AS �кθ�_��ȭ��ȣ  , PRNT_CEL AS �кθ�_�ڵ���   , PRNT_JOB AS �кθ�_����   , PRNT_W_TEL AS �кθ�_������ȭ ,"
'    sStr = sStr & "         PHOTO_PATH AS �����������, "
'    sStr = sStr & "         DECODE(R_WAY,'1','�п����','2','���ͳݵ��','3','�п����') AS ��Ϲ�ȣ, "
'    sStr = sStr & "         ORD_NO AS �ֹ���ȣ, "
'    sStr = sStr & "         ACID||EXMID AS �̹������ϸ�, "
'    'sStr = sStr & "         DECODE(ACID,'" & Trim(basModule.SchCD) & "','',ACID) AS WANT_ACID "
'    'sStr = sStr & "         DECODE(ACID,'" & Trim(basModule.SchCD) & "','" & Trim(basModule.SchCD) & "',ACID) AS WANT_ACID, "       '< TEST
'    sStr = sStr & "         REGDATE AS �������, GET_PAYGUBN(ORD_NO) AS ������, "
'    sStr = sStr & "         DECODE(MU_TYPE,'1','������','2','6�� �򰡿�','3','9�� �򰡿�','4','6�� �򰡿�','9','���ŵ��','5','9�� �򰡿�','') AS ���, "
'    sStr = sStr & "         CL_CLOSE AS �Ϸ���  "
'
'    sStr = sStr & "    FROM ( "
'
'            sStr = sStr & "  SELECT SCHNO           ,"
'            sStr = sStr & "         MAX(ACID      ) AS ACID       ,"
'            sStr = sStr & "         MAX(EXMID     ) AS EXMID      ,"
'            sStr = sStr & "         MAX(STDNM     ) AS STDNM      ,"
'            sStr = sStr & "         MAX(D_UNIVCD     ) AS D_UNIVCD      ,"
'            sStr = sStr & "         MAX(D_MAJORCD     ) AS D_MAJORCD      ,"
'            sStr = sStr & "         MAX(birth_ymd     ) AS birth_ymd      ,"
'            sStr = sStr & "         MAX(EXMTYPE   ) AS EXMTYPE    , MAX(KAEYOL    ) AS KAEYOL     ,"
'            sStr = sStr & "         MAX(SEL1      ) AS SEL1       , MAX(SEL2      ) AS SEL2       , MAX(SEL3      ) AS SEL3      , MAX(SEL4      ) AS SEL4      , MAX(SEL5      ) AS  SEL5      ,"
'            sStr = sStr & "         MAX(K_NUM     ) AS K_NUM      , MAX(M_NUM     ) AS M_NUM      , MAX(E_NUM     ) AS E_NUM     , MAX(TOT_NUM   ) AS TOT_NUM   , MAX(N_NUM     ) AS N_NUM     "
'            sStr = sStr & "         MAX(SEL1_SCH  ) AS SEL1_SCH   , MAX(SEL2_SCH  ) AS SEL2_SCH   ,"
'            sStr = sStr & "         MAX(PASS1     ) AS PASS1      , MAX(PASS2     ) AS PASS2      , MAX(PASS3     ) AS PASS3     , MAX(PASS4     ) AS PASS4     , MAX(CL_CLOSE  ) AS  CL_CLOSE  ,"
'            sStr = sStr & "         MAX(CY_ACNT   ) AS CY_ACNT    , MAX(TOT_AMT   ) AS TOT_AMT    ,"
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
'            sStr = sStr & "         MAX(PHOTO_PATH) AS PHOTO_PATH , MAX(R_WAY     ) AS R_WAY      , MAX(ORD_NO    ) AS ORD_NO    , "
'            sStr = sStr & "         MAX(TO_CHAR(REGDATE,'YYYY-MM-DD HH24:MI:SS')) AS REGDATE      , MAX(MU_TYPE   ) AS MU_TYPE   "
'
'            sStr = sStr & "    FROM ("
'            '---------------------------------------------------------------------------- ��ü�л� ��ȸ START
'            sStr = sStr & "          SELECT *"
'            sStr = sStr & "            FROM CLSTD01TB"
'            sStr = sStr & "           WHERE ACID = '" & Trim(basModule.SchCD) & "'"
'            sStr = sStr & "             AND EXMID > ' ' "
'            sStr = sStr & "             AND BIGO2 IS NULL "
'            sStr = sStr & "          UNION ALL"
'            '---------------------------------------------------------------------------- ��ü�л� ��ȸ END
'            '---------------------------------------------------------------------------- �հ��� ��ȸ START
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
'            '---------------------------------------------------------------------------- �հ��� ��ȸ END
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
'
'
''>> �п�
''        sTmp = Trim(basModule.SchCD)
''        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
''            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
'
'''>> �����ȣ
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
''>> �л���
''        If Trim(txtStdNM.Text) > " " Then
''            sTmp = "%" & Trim(txtStdNM.Text) & "%"
''            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
''                Set DBParam = DBCmd.CreateParameter("STDNM", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
''        End If
'
'
'    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic���·� ���ԵǸ� record count�� �� �� ����.
'    Do While DBRec.State And adStateExecuting
'        DoEvents
'    Loop
'
'    With DBRec
'        If .RecordCount = 0 Then
'
'            MsgBox "�ش���ȸ����ڰ� �����ϴ�.", vbExclamation + vbOKOnly, "��ü�л� ��ȸ"
'
'        ElseIf .RecordCount > 0 Then
'
'            '## ��������
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
'    MsgBox "�����ڷ� �ۼ��Ϸ��Ͽ����ϴ�.", vbInformation + vbOKOnly, "��ü�л� ��ȸ"
'
'    Set DBCmd = Nothing
'    Set DBRec = Nothing
'
'    Exit Sub
'
'ErrStmt1:
'    MsgBox "������ �������� ����ϼ���.", vbExclamation + vbOKOnly, Me.Caption
'    Exit Sub
'
'ErrStmt:
'    Set DBCmd = Nothing
'    Set DBRec = Nothing
'
'    MsgBox "��ü�л� ��ȸ�� ������ �߻��Ͽ����ϴ�." & vbCrLf & _
'           Trim(CStr(Err.Number)) & ":" & Trim(Err.Description), vbCritical + vbOKOnly, "��ü�л� ��ȸ"
'
'    On Error GoTo 0
End Sub

















'�л� �������κ� ��� ==========================================================================

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


'���
Private Sub sprPoint_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)

    
    '������ ��ȸ
        
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
    
    '> �ʱ�ȭ
    
    Dim sPyoJum         As String
    Dim sSubjCD         As String
    
    If cmdAddPoint.Tag = "ACC" Then Exit Sub
    If sprPoint.MaxRows = 0 Then Exit Sub
    
    sprPoint.Row = Row
    sprPoint.Col = 2
    
    If Trim(sprPoint.Text) = "" Then
        MsgBox "������ �����ϼ���.", vbExclamation + vbOKOnly, "�л�ǥ������ ó��"
        Exit Sub
    Else
        sSubjCD = Trim(sprPoint.Text)
    End If
    
    sprPoint.Col = 4
        sPyoJum = Trim(sprPoint.Text)
        If Trim(sPyoJum) = "" Then sPyoJum = "0"
    
    On Error GoTo ErrStmt
    
    sStr = ""
    
    '���ѿ� ���� ��û ���� ������
    If Trim(basModule.SchCD) = "K" Then
        sStr = sStr & " SELECT NVL(TO_CHAR(BAK_J),0) AS BAK_NUM"
    Else
        sStr = sStr & " SELECT NVL(TO_CHAR(DNG_J),0) AS BAK_NUM"
    End If
    
    sStr = sStr & "   FROM DMSITEMGR.DMINF28TB"
    sStr = sStr & "  WHERE YY = '2012'"
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
    


    
'>> �п�
'        sTmp = Trim(basModule.SchCD)
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
    
    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic���·� ���ԵǸ� record count�� �� �� ����.
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
        
            Select Case Trim(basModule.SchCD)
                Case "N", "B"
                    Select Case Trim(Right(cboKaeyol.Text, 30))
                        Case "01", "03", "07", "09", "11", "13"
                            '<< �ι�
                            sGbn = "CULT"
                            
                        Case "02", "04", "08", "10", "12", "14"
                            '<< �ڿ�
                            sGbn = "SCI"
                            
                        Case "05", "15"
                            '<< �ι�
                            sGbn = "CULT"
                        Case "06", "16"
                            '<< �ڿ�
                            sGbn = "SCI"
                            
                        '2011-01-11 ���ѿ� �뷮�� �λ꿡 ���� �����,PS,PM ���� �߰�
                        Case "21", "23"
                            '<< �ι�
                            sGbn = "CULT"
                        Case "22", "24"
                            '<< �ڿ�
                            sGbn = "SCI"
                    End Select
                
                    
                Case "S", "P", "J"
                    Select Case Trim(Right(cboKaeyol.Text, 30))
                        Case "01", "03", "05", "11", "18"                                             '< 2008.02.15 : �迭 - ����, ����, ����      2009.06.02 : �迭�߰�
                            '<< �ι�
                            sGbn = "CULT"
                            
                        Case "02", "04", "06", "08", "12", "19"                                       '< 2008.02.15 : �迭 - ����, ����, ����      2009.06.02 : �迭�߰�
                            '<< �ڿ�
                            sGbn = "SCI"
                            
                    End Select
                Case Else
                    Select Case Trim(Right(cboKaeyol.Text, 30))
                        Case "01", "03", "04", "06", "11", "16"                         '< 2008.01.10 : �迭 - ����
                            '<< �ι�
                            sGbn = "CULT"
                            
                        Case "02", "05", "07", "12", "17"                               '< 2008.01.10 : �迭 - ����
                            '<< �ڿ�
                            sGbn = "SCI"
                            
                    End Select
            End Select
    
            sComboList = ""
            
            If sGbn = "CULT" Then
                sComboList = sComboList & "���                     37" + Chr$(9)
                sComboList = sComboList & "��������                 43" + Chr$(9)
                sComboList = sComboList & "�ܱ���                   39" + Chr$(9)
                
                
                '��Ž �߰�.
                Dim ni As Long
                For ni = 0 To SATAM_COUNT - 1
                    sComboList = sComboList & constSatams(ni) & "               " & constSatamCodes(ni) + Chr$(9)
                Next ni
                
                sComboList = sComboList & "����                     31" + Chr$(9)
                sComboList = sComboList & "�Ͼ�                     32" + Chr$(9)
                sComboList = sComboList & "�����ĳ�                 33" + Chr$(9)
                sComboList = sComboList & "�Ҿ�                     34" + Chr$(9)
                sComboList = sComboList & "�߱���                   35" + Chr$(9)
                sComboList = sComboList & "�ѹ�                     36" + Chr$(9)
                sComboList = sComboList & "�ƶ���                   42"

            Else
                sComboList = sComboList & "���                     37" + Chr$(9)
                sComboList = sComboList & "��������                 38" + Chr$(9)
                sComboList = sComboList & "�ܱ���                   39" + Chr$(9)
                sComboList = sComboList & "����1                    51" + Chr$(9)
                sComboList = sComboList & "ȭ��1                    52" + Chr$(9)
                sComboList = sComboList & "�������1                    53" + Chr$(9)
                sComboList = sComboList & "��������1                54" + Chr$(9)
                sComboList = sComboList & "����2                    55" + Chr$(9)
                sComboList = sComboList & "ȭ��2                    56" + Chr$(9)
                sComboList = sComboList & "�������2                    57" + Chr$(9)
                sComboList = sComboList & "��������2                58" + Chr$(9)
                sComboList = sComboList & "������                   81" + Chr$(9)
                sComboList = sComboList & "�̻����                 82" + Chr$(9)
                sComboList = sComboList & "Ȯ�����                 83"

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
        sprPoint.TypeButtonText = "���"
        
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
        MsgBox "����� ������ �����ϴ�.", vbExclamation + vbOKOnly, "�л��������"
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
        
            sprPoint.Col = 7                '����ó��
            If sprPoint.value = 1 Then
                
                sStr = sStr & " DELETE CLSTD03TB "
                sStr = sStr & "  WHERE SCHNO   = '" & sSchNO & "'"
                sStr = sStr & "    AND SUB_ID  = '" & sSubID & "'"
            Else
                
                sprPoint.Col = 9
                If sprPoint.value = 1 Then              '���ŵ��
                    
                    sStr = sStr & " UPDATE CLSTD03TB "
                    sStr = sStr & "    SET SUB_NUM = '" & sSubNum & "', "
                    sStr = sStr & "        SUB_BAK = '" & sSubBak & "' "
                    sStr = sStr & "  WHERE SCHNO   = '" & sSchNO & "'"
                    sStr = sStr & "    AND SUB_ID  = '" & sSubID & "'"
                    
                Else
                
                    sprPoint.Col = 8                    '�űԵ��
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
    MsgBox "����Ͽ����ϴ�.", vbInformation + vbOKOnly, "�л����� ���"
    Exit Sub
    
ErrStmt:
    basDataBase.DBConn.RollbackTrans
    
    Set DBCmd = Nothing
    Set DBParam = Nothing
    
    MsgBox "��Ͻ� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "�л����� ���"
    
End Sub













Private Sub cmdAddPoint_Click()
    
    '������ ��ȸ
        
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
    
    
    '> �ʱ�ȭ
    If Trim(txtSchNo.Text) = "" Then
        MsgBox "�л��� �����ϼ���.", vbExclamation + vbOKOnly, "�л� ������"
        Exit Sub
    End If
    
    sprPoint.MaxRows = 0
    
    On Error GoTo ErrStmt
    
    sStr = ""
    sStr = sStr & "  SELECT SCHNO, SUB_ID,"
    

    

    sStr = sStr & "         CASE WHEN      INSTR(SUB_ID,'37') > 0 THEN     /* ��� */"
    sStr = sStr & "             '���'"
    sStr = sStr & "         ELSE CASE WHEN INSTR(SUB_ID,'38') > 0 THEN     /* �������� */"
    sStr = sStr & "             '��������'"
    sStr = sStr & "         ELSE CASE WHEN INSTR(SUB_ID,'39') > 0 THEN     /* �ܱ��� */"
    sStr = sStr & "             '�ܱ���' "

    sStr = sStr & "     /* ��Ž, ��Ž �и� */"

    For ni = 0 To SATAM_COUNT - 1
        sStr = sStr & "         ELSE CASE WHEN INSTR(SUB_ID,'" & constSatamCodes(ni) & "') > 0 THEN     /* ��Ž-" & constSatams(ni) & " */"
        sStr = sStr & "             '" & constSatams(ni) & "'"
    Next ni

    
    sStr = sStr & "         ELSE CASE WHEN INSTR(SUB_ID,'51') > 0 THEN     /* ��Ž-����1 */"
    sStr = sStr & "             '����1'"
    sStr = sStr & "         ELSE CASE WHEN INSTR(SUB_ID,'52') > 0 THEN     /* ��Ž-ȭ��1 */"
    sStr = sStr & "             'ȭ��1'"
    sStr = sStr & "         ELSE CASE WHEN INSTR(SUB_ID,'53') > 0 THEN     /* ��Ž-�������1 */"
    sStr = sStr & "             '�������1'"
    sStr = sStr & "         ELSE CASE WHEN INSTR(SUB_ID,'54') > 0 THEN     /* ��Ž-��������1 */"
    sStr = sStr & "             '����1'"
    sStr = sStr & "         ELSE CASE WHEN INSTR(SUB_ID,'55') > 0 THEN     /* ��Ž-����2 */"
    sStr = sStr & "             '����2'"
    sStr = sStr & "         ELSE CASE WHEN INSTR(SUB_ID,'56') > 0 THEN     /* ��Ž-ȭ��2 */"
    sStr = sStr & "             'ȭ��2'"
    sStr = sStr & "         ELSE CASE WHEN INSTR(SUB_ID,'57') > 0 THEN     /* ��Ž-�������2 */"
    sStr = sStr & "             '�������2'"
    sStr = sStr & "         ELSE CASE WHEN INSTR(SUB_ID,'58') > 0 THEN     /* ��Ž-��������2 */" '11
    sStr = sStr & "             '����2'"
    
    sStr = sStr & "      /* ��2�ܱ��� & ���� */"
    sStr = sStr & "         ELSE CASE WHEN INSTR(SUB_ID,'31') > 0 THEN '����'"
    sStr = sStr & "         ELSE CASE WHEN INSTR(SUB_ID,'32') > 0 THEN '�Ͼ�'"
    sStr = sStr & "         ELSE CASE WHEN INSTR(SUB_ID,'33') > 0 THEN '�����ĳ�'"
    sStr = sStr & "         ELSE CASE WHEN INSTR(SUB_ID,'34') > 0 THEN '�Ҿ�'"
    sStr = sStr & "         ELSE CASE WHEN INSTR(SUB_ID,'35') > 0 THEN '�߱���'"
    sStr = sStr & "         ELSE CASE WHEN INSTR(SUB_ID,'36') > 0 THEN '�ѹ�'"
    sStr = sStr & "         ELSE CASE WHEN INSTR(SUB_ID,'42') > 0 THEN '�ƶ���'"

    sStr = sStr & "         ELSE CASE WHEN INSTR(SUB_ID,'81') > 0 THEN '������'"
    sStr = sStr & "         ELSE CASE WHEN INSTR(SUB_ID,'82') > 0 THEN '�̻����'"
    sStr = sStr & "         ELSE CASE WHEN INSTR(SUB_ID,'83') > 0 THEN 'Ȯ�����'"
    sStr = sStr & "         ELSE CASE WHEN INSTR(SUB_ID,'43') > 0 THEN '��������'" '22
'    sStr = sStr & "         END END END END END END END END END END END END END END END END END END END END END END  SUBJNM, "
'
    For ni = 0 To SATAM_COUNT - 1
        sStr = sStr & " END "
    Next ni
    
    sStr = sStr & "         END END END END END END END END END END END END END END END END END END END END END END SUBJNM, "
    
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
    

    
'>> �п�
'        sTmp = Trim(basModule.SchCD)
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
    
    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic���·� ���ԵǸ� record count�� �� �� ����.
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
                    sprPoint.TypeButtonText = "���"
                    
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
    
    MsgBox "�л��ڷ� ��ȸ�� ������ �߻��Ͽ����ϴ�." & vbCrLf & _
           Trim(CStr(Err.Number)) & ":" & Trim(Err.Description), vbCritical + vbOKOnly, "�л�������ȸ"
    
    On Error GoTo 0
    
    fraPoint.Visible = True
    
    cmdAddPoint.Tag = ""
    
End Sub


Private Sub init_FraAddrForm()
    fpBirth_ymdS.Text = ""
    fpZip.Text = ""
    txtAddr1.Text = ""
    txtAddr2.Text = ""
    txtEmail.Text = ""
End Sub




'## ���׸� ��ȸ
Private Sub Label49_Click()         '< �ݱ�
    fraAddr.Visible = False
    
End Sub

Private Sub cmdChgAddr_Click()

    Call init_FraAddrForm ' ���ʱ�ȭ
    
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
        MsgBox "�л��� ��ȸ�ϼ���.", vbExclamation + vbOKOnly, "�󼼳��� ����"
        Exit Sub
    End If
    
    sBirth_ymd = Trim(fpBirth_ymdS.UnFmtText)
    sZip = Trim(fpZip.UnFmtText)
    sAddr1 = Trim(txtAddr1.Text)
    sAddr2 = Trim(txtAddr2.Text)
    sEmail = Trim(txtEmail.Text)
    
    If MsgBox("�� " & Trim(txtStdNM.Text) & " ��" & " �л��� �󼼳����� �����Ͻðڽ��ϱ�?", vbQuestion + vbYesNo, "�󼼳��� ����") = vbNo Then
        Exit Sub
    End If
    
    basDataBase.DBConn.BeginTrans
    
    Set DBCmd = New ADODB.Command
    Set DBParam = New ADODB.Parameter
    
    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
                

    
    
    sStr = ""
    sStr = sStr & "  UPDATE CLSTD01TB"
    sStr = sStr & "     SET birth_ymd = '" & Trim(fpBirth_ymdS.UnFmtText) & "',"
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
        MsgBox "����Ͽ����ϴ�.", vbInformation + vbOKOnly, "�󼼳��� ����"
    Else
        basDataBase.DBConn.RollbackTrans
        MsgBox "��Ͻ� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "�󼼳��� ����"
    End If
    
    Set DBCmd = Nothing
    Set DBParam = Nothing
    
    
    Exit Sub
ErrStmt:
    basDataBase.DBConn.RollbackTrans
    
    Set DBCmd = Nothing
    Set DBParam = Nothing
    
    
    MsgBox "��Ͻ� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "�󼼳��� ����"
End Sub



















'## �����׸� �ޱ�
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
    
    '> �ʱ�ȭ
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
            MsgBox "������ ������ �����ϴ�.", vbExclamation + vbOKOnly, Me.Caption
            Exit Sub
        End If
        
        sExcelFileName = .fileName
        
        ni = InStrRev(sExcelFileName, "\", -1, vbTextCompare)
        sExcelLogFile = Mid(sExcelFileName, 1, ni) & "\" & Mid(sExcelFileName, ni + 1, Len(sExcelFileName) - ni + 1 - 5)
        
    End With
    
    On Error GoTo 0
    
    On Error GoTo ErrStmt
    
    
    
    '>>>> �л� ���� ���� ������
    sStr = basCommonSTD.Get_StdExcuteSqlToExcel(cboKaeyol_F.Text)
    
    
        
    Set DBCmd = New ADODB.Command
    Set DBRec = New ADODB.Recordset
    Set DBParam = New ADODB.Parameter
    
    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
    DBCmd.CommandText = sStr
    DBCmd.CommandType = adCmdText
    DBCmd.CommandTimeout = 30
    

    
'>> �п�
'        sTmp = Trim(basModule.SchCD)
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
        
''>> �����ȣ
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
'>> �л���
'        If Trim(txtStdNM.Text) > " " Then
'            sTmp = "%" & Trim(txtStdNM.Text) & "%"
'            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'                Set DBParam = DBCmd.CreateParameter("STDNM", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
'        End If
        
    
    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic���·� ���ԵǸ� record count�� �� �� ����.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop
    
    With DBRec
        If .RecordCount = 0 Then
            
            MsgBox "�ش���ȸ����ڰ� �����ϴ�.", vbExclamation + vbOKOnly, "��ü�л� ��ȸ"
            
        ElseIf .RecordCount > 0 Then
            
            '## ��������
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
    '��ü�л��ڷ�
    nRet = sprStdData.ExportToExcel(sExcelFileName, "Sheet1", sExcelLogFile)
    MsgBox "�����ڷ� �ۼ��Ϸ��Ͽ����ϴ�.", vbInformation + vbOKOnly, "��ü�л� ��ȸ"
    
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    Exit Sub
    
ErrStmt1:
    MsgBox "������ �������� ����ϼ���.", vbExclamation + vbOKOnly, Me.Caption
    Exit Sub
    
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    MsgBox "��ü�л� ��ȸ�� ������ �߻��Ͽ����ϴ�." & vbCrLf & _
           Trim(CStr(Err.Number)) & ":" & Trim(Err.Description), vbCritical + vbOKOnly, "��ü�л� ��ȸ"
    
    On Error GoTo 0
End Sub








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









'## ���系�� ����
Private Sub cmdPayChg_Click()
    
    If Trim(txtSchNo.Text) = "" Then
        MsgBox "�л��� ��ȸ�ϼ���.", vbExclamation + vbOKOnly, Me.Caption
        Exit Sub
    End If
    
    txtPay.Text = ""
    txtPayChg.Text = ""
    
    With FraPay         '< �������� ��� : 2010.01.13
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
    '## ����ϱ�
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
    
    '> �����ȣ + 1
    sSql = ""
    sSql = sSql & " UPDATE CLSTD02TB "
    sSql = sSql & "    SET NOW_NUM = NOW_NUM + 1"
    Select Case Trim(basModule.SchCD)
        Case "K", "W", "Q"
            sSql = sSql & "  WHERE ACSID   = 'K'"
        Case Else
            sSql = sSql & "  WHERE ACSID   = '" & Trim(basModule.SchCD) & "'"
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

    '> ���纯��
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
        Select Case Trim(basModule.SchCD)
            Case "K", "W", "Q"
                sSql = sSql & "  WHERE ACSID   = 'K'"
            Case Else
                sSql = sSql & "  WHERE ACSID   = '" & Trim(basModule.SchCD) & "'"
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
        
        DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic���·� ���ԵǸ� record count�� �� �� ����.
        Do While DBRec.State And adStateExecuting
            DoEvents
        Loop
        
        sNo = ""
        If DBRec.RecordCount = 1 Then
            DBRec.MoveFirst
            sNo = Trim(DBRec.Fields("TN"))
        End If
            
        
    '> ���纯��
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
    MsgBox "����Ͽ����ϴ�.", vbInformation + vbOKOnly, Me.Caption
    
    Exit Sub
Err:
        
    Set DBCmd = Nothing
    Set DBRec = Nothing
    MsgBox "��Ͻ� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, Me.Caption
    
End Sub



































