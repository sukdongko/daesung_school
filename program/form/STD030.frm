VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form STD030 
   Caption         =   "���л��� >> ��ϱ� �� ������� �ο� OLD"
   ClientHeight    =   11310
   ClientLeft      =   45
   ClientTop       =   2055
   ClientWidth     =   19110
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   11310
   ScaleWidth      =   19110
   Begin VB.Frame Frame1 
      BackColor       =   &H0082C8E8&
      BorderStyle     =   0  '����
      Caption         =   "Frame1"
      Height          =   1095
      Left            =   30
      TabIndex        =   52
      Top             =   30
      Width           =   15045
      Begin VB.Frame Frame2 
         BackColor       =   &H00D2EAF5&
         BorderStyle     =   0  '����
         Caption         =   "Frame2"
         Height          =   1035
         Left            =   30
         TabIndex        =   53
         Top             =   30
         Width           =   14985
         Begin VB.ComboBox cboExmType 
            Height          =   300
            Left            =   5310
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   2
            Top             =   165
            Width           =   1005
         End
         Begin VB.ComboBox cboKaeyol 
            Height          =   300
            Left            =   3660
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   7
            Top             =   600
            Width           =   975
         End
         Begin VB.TextBox txtStdNM 
            Height          =   345
            Left            =   5310
            TabIndex        =   8
            Text            =   "txtStdNM"
            Top             =   578
            Width           =   1005
         End
         Begin VB.ComboBox cboHakwon 
            Height          =   300
            Left            =   3660
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   1
            Top             =   165
            Width           =   975
         End
         Begin VB.CommandButton cmdFind 
            Caption         =   "��ȸ�ϱ� (&F)"
            Height          =   450
            Left            =   480
            TabIndex        =   0
            Top             =   30
            Width           =   1365
         End
         Begin EditLib.fpMask fpJumin 
            Height          =   345
            Left            =   7470
            TabIndex        =   9
            Top             =   585
            Width           =   1515
            _Version        =   196608
            _ExtentX        =   2672
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
            Mask            =   "999999-9999999"
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
            TabIndex        =   3
            Top             =   143
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
            Left            =   8580
            TabIndex        =   4
            Top             =   143
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
         Begin VB.Frame Frame3 
            BackColor       =   &H00D2EAF5&
            Height          =   555
            Left            =   120
            TabIndex        =   54
            Top             =   450
            Width           =   2955
            Begin VB.OptionButton optOkY 
               BackColor       =   &H00D2EAF5&
               Caption         =   "�ο��� �л�"
               Height          =   285
               Left            =   1620
               TabIndex        =   6
               Top             =   180
               Width           =   1275
            End
            Begin VB.OptionButton optOkN 
               BackColor       =   &H00D2EAF5&
               Caption         =   "��ϱ� �ο��� �л�"
               Height          =   375
               Left            =   30
               TabIndex        =   5
               Top             =   150
               Width           =   1485
            End
         End
         Begin VB.Label Label9 
            Alignment       =   1  '������ ����
            BackStyle       =   0  '����
            Caption         =   "����"
            Height          =   210
            Left            =   4890
            TabIndex        =   96
            Top             =   210
            Width           =   375
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
            Left            =   150
            TabIndex        =   60
            Top             =   150
            Width           =   2625
         End
         Begin VB.Label Label28 
            Alignment       =   1  '������ ����
            BackStyle       =   0  '����
            Caption         =   "��  ��"
            Height          =   210
            Left            =   2640
            TabIndex        =   59
            Top             =   645
            Width           =   975
         End
         Begin VB.Label Label2 
            Alignment       =   1  '������ ����
            BackStyle       =   0  '����
            Caption         =   "�л���"
            Height          =   210
            Left            =   4290
            TabIndex        =   58
            Top             =   645
            Width           =   975
         End
         Begin VB.Label Label3 
            Alignment       =   1  '������ ����
            BackStyle       =   0  '����
            Caption         =   "�ֹι�ȣ"
            Height          =   210
            Left            =   6450
            TabIndex        =   57
            Top             =   645
            Width           =   975
         End
         Begin VB.Label Label1 
            BackStyle       =   0  '����
            Caption         =   "�����ȣ            ����             ����"
            Height          =   210
            Left            =   6720
            TabIndex        =   56
            Top             =   210
            Width           =   3405
         End
         Begin VB.Label Label4 
            Alignment       =   1  '������ ����
            BackStyle       =   0  '����
            Caption         =   "����п�"
            Height          =   210
            Left            =   2640
            TabIndex        =   55
            Top             =   210
            Width           =   975
         End
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C6AD84&
      BorderStyle     =   0  '����
      Caption         =   "Frame4"
      Height          =   8415
      Left            =   30
      TabIndex        =   45
      Top             =   1170
      Width           =   15045
      Begin VB.Frame Frame5 
         BackColor       =   &H00F7EFE7&
         BorderStyle     =   0  '����
         Caption         =   "Frame5"
         Height          =   8355
         Left            =   30
         TabIndex        =   46
         Top             =   30
         Width           =   14985
         Begin VB.CheckBox chkAll 
            BackColor       =   &H00D2EAF5&
            Caption         =   "���"
            Height          =   225
            Left            =   6810
            TabIndex        =   42
            Top             =   720
            Width           =   675
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "�л� ������ ��� (&S)"
            Height          =   500
            Left            =   12660
            TabIndex        =   44
            Top             =   7680
            Width           =   2000
         End
         Begin VB.Frame fraAMT 
            BackColor       =   &H00C6AD84&
            BorderStyle     =   0  '����
            Caption         =   "Frame8"
            Height          =   7545
            Left            =   30
            TabIndex        =   61
            Top             =   540
            Width           =   2235
            Begin VB.Frame fraBase 
               BackColor       =   &H00C0C0FF&
               BorderStyle     =   0  '����
               Caption         =   "Frame7"
               Height          =   3105
               Left            =   30
               TabIndex        =   64
               Top             =   30
               Width           =   2175
               Begin VB.CommandButton cmdAmt 
                  Caption         =   "�ݾ׵��(&T)"
                  Height          =   360
                  Left            =   900
                  TabIndex        =   41
                  Top             =   30
                  Width           =   1230
               End
               Begin EditLib.fpDoubleSingle fpBase 
                  Height          =   300
                  Index           =   1
                  Left            =   930
                  TabIndex        =   14
                  Top             =   450
                  Width           =   1155
                  _Version        =   196608
                  _ExtentX        =   2037
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
                  TabIndex        =   15
                  Top             =   780
                  Width           =   1155
                  _Version        =   196608
                  _ExtentX        =   2037
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
                  TabIndex        =   16
                  Top             =   1110
                  Width           =   1155
                  _Version        =   196608
                  _ExtentX        =   2037
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
                  TabIndex        =   17
                  Top             =   1440
                  Width           =   1155
                  _Version        =   196608
                  _ExtentX        =   2037
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
                  Index           =   8
                  Left            =   930
                  TabIndex        =   21
                  Top             =   2760
                  Width           =   1155
                  _Version        =   196608
                  _ExtentX        =   2037
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
                  Index           =   5
                  Left            =   930
                  TabIndex        =   18
                  Top             =   1770
                  Width           =   1155
                  _Version        =   196608
                  _ExtentX        =   2037
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
                  Index           =   6
                  Left            =   930
                  TabIndex        =   19
                  Top             =   2100
                  Width           =   1155
                  _Version        =   196608
                  _ExtentX        =   2037
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
                  Index           =   7
                  Left            =   930
                  TabIndex        =   20
                  Top             =   2430
                  Width           =   1155
                  _Version        =   196608
                  _ExtentX        =   2037
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
               Begin VB.Label Label44 
                  Alignment       =   1  '������ ����
                  BackStyle       =   0  '����
                  Caption         =   "���4"
                  Height          =   210
                  Left            =   -120
                  TabIndex        =   95
                  Top             =   2835
                  Width           =   945
               End
               Begin VB.Label Label43 
                  Alignment       =   1  '������ ����
                  BackStyle       =   0  '����
                  Caption         =   "���3"
                  Height          =   210
                  Left            =   -120
                  TabIndex        =   94
                  Top             =   2505
                  Width           =   945
               End
               Begin VB.Label Label42 
                  Alignment       =   1  '������ ����
                  BackStyle       =   0  '����
                  Caption         =   "���2"
                  Height          =   210
                  Left            =   -120
                  TabIndex        =   93
                  Top             =   2175
                  Width           =   945
               End
               Begin VB.Label Label41 
                  Alignment       =   1  '������ ����
                  BackStyle       =   0  '����
                  Caption         =   "���1"
                  Height          =   210
                  Left            =   -120
                  TabIndex        =   92
                  Top             =   1845
                  Width           =   945
               End
               Begin VB.Label Label40 
                  BackStyle       =   0  '����
                  Caption         =   ">�ݾ� ------"
                  BeginProperty Font 
                     Name            =   "����"
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
                  TabIndex        =   91
                  Top             =   120
                  Width           =   1665
               End
               Begin VB.Label Label15 
                  Alignment       =   1  '������ ����
                  BackStyle       =   0  '����
                  Caption         =   "����ȸ����"
                  Height          =   210
                  Left            =   0
                  TabIndex        =   68
                  Top             =   1485
                  Width           =   945
               End
               Begin VB.Label Label14 
                  Alignment       =   1  '������ ����
                  BackStyle       =   0  '����
                  Caption         =   "�α����"
                  Height          =   210
                  Left            =   90
                  TabIndex        =   67
                  Top             =   1155
                  Width           =   765
               End
               Begin VB.Label Label13 
                  Alignment       =   1  '������ ����
                  BackStyle       =   0  '����
                  Caption         =   "�����"
                  Height          =   210
                  Left            =   300
                  TabIndex        =   66
                  Top             =   825
                  Width           =   555
               End
               Begin VB.Label Label6 
                  Alignment       =   1  '������ ����
                  BackStyle       =   0  '����
                  Caption         =   "��ϱ�"
                  Height          =   210
                  Left            =   300
                  TabIndex        =   65
                  Top             =   525
                  Width           =   555
               End
            End
            Begin VB.Frame fraGwatam 
               BackColor       =   &H00C0FFC0&
               BorderStyle     =   0  '����
               Caption         =   "Frame9"
               Height          =   3195
               Left            =   1170
               TabIndex        =   63
               Top             =   3150
               Width           =   2175
               Begin EditLib.fpDoubleSingle fpGwatam 
                  Height          =   300
                  Index           =   1
                  Left            =   930
                  TabIndex        =   33
                  Top             =   330
                  Width           =   1155
                  _Version        =   196608
                  _ExtentX        =   2037
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
                  TabIndex        =   34
                  Top             =   660
                  Width           =   1155
                  _Version        =   196608
                  _ExtentX        =   2037
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
                  TabIndex        =   35
                  Top             =   990
                  Width           =   1155
                  _Version        =   196608
                  _ExtentX        =   2037
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
                  TabIndex        =   36
                  Top             =   1320
                  Width           =   1155
                  _Version        =   196608
                  _ExtentX        =   2037
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
                  TabIndex        =   37
                  Top             =   1770
                  Width           =   1155
                  _Version        =   196608
                  _ExtentX        =   2037
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
                  TabIndex        =   38
                  Top             =   2100
                  Width           =   1155
                  _Version        =   196608
                  _ExtentX        =   2037
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
                  TabIndex        =   39
                  Top             =   2430
                  Width           =   1155
                  _Version        =   196608
                  _ExtentX        =   2037
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
                  TabIndex        =   40
                  Top             =   2760
                  Width           =   1155
                  _Version        =   196608
                  _ExtentX        =   2037
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
               Begin VB.Label Label39 
                  BackStyle       =   0  '����
                  Caption         =   ">��Ž ------"
                  BeginProperty Font 
                     Name            =   "����"
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
                  TabIndex        =   90
                  Top             =   90
                  Width           =   1665
               End
               Begin VB.Label Label38 
                  Alignment       =   1  '������ ����
                  BackStyle       =   0  '����
                  Caption         =   "��������2"
                  Height          =   210
                  Left            =   -60
                  TabIndex        =   88
                  Top             =   2835
                  Width           =   915
               End
               Begin VB.Label Label37 
                  Alignment       =   1  '������ ����
                  BackStyle       =   0  '����
                  Caption         =   "���� 2"
                  Height          =   210
                  Left            =   300
                  TabIndex        =   87
                  Top             =   2505
                  Width           =   555
               End
               Begin VB.Label Label36 
                  Alignment       =   1  '������ ����
                  BackStyle       =   0  '����
                  Caption         =   "ȭ�� 2"
                  Height          =   210
                  Left            =   300
                  TabIndex        =   86
                  Top             =   2175
                  Width           =   555
               End
               Begin VB.Label Label35 
                  Alignment       =   1  '������ ����
                  BackStyle       =   0  '����
                  Caption         =   "���� 2"
                  Height          =   210
                  Left            =   300
                  TabIndex        =   85
                  Top             =   1845
                  Width           =   555
               End
               Begin VB.Label Label34 
                  Alignment       =   1  '������ ����
                  BackStyle       =   0  '����
                  Caption         =   "��������1"
                  Height          =   210
                  Left            =   -30
                  TabIndex        =   84
                  Top             =   1395
                  Width           =   885
               End
               Begin VB.Label Label33 
                  Alignment       =   1  '������ ����
                  BackStyle       =   0  '����
                  Caption         =   "���� 1"
                  Height          =   210
                  Left            =   300
                  TabIndex        =   83
                  Top             =   1065
                  Width           =   555
               End
               Begin VB.Label Label32 
                  Alignment       =   1  '������ ����
                  BackStyle       =   0  '����
                  Caption         =   "ȭ�� 1"
                  Height          =   210
                  Left            =   300
                  TabIndex        =   82
                  Top             =   735
                  Width           =   555
               End
               Begin VB.Label Label31 
                  Alignment       =   1  '������ ����
                  BackStyle       =   0  '����
                  Caption         =   "���� 1"
                  Height          =   210
                  Left            =   300
                  TabIndex        =   81
                  Top             =   390
                  Width           =   555
               End
            End
            Begin VB.Frame fraSatam 
               BackColor       =   &H00C0FFFF&
               BorderStyle     =   0  '����
               Caption         =   "Frame9"
               Height          =   4215
               Left            =   30
               TabIndex        =   62
               Top             =   3240
               Width           =   2175
               Begin EditLib.fpDoubleSingle fpSatam 
                  Height          =   300
                  Index           =   1
                  Left            =   930
                  TabIndex        =   22
                  Top             =   300
                  Width           =   1155
                  _Version        =   196608
                  _ExtentX        =   2037
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
                  TabIndex        =   23
                  Top             =   630
                  Width           =   1155
                  _Version        =   196608
                  _ExtentX        =   2037
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
                  TabIndex        =   24
                  Top             =   960
                  Width           =   1155
                  _Version        =   196608
                  _ExtentX        =   2037
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
                  TabIndex        =   25
                  Top             =   1290
                  Width           =   1155
                  _Version        =   196608
                  _ExtentX        =   2037
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
                  TabIndex        =   26
                  Top             =   1740
                  Width           =   1155
                  _Version        =   196608
                  _ExtentX        =   2037
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
                  TabIndex        =   27
                  Top             =   2070
                  Width           =   1155
                  _Version        =   196608
                  _ExtentX        =   2037
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
                  TabIndex        =   28
                  Top             =   2400
                  Width           =   1155
                  _Version        =   196608
                  _ExtentX        =   2037
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
                  TabIndex        =   29
                  Top             =   2730
                  Width           =   1155
                  _Version        =   196608
                  _ExtentX        =   2037
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
                  TabIndex        =   30
                  Top             =   3180
                  Width           =   1155
                  _Version        =   196608
                  _ExtentX        =   2037
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
                  TabIndex        =   31
                  Top             =   3510
                  Width           =   1155
                  _Version        =   196608
                  _ExtentX        =   2037
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
                  TabIndex        =   32
                  Top             =   3840
                  Width           =   1155
                  _Version        =   196608
                  _ExtentX        =   2037
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
                  BackStyle       =   0  '����
                  Caption         =   ">��Ž ------"
                  BeginProperty Font 
                     Name            =   "����"
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
                  TabIndex        =   89
                  Top             =   90
                  Width           =   1665
               End
               Begin VB.Label Label30 
                  Alignment       =   1  '������ ����
                  BackStyle       =   0  '����
                  Caption         =   "��������"
                  Height          =   210
                  Left            =   120
                  TabIndex        =   80
                  Top             =   3900
                  Width           =   735
               End
               Begin VB.Label Label29 
                  Alignment       =   1  '������ ����
                  BackStyle       =   0  '����
                  Caption         =   "������ȸ"
                  Height          =   210
                  Left            =   120
                  TabIndex        =   79
                  Top             =   3585
                  Width           =   735
               End
               Begin VB.Label Label27 
                  Alignment       =   1  '������ ����
                  BackStyle       =   0  '����
                  Caption         =   "��ȸ��ȭ"
                  Height          =   210
                  Left            =   90
                  TabIndex        =   78
                  Top             =   3255
                  Width           =   765
               End
               Begin VB.Label Label26 
                  Alignment       =   1  '������ ����
                  BackStyle       =   0  '����
                  Caption         =   "��ġ"
                  Height          =   210
                  Left            =   300
                  TabIndex        =   77
                  Top             =   2805
                  Width           =   555
               End
               Begin VB.Label Label25 
                  Alignment       =   1  '������ ����
                  BackStyle       =   0  '����
                  Caption         =   "�ѱ�����"
                  Height          =   210
                  Left            =   120
                  TabIndex        =   76
                  Top             =   2475
                  Width           =   735
               End
               Begin VB.Label Label23 
                  Alignment       =   1  '������ ����
                  BackStyle       =   0  '����
                  Caption         =   "��������"
                  Height          =   210
                  Left            =   90
                  TabIndex        =   75
                  Top             =   2145
                  Width           =   765
               End
               Begin VB.Label Label22 
                  Alignment       =   1  '������ ����
                  BackStyle       =   0  '����
                  Caption         =   "�����"
                  Height          =   210
                  Left            =   300
                  TabIndex        =   74
                  Top             =   1800
                  Width           =   555
               End
               Begin VB.Label Label21 
                  Alignment       =   1  '������ ����
                  BackStyle       =   0  '����
                  Caption         =   "�ѱ�������"
                  Height          =   210
                  Left            =   0
                  TabIndex        =   73
                  Top             =   1365
                  Width           =   945
               End
               Begin VB.Label Label20 
                  Alignment       =   1  '������ ����
                  BackStyle       =   0  '����
                  Caption         =   "����"
                  Height          =   210
                  Left            =   90
                  TabIndex        =   72
                  Top             =   1035
                  Width           =   765
               End
               Begin VB.Label Label19 
                  Alignment       =   1  '������ ����
                  BackStyle       =   0  '����
                  Caption         =   "����"
                  Height          =   210
                  Left            =   300
                  TabIndex        =   71
                  Top             =   705
                  Width           =   555
               End
               Begin VB.Label Label17 
                  Alignment       =   1  '������ ����
                  BackStyle       =   0  '����
                  Caption         =   "����"
                  Height          =   210
                  Left            =   300
                  TabIndex        =   69
                  Top             =   375
                  Width           =   555
               End
            End
         End
         Begin FPSpread.vaSpread sprTamgu 
            Height          =   7035
            Left            =   2280
            TabIndex        =   43
            Top             =   570
            Width           =   12675
            _Version        =   393216
            _ExtentX        =   22357
            _ExtentY        =   12409
            _StockProps     =   64
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
            MaxCols         =   43
            ProcessTab      =   -1  'True
            SpreadDesigner  =   "STD030.frx":0000
         End
         Begin VB.Frame Frame6 
            BackColor       =   &H00C0E0FF&
            BorderStyle     =   0  '����
            Caption         =   "Frame6"
            Height          =   525
            Left            =   30
            TabIndex        =   47
            Top             =   30
            Width           =   9675
            Begin VB.CommandButton cmdSort 
               Caption         =   "����"
               Height          =   375
               Left            =   2010
               TabIndex        =   13
               Top             =   90
               Width           =   645
            End
            Begin EditLib.fpLongInteger fpSort 
               Height          =   315
               Index           =   0
               Left            =   2820
               TabIndex        =   10
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
               TabIndex        =   11
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
               TabIndex        =   12
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
               Left            =   1800
               TabIndex        =   51
               Top             =   180
               Width           =   645
            End
            Begin VB.Label Label7 
               Alignment       =   1  '������ ����
               BackStyle       =   0  '����
               Caption         =   "�����ȣ"
               Height          =   210
               Left            =   2700
               TabIndex        =   50
               Top             =   15
               Width           =   765
            End
            Begin VB.Label Label8 
               Alignment       =   1  '������ ����
               BackStyle       =   0  '����
               Caption         =   "����"
               Height          =   210
               Left            =   3540
               TabIndex        =   49
               Top             =   15
               Width           =   405
            End
            Begin VB.Label Label11 
               Alignment       =   1  '������ ����
               BackStyle       =   0  '����
               Caption         =   "�迭"
               Height          =   210
               Left            =   4080
               TabIndex        =   48
               Top             =   15
               Width           =   465
            End
         End
         Begin EditLib.fpLongInteger fpTotCnt 
            Height          =   345
            Left            =   13950
            TabIndex        =   97
            Top             =   120
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
         Begin VB.Label Label46 
            BackStyle       =   0  '����
            Caption         =   "��ȸ�ο�"
            Height          =   210
            Left            =   13110
            TabIndex        =   98
            Top             =   210
            Width           =   975
         End
      End
   End
   Begin VB.Label Label18 
      Alignment       =   1  '������ ����
      BackStyle       =   0  '����
      Caption         =   "�⺻ 1"
      Height          =   210
      Left            =   -30
      TabIndex        =   70
      Top             =   75
      Width           =   555
   End
End
Attribute VB_Name = "STD030"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################################################
'   �� �� ��  �� : �뼺�п� ���л���, �ݹ��� & �ð�ǥ ���α׷�
'   ����ý��۸� :
'   ��   ��   �� : STD030
'   �� ��  �� �� : ��ϱ� �� ������� �ο�
'
'   ��   ��   �� : 2007/08/27
'   ��   ��   �� : ���ϱ�
' --------------------------------------------------------------------------------------------------------------
' --------------------------------------------------------------------------------------------------------------
'                 ��     ��     ��     ��
' --------------------------------------------------------------------------------------------------------------
'   1. ������ :
'   2. ��  �� :
'################################################################################################################

Option Explicit

Private sini_Path      As String        '>> �뼺�п�

Private Sub Form_Terminate()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim sData       As String * 255
    Dim sGbn        As String
    Dim nRtn        As Long
    
    Dim sBase       As String           '<< �⺻�ݾ�
    Dim sSatam      As String           '<< ��Ž�ݾ�
    Dim sGwatam     As String           '<< ��Ž�ݾ�
    Dim sSort       As String           '<< sort
    
    Me.Move 0, 0, 15255, 9980
    
    Me.Tag = "LOAD"
    
        fraBase.Move 30, 30, fraAMT.Width - 60, fraBase.Height
        fraSatam.Move 30, 30 + fraBase.Height + 15, fraAMT.Width - 60, fraAMT.Height - fraBase.Height - 75:     fraSatam.Visible = False
        fraGwatam.Move 30, 30 + fraBase.Height + 15, fraAMT.Width - 60, fraAMT.Height - fraBase.Height - 75:    fraGwatam.Visible = False
        
        fpTotCnt.Value = 0
        
        With sprTamgu
            .ShadowColor = basModule.ShadowColor2
            .ShadowDark = basModule.ShadowDark2
            .ShadowText = basModule.ShadowText2
            .GridColor = basModule.GridColor2
            .GrayAreaBackColor = basModule.GrayAreaBackColor2
            
            .Tag = "0"      '<< ���߼���
        End With
        
        With cboHakwon
            .Clear
            .AddItem "�뷮��" & Space(30) & "N"
            .AddItem "����" & Space(30) & "K"
            .AddItem "����" & Space(30) & "S"
            .AddItem "���� M" & Space(30) & "P"
            .AddItem "���� M" & Space(30) & "M"
            
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
            End Select
        End With
        
        With cboExmType
            .Clear
            .AddItem "��ü" & Space(30) & "ALL"
            .AddItem "������" & Space(30) & "0"
            .AddItem "������" & Space(30) & "1"
            
            .ListIndex = 0
        End With
        
        
'>> �迭
        With cboKaeyol
            .Clear
            .AddItem "�ι�" & Space(30) & "01"
            .AddItem "�ڿ�" & Space(30) & "02"
            .ListIndex = 0
        End With
        
        sini_Path = App.Path & "\DAESUNG.INI"
        
        '>> ���α׷� INI ����
        If Dir(sini_Path) = "" Then                                     '<< ������ ������ ����
            sBase = insert_AMT_ini_File("BASE", "0/0/0/0/0/0/0/0/")
            sSatam = insert_AMT_ini_File("SATAM", "0/0/0/0/0/0/0/0/0/0/0/")
            sGwatam = insert_AMT_ini_File("GWATAM", "0/0/0/0/0/0/0/0/")
        End If
        
        sGbn = "STD030"
            sData = ""
            nRtn = basModule.GetPrivateProfileString(sGbn, "BASE", "", sData, 255, sini_Path)           '>> �⺻�ݾ�
            sBase = Trim(Replace(Left(Trim(sData), nRtn), Chr(0), "", 1, -1, vbTextCompare))
            If nRtn = 0 Then
                sBase = insert_AMT_ini_File("BASE", "0/0/0/0/0/0/0/0/")
            End If
            
            sData = ""
            nRtn = basModule.GetPrivateProfileString(sGbn, "SATAM", "", sData, 255, sini_Path)          '>> ��Ž�ݾ�
            sSatam = Trim(Replace(Left(Trim(sData), nRtn), Chr(0), "", 1, -1, vbTextCompare))
            If nRtn = 0 Then
                sSatam = insert_AMT_ini_File("SATAM", "0/0/0/0/0/0/0/0/0/0/0/")
            End If
            
            sData = ""
            nRtn = basModule.GetPrivateProfileString(sGbn, "GWATAM", "", sData, 255, sini_Path)         '>> ��Ž�ݾ�
            sGwatam = Trim(Replace(Left(Trim(sData), nRtn), Chr(0), "", 1, -1, vbTextCompare))
            If nRtn = 0 Then
                sGwatam = insert_AMT_ini_File("GWATAM", "0/0/0/0/0/0/0/0/")
            End If
            
            sData = ""
            nRtn = basModule.GetPrivateProfileString(sGbn, "SORT", "", sData, 255, sini_Path)         '>> SORT ����
            sSort = Trim(Replace(Left(Trim(sData), nRtn), Chr(0), "", 1, -1, vbTextCompare))
            If nRtn = 0 Then
                sSort = insert_AMT_ini_File("SORT", "0,3/1,2/2,1/")
            End If
            
        Call init_Form(sBase, sSatam, sGwatam, sSort)
        
    Me.Tag = ""
End Sub

'>> �ݾ׵��
Private Sub cboKaeyol_Click()
    Select Case Trim(Right(cboKaeyol.Text, 30))
        Case "01"
            fraSatam.Visible = True
            fraGwatam.Visible = False
            
            '>> spread header ����
            With sprTamgu
                .Row = SpreadHeader
                    .Col = 9:           .Text = "�ι� ���ÿ��� ����"
                    .Col = 21:          .Text = "�ι� �������"
                    .Col = 33:          .Text = "�ι� ���ÿ��� �ݾ׳���"
                    
                .Row = SpreadHeader + 1
                    .Col = 9:           .Text = "����"
                    .Col = 10:           .Text = "����"
                    .Col = 11:          .Text = "����"
                    .Col = 12:          .Text = "�ѱ�"
                    .Col = 13:          .Text = "����"
                    .Col = 14:          .Text = "����"
                    .Col = 15:          .Text = "����"
                    .Col = 16:          .Text = "��ġ"
                    .Col = 17:          .Text = "�繮"
                    .Col = 18:          .Text = "����"
                    .Col = 19:          .Text = "����"
                    
                    .Col = 20:          .Text = "��2��"
                    
                    .Col = 21:          .Text = "���"
                    .Col = 22:          .Text = "����"
                    .Col = 23:          .Text = "��Ž"
                    .Col = 24:          .Text = "��Ž"
                    
                    .Col = 33:          .Text = "����"
                    .Col = 34:          .Text = "����"
                    .Col = 35:          .Text = "����"
                    .Col = 36:          .Text = "�ѱ�"
                    .Col = 37:          .Text = "����"
                    .Col = 38:          .Text = "����"
                    .Col = 39:          .Text = "����"
                    .Col = 40:          .Text = "��ġ"
                    .Col = 41:          .Text = "�繮"
                    .Col = 42:          .Text = "����"
                    .Col = 43:          .Text = "����"
            End With
            
        Case "02"
            fraSatam.Visible = False
            fraGwatam.Visible = True
            
            '>> spread header ����
            With sprTamgu
                .Row = SpreadHeader
                    .Col = 9:           .Text = "�ڿ� ���ÿ��� ����"
                    .Col = 21:          .Text = "�ڿ� �������"
                    .Col = 33:          .Text = "�ڿ� ���ÿ��� �ݾ׳���"
                    
                .Row = SpreadHeader + 1
                    .Col = 9:           .Text = "��1"
                    .Col = 10:           .Text = "ȭ1"
                    .Col = 11:          .Text = "��1"
                    .Col = 12:          .Text = "��1"
                    .Col = 13:          .Text = "��2"
                    .Col = 14:          .Text = "ȭ2"
                    .Col = 15:          .Text = "��2"
                    .Col = 16:          .Text = "��2"
                    .Col = 17:          .Text = "-"
                    .Col = 18:          .Text = "-"
                    .Col = 19:          .Text = "-"
                    
                    .Col = 20:          .Text = "��2��"
                    
                    .Col = 21:          .Text = "���"
                    .Col = 22:          .Text = "����"
                    .Col = 23:          .Text = "��Ž"
                    .Col = 24:          .Text = "��Ž"
                    
                    .Col = 33:          .Text = "��1"
                    .Col = 34:          .Text = "ȭ1"
                    .Col = 35:          .Text = "��1"
                    .Col = 36:          .Text = "��1"
                    .Col = 37:          .Text = "��2"
                    .Col = 38:          .Text = "ȭ2"
                    .Col = 39:          .Text = "��2"
                    .Col = 40:          .Text = "��2"
                    .Col = 41:          .Text = "-"
                    .Col = 42:          .Text = "-"
                    .Col = 43:          .Text = "-"
                    
            End With
            
    End Select
End Sub

Private Sub init_Form(ByVal aBase As String, ByVal aSatam As String, ByVal aGwatam As String, ByVal aSort As String)
    Dim ni      As Integer
    Dim sDivs() As String
    Dim sDivC() As String
    
    optOkN.Value = True
    optOkY.Value = False

    txtStdNM.Text = ""
    fpJumin.Text = ""

    sprTamgu.MaxRows = 0
    
    For ni = 1 To 8 Step 1
        fpBase(ni).Value = 0
    Next ni

    For ni = 1 To 11 Step 1
        fpSatam(ni).Value = 0
    Next ni
    
    For ni = 1 To 8 Step 1
        fpGwatam(ni).Value = 0
    Next ni

    '>> ����
        sDivs() = Split(aBase, "/", -1, vbTextCompare)
        For ni = 0 To UBound(sDivs) - 1 Step 1
            fpBase(ni + 1).Value = CLng(sDivs(ni))
        Next ni
    
    '>> ��Ž
        sDivs() = Split(aSatam, "/", -1, vbTextCompare)
        For ni = 0 To UBound(sDivs) - 1 Step 1
            fpSatam(ni + 1).Value = CLng(sDivs(ni))
        Next ni
    
    '>> ��Ž
        sDivs() = Split(aGwatam, "/", -1, vbTextCompare)
        For ni = 0 To UBound(sDivs) - 1 Step 1
            fpGwatam(ni + 1).Value = CLng(sDivs(ni))
        Next ni
    
    '>> sort
        sDivs() = Split(aSort, "/", -1, vbTextCompare)
        For ni = 0 To UBound(sDivs) - 1 Step 1
            sDivC = Split(sDivs(ni), ",", -1, vbTextCompare)
            
            fpSort(CInt(sDivC(0))).Value = CInt(sDivC(1))
        Next ni
    
End Sub

Private Function insert_AMT_ini_File(ByVal aGbn As String, ByVal aData As String) As String
    Dim sGbn        As String
    Dim nRtn        As Long
    
    Dim sReturn     As String
    
    sGbn = "STD030"
         nRtn = basModule.WritePrivateProfileString(sGbn, aGbn, aData, sini_Path)        '<< �����ڿ� ���� �ݾ׵��
         
    insert_AMT_ini_File = aData
    
End Function




'>> �հݻ� ��ȸ
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
    
    chkAll.Value = 0
    sprTamgu.MaxRows = 0
    fpTotCnt.Value = 0
    
    sStr = ""
    sStr = sStr & "  SELECT SCHNO, EXMID, "
    sStr = sStr & "         STDNM,"
    sStr = sStr & "         CASE WHEN SEL1 > ' ' THEN"
    sStr = sStr & "             '01'"
    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' THEN"
    sStr = sStr & "             '02'"
    sStr = sStr & "         END END GAEYUL_CD,"
    sStr = sStr & "         CASE WHEN SEL1 > ' ' THEN"
    sStr = sStr & "             '��Ž'"
    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' THEN"
    sStr = sStr & "             '��Ž'"
    sStr = sStr & "         END END GAEYUL,"
    sStr = sStr & "  "
    sStr = sStr & "         CY_ACNT,"
    sStr = sStr & "         TOT_AMT,"
    sStr = sStr & "         0 AS CHKS,"
    sStr = sStr & "  "
    sStr = sStr & "     /* ��Ž, ��Ž �и� */"
    sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'01|') > 0 THEN          /* ��Ž-���� */"
    sStr = sStr & "             '01'"
    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'51|') > 0 THEN     /* ��Ž-����1 */"
    sStr = sStr & "             '51'"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             '00'"
    sStr = sStr & "         END END SEL1,"
    sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'02|') > 0 THEN          /* ��Ž-���� */"
    sStr = sStr & "             '02'"
    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'52|') > 0 THEN     /* ��Ž-ȭ��1 */"
    sStr = sStr & "             '52'"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             '00'"
    sStr = sStr & "         END END SEL2,"
    sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'03|') > 0 THEN          /* ��Ž-���� */"
    sStr = sStr & "             '03'"
    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'53|') > 0 THEN     /* ��Ž-����1 */"
    sStr = sStr & "             '53'"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             '00'"
    sStr = sStr & "         END END SEL3,"
    sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'04|') > 0 THEN          /* ��Ž-�ѱ������� */"
    sStr = sStr & "             '04'"
    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'54|') > 0 THEN     /* ��Ž-��������1 */"
    sStr = sStr & "             '54'"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             '00'"
    sStr = sStr & "         END END SEL4,"
    sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'05|') > 0 THEN          /* ��Ž-����� */"
    sStr = sStr & "             '05'"
    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'55|') > 0 THEN     /* ��Ž-����2 */"
    sStr = sStr & "             '55'"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             '00'"
    sStr = sStr & "         END END SEL5,"
    sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'06|') > 0 THEN          /* ��Ž-�������� */"
    sStr = sStr & "             '06'"
    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'56|') > 0 THEN     /* ��Ž-ȭ��2 */"
    sStr = sStr & "             '56'"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             '00'"
    sStr = sStr & "         END END SEL6,"
    sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'07|') > 0 THEN          /* ��Ž-�ѱ����� */"
    sStr = sStr & "             '07'"
    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'57|') > 0 THEN     /* ��Ž-����2 */"
    sStr = sStr & "             '57'"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             '00'"
    sStr = sStr & "         END END SEL7,"
    sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'08|') > 0 THEN          /* ��Ž-��ġ */"
    sStr = sStr & "             '08'"
    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'58|') > 0 THEN     /* ��Ž-��������2 */"
    sStr = sStr & "             '58'"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             '00'"
    sStr = sStr & "         END END SEL8,"
    sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'09|') > 0 THEN          /* ��Ž-��ȸ��ȭ */"
    sStr = sStr & "             '09'"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             '00'"
    sStr = sStr & "         END SEL9,"
    sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'10|') > 0 THEN          /* ��Ž-������ȸ */"
    sStr = sStr & "             '10'"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             '00'"
    sStr = sStr & "         END SEL10,"
    sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'11|') > 0 THEN          /* ��Ž-�������� */"
    sStr = sStr & "             '11'"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             '00'"
    sStr = sStr & "         END SEL11,"
    sStr = sStr & "  "
    sStr = sStr & "      /* ��2�ܱ��� & ���� */"
    sStr = sStr & "              CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'31|') > 0 THEN '31'"
    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'32|') > 0 THEN '32'"
    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'33|') > 0 THEN '33'"
    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'34|') > 0 THEN '34'"
    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'35|') > 0 THEN '35'"
    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'36|') > 0 THEN '36'"
    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL4,'81|') > 0 THEN '81'"
    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL4,'82|') > 0 THEN '82'"
    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL4,'83|') > 0 THEN '83'"
    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL4,'84|') > 0 THEN '84'"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             '00'"
    sStr = sStr & "         END END END END END END END END END END SEL_X2,"
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
    sStr = sStr & "         END SEL_N4,"
    sStr = sStr & "  "
    sStr = sStr & "         NVL(BASE_AMT1    ,0) AS BASE_AMT1  ,"
    sStr = sStr & "         NVL(BASE_AMT2    ,0) AS BASE_AMT2  ,"
    sStr = sStr & "         NVL(BASE_AMT3    ,0) AS BASE_AMT3  ,"
    sStr = sStr & "         NVL(BASE_AMT4    ,0) AS BASE_AMT4  ,"
    sStr = sStr & "         NVL(BASE_AMT5    ,0) AS BASE_AMT5  ,"
    sStr = sStr & "         NVL(BASE_AMT6    ,0) AS BASE_AMT6  ,"
    sStr = sStr & "         NVL(BASE_AMT7    ,0) AS BASE_AMT7  ,"
    sStr = sStr & "         NVL(BASE_AMT8    ,0) AS BASE_AMT8  ,"
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
    sStr = sStr & "         NVL(TAMGU_AMT11  ,0) AS TAMGU_AMT11"
    sStr = sStr & "  "
    sStr = sStr & "    FROM CLSTD01TB"
    sStr = sStr & "   WHERE (PASS1 = ? OR"
    sStr = sStr & "          PASS2 = ? OR"
    sStr = sStr & "          PASS3 = ? OR"
    sStr = sStr & "          PASS4 = ? )"
'>> ��ϱ� ��Ͽ���
    If optOkN.Value = True Then
        sStr = sStr & " AND TOT_AMT = 0 "
    ElseIf optOkY.Value = True Then
        sStr = sStr & " AND TOT_AMT > 0 "
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
'>> �迭
    Select Case Trim(Right(cboKaeyol, 30))
        Case "XX"
            ' no action
        Case "01", "03"
            sStr = sStr & " AND SEL1 > ' ' "
        Case "02"
            sStr = sStr & " AND SEL3 > ' ' "
    End Select
'>> �����ȣ
    If Trim(fpExmID_S.UnFmtText) > " " And Trim(fpExmID_E.UnFmtText) > " " Then
        sStr = sStr & " AND EXMID BETWEEN ? AND ? "
    ElseIf Trim(fpExmID_S.UnFmtText) > " " And Trim(fpExmID_E.UnFmtText) = " " Then
        sStr = sStr & " AND EXMID BETWEEN ? AND '99999' "
    ElseIf Trim(fpExmID_S.UnFmtText) = " " And Trim(fpExmID_E.UnFmtText) > " " Then
        sStr = sStr & " AND EXMID BETWEEN '00000' AND ? "
    ElseIf Trim(fpExmID_S.UnFmtText) = " " And Trim(fpExmID_E.UnFmtText) = " " Then
        ' no action
    End If
'>> �л���
    If Trim(txtStdNM.Text) > " " Then
        sStr = sStr & " AND STDNM LIKE ? "
    End If
'>> �ֹι�ȣ
    If Trim(fpJumin.UnFmtText) > " " Then
        sStr = sStr & " AND JUMIN LIKE ? "
    End If
'>> �ϷῩ�� : ����Ǹ� YYMM���� ��.
    sStr = sStr & " AND CL_CLOSE IS NULL "
    
    
    Set DBCmd = New ADODB.Command
    Set DBRec = New ADODB.Recordset
    Set DBParam = New ADODB.Parameter
    
    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
    DBCmd.CommandText = sStr
    DBCmd.CommandType = adCmdText
    DBCmd.CommandTimeout = 30
    
    '>> ���� �Ķ���Ͱ� ���� ������ �޸𸮿��� ������.
    For ni = 0 To DBCmd.Parameters.Count - 1 Step 1
        DBCmd.Parameters.Delete (0)
    Next ni
    
    '>> �п�
        For ni = 1 To 4 Step 1
            sTmp = Trim(Right(cboHakwon.Text, 30))
            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
                Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
        Next ni
        
    '>> �����ȣ
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
    
    '>> �л���
        If Trim(txtStdNM.Text) > " " Then
            sTmp = "%" & Trim(txtStdNM.Text) & "%"
            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
                Set DBParam = DBCmd.CreateParameter("STDNM", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
        End If
    '>> �ֹι�ȣ
        If Trim(fpJumin.UnFmtText) > " " Then
            sTmp = "%" & Trim(fpJumin.UnFmtText) & "%"
            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
                Set DBParam = DBCmd.CreateParameter("JUMIN", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
        End If
    
    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic���·� ���ԵǸ� record count�� �� �� ����.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop
    
    With DBRec
        If .RecordCount > 0 Then
            .MoveFirst
            
            For nRec = 1 To .RecordCount Step 1
            
                fpTotCnt.Value = fpTotCnt.Value + 1
            
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
                    sTmp = " ":     If IsNull(.Fields("TOT_AMT")) = False Then nTmp = CDbl(.Fields("TOT_AMT"))
                        Call basFunction.Set_SprType_Numeric(sprTamgu, 0, 0, 999999999, ",", nTmp)
                                    
                sprTamgu.Col = sprTamgu.Col + 1:    Call basFunction.Set_SprType_ChkBox(sprTamgu)
                
                
                sprTamgu.SetCellBorder sprTamgu.Col, sprTamgu.Row, sprTamgu.Col, sprTamgu.Row, 2, basModule.SectionColor1, CellBorderStyleSolid
                
                
            '>> ���ð��� (��Ž/ ��Ž)
                For ni = 1 To 11 Step 1
                
                    If ni Mod 4 = 1 Then
                        sprTamgu.SetCellBorder sprTamgu.Col, sprTamgu.Row, sprTamgu.Col, sprTamgu.Row, 2, basModule.SectionColor2, CellBorderStyleSolid
                    End If
                
                    sprTamgu.Col = sprTamgu.Col + 1
                    
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
                        Call basFunction.Set_SprType_Text(sprTamgu, "CENTER", "LEFT", 10, "")
                    Else
                        sTmp = IIf(Trim(.Fields(sGbn)) = "00", "", Trim(.Fields(sGbn)))
                        
                        If IsNull(.Fields(sGbn)) = False Then
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
                            End If
                            Call basFunction.Set_SprType_Text(sprTamgu, "CENTER", "LEFT", LenB(sTmp), sTmp)
                        End If
                    End If
                Next ni
                
                sprTamgu.SetCellBorder sprTamgu.Col, sprTamgu.Row, sprTamgu.Col, sprTamgu.Row, 2, basModule.SectionColor1, CellBorderStyleSolid
                
                sprTamgu.Col = sprTamgu.Col + 1
                If IsNull(.Fields("SEL_X2")) = True Then
                    Call basFunction.Set_SprType_Text(sprTamgu, "CENTER", "LEFT", 10, "")
                Else
                    If Trim(.Fields("SEL_X2")) = "00" Then
                        Call basFunction.Set_SprType_Text(sprTamgu, "CENTER", "LEFT", 10, "")
                    Else
                        Select Case Trim(.Fields("SEL_X2"))
                        
                            Case "31":  sTmp = "����"
                            Case "32":  sTmp = "�Ͼ�"
                            Case "33":  sTmp = "�����ĳľ�"
                            Case "34":  sTmp = "�Ҿ�"
                            Case "35":  sTmp = "�߱���"
                            Case "36":  sTmp = "�ѹ�"
                            
                            Case "81":  sTmp = "������"
                            Case "82":  sTmp = "�̻����"
                            Case "83":  sTmp = "Ȯ�����"
                            Case "84":  sTmp = "��������"
                            
                        End Select
                        Call basFunction.Set_SprType_Text(sprTamgu, "CENTER", "LEFT", LenB(sTmp), sTmp)
                    End If
                End If
                
                sprTamgu.SetCellBorder sprTamgu.Col, sprTamgu.Row, sprTamgu.Col, sprTamgu.Row, 2, basModule.SectionColor1, CellBorderStyleSolid
                
            '>> ���
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
                                    Case "91":  sTmp = "���"
                                    Case "92":  sTmp = "����"
                                    Case "93":  sTmp = "��Ž"
                                    Case "94":  sTmp = "��Ž"
                                    
                                End Select
                            End If
                            Call basFunction.Set_SprType_Text(sprTamgu, "CENTER", "LEFT", LenB(sTmp), sTmp)
                        End If
                    End If
                Next ni
                
                sprTamgu.SetCellBorder sprTamgu.Col, sprTamgu.Row, sprTamgu.Col, sprTamgu.Row, 2, basModule.SectionColor1, CellBorderStyleSolid
            
            '>> �ݾ�
                For ni = 1 To 8 Step 1
                    sprTamgu.Col = sprTamgu.Col + 1:    nTmp = 0
                    
                    If ni Mod 4 = 0 Then
                        sprTamgu.SetCellBorder sprTamgu.Col, sprTamgu.Row, sprTamgu.Col, sprTamgu.Row, 2, basModule.SectionColor2, CellBorderStyleSolid
                    End If
                    
                    sGbn = "BASE_AMT" & Trim(CStr(ni))
                    
                    If IsNull(.Fields(sGbn)) = False Then
                        nTmp = CDbl(.Fields(sGbn))
                    End If
                    Call basFunction.Set_SprType_Numeric(sprTamgu, 0, 0, 999999999, ",", nTmp)
                Next ni
                
            '>> Ž��
                For ni = 1 To 11 Step 1
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
                
                
            '## formula ##
                    sprTamgu.Col = 7
                    
                    sprTamgu.FormulaSync = False
                    sprTamgu.Formula = "SUM(X#:AQ#)"
                
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
            sprTamgu.Col = 8:       sprTamgu.Col2 = 24
            sprTamgu.BlockMode = True
                sprTamgu.Lock = True
                sprTamgu.Protect = True
            sprTamgu.BlockMode = False
            
        End If
    End With
    
    MsgBox "�л� ��ȸ�Ͽ����ϴ�.", vbInformation + vbOKOnly, "�հݻ� ��ȸ"
    
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    Exit Sub
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    On Error GoTo 0
    MsgBox "�հݻ� ��ȸ�� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "�հݻ� ��ȸ"
End Sub


'>> ���� ## multi ����
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
                    .Col = 8
                        .Value = 0
                Next nRow
                
                .Row = Row:     .Row2 = .Row
                .Col = 1:       .Col2 = .MaxCols
                .BlockMode = True
                .BackColor = basModule.SelectColor2
                .BackColorStyle = BackColorStyleUnderGrid
                .BlockMode = False
                
                .Col = 8
                    .Value = 1
                
                .Tag = Trim(CStr(Row))
            ElseIf .Tag > "0" Then
                .Row = Row
                .Col = 8
                If .Value = 1 Then
                    .Value = 0
                    
                    .Row = Row:     .Row2 = .Row
                    .Col = 1:       .Col2 = .MaxCols
                    .BlockMode = True
                    '.BackColor = basModule.BackColor2
                    .BackColor = &HFFFFFF
                    .BackColorStyle = BackColorStyleUnderGrid
                    .BlockMode = False
                    
                    '.Tag = Trim(CStr(Row))
                Else
                    .Value = 1
                    
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

Private Sub sprTamgu_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
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
                            .Col = 8
                                .Value = 1
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
    
    With sprTamgu
        If .MaxRows = 0 Then Exit Sub
            
        If chkAll.Value = 0 Then
            For ni = 1 To .MaxRows Step 1
                .Row = ni
                .Col = 8
                    .Value = 0
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
                .Col = 8
                    .Value = 1
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




'>> �ݾ׵��
Private Sub cmdAmt_Click()
    Dim ni      As Long
    Dim nRec    As Long
    
    Dim sTmp    As String
    
    If Me.Tag = "LOAD" Then Exit Sub
    
    With sprTamgu
    
        '## ����� �ݾ����� ini ���Ͽ� ���
    '> base
        sTmp = ""
        For ni = 1 To 8 Step 1
            sTmp = sTmp & Trim(CStr(fpBase(ni).Value)) & "/"
        Next ni
        sTmp = insert_AMT_ini_File("BASE", sTmp)
    '> satam
        sTmp = ""
        For ni = 1 To 11 Step 1
            sTmp = sTmp & Trim(CStr(fpSatam(ni).Value)) & "/"
        Next ni
        sTmp = insert_AMT_ini_File("SATAM", sTmp)
    '> gwatam
        sTmp = ""
        For ni = 1 To 8 Step 1
            sTmp = sTmp & Trim(CStr(fpGwatam(ni).Value)) & "/"
        Next ni
        sTmp = insert_AMT_ini_File("GWATAM", sTmp)
        
    
        If .MaxRows = 0 Then Exit Sub
        
        nRec = 0
        For ni = 1 To .MaxRows Step 1
            .Row = ni
            .Col = 8
            If .Value = 1 Then
                nRec = 1
                Exit For
            End If
        Next ni
        
        If nRec = 0 Then
            MsgBox "�ݾ� ����� �л��� �����Ͽ� �ֽʽÿ�.", vbExclamation + vbOKOnly, "�ݾ� ���"
            Exit Sub
        End If
        
        For ni = 1 To .MaxRows Step 1
            .Row = ni
            .Col = 8
            
            If .Value = 1 Then
                Select Case Trim(Right(cboKaeyol.Text, 30))
                
                '>> ��Ž
                    Case "01"
                    '>> �⺻�ݾ�
                        For nRec = 1 To 4 Step 1
                            .Col = 25 + nRec - 1
                            .Value = fpBase(nRec).Value
                        Next nRec
                        
                    '>> ����ݾ�
                        .Col = 21:  If StrComp(Trim(.Text), "���", vbTextCompare) = 0 Then .Col = 29:      .Value = fpBase(5).Value
                        .Col = 22:  If StrComp(Trim(.Text), "����", vbTextCompare) = 0 Then .Col = 30:      .Value = fpBase(6).Value
                        .Col = 23:  If StrComp(Trim(.Text), "��Ž", vbTextCompare) = 0 Then .Col = 31:      .Value = fpBase(7).Value
                        .Col = 24:  If StrComp(Trim(.Text), "��Ž", vbTextCompare) = 0 Then .Col = 32:      .Value = fpBase(8).Value
                        
                    '>> ��Ž�ݾ�
                        .Col = 9:   If StrComp(Trim(.Text), "����", vbTextCompare) = 0 Then .Col = 33:      .Value = fpSatam(1).Value
                        .Col = 10:  If StrComp(Trim(.Text), "����", vbTextCompare) = 0 Then .Col = 34:      .Value = fpSatam(2).Value
                        .Col = 11:  If StrComp(Trim(.Text), "����", vbTextCompare) = 0 Then .Col = 35:      .Value = fpSatam(3).Value
                        .Col = 12:  If StrComp(Trim(.Text), "�ѱ�", vbTextCompare) = 0 Then .Col = 36:      .Value = fpSatam(4).Value
                        .Col = 13:  If StrComp(Trim(.Text), "�����", vbTextCompare) = 0 Then .Col = 37:    .Value = fpSatam(5).Value
                        .Col = 14:  If StrComp(Trim(.Text), "����", vbTextCompare) = 0 Then .Col = 38:      .Value = fpSatam(6).Value
                        .Col = 15:  If StrComp(Trim(.Text), "����", vbTextCompare) = 0 Then .Col = 39:      .Value = fpSatam(7).Value
                        .Col = 16:  If StrComp(Trim(.Text), "��ġ", vbTextCompare) = 0 Then .Col = 40:      .Value = fpSatam(8).Value
                        .Col = 17:  If StrComp(Trim(.Text), "�繮", vbTextCompare) = 0 Then .Col = 41:      .Value = fpSatam(9).Value
                        .Col = 18:  If StrComp(Trim(.Text), "����", vbTextCompare) = 0 Then .Col = 42:      .Value = fpSatam(10).Value
                        .Col = 19:  If StrComp(Trim(.Text), "����", vbTextCompare) = 0 Then .Col = 43:      .Value = fpSatam(11).Value
                        
                        
                '>> ��Ž
                    Case "02"
                    '>> �⺻�ݾ�
                        For nRec = 1 To 4 Step 1
                            .Col = 25 + nRec - 1
                            .Value = fpBase(nRec).Value
                        Next nRec
                        
                    '>> ����ݾ�
                        .Col = 21:  If StrComp(Trim(.Text), "���", vbTextCompare) = 0 Then .Col = 29:      .Value = fpBase(5).Value
                        .Col = 22:  If StrComp(Trim(.Text), "����", vbTextCompare) = 0 Then .Col = 30:      .Value = fpBase(6).Value
                        .Col = 23:  If StrComp(Trim(.Text), "��Ž", vbTextCompare) = 0 Then .Col = 31:      .Value = fpBase(7).Value
                        .Col = 24:  If StrComp(Trim(.Text), "��Ž", vbTextCompare) = 0 Then .Col = 32:      .Value = fpBase(8).Value
                        
                    '>> ��Ž�ݾ�
                        .Col = 9:   If StrComp(Trim(.Text), "��1", vbTextCompare) = 0 Then .Col = 33:       .Value = fpGwatam(1).Value
                        .Col = 10:  If StrComp(Trim(.Text), "ȭ1", vbTextCompare) = 0 Then .Col = 34:       .Value = fpGwatam(2).Value
                        .Col = 11:  If StrComp(Trim(.Text), "��1", vbTextCompare) = 0 Then .Col = 35:       .Value = fpGwatam(3).Value
                        .Col = 12:  If StrComp(Trim(.Text), "��1", vbTextCompare) = 0 Then .Col = 36:       .Value = fpGwatam(4).Value
                        .Col = 13:  If StrComp(Trim(.Text), "��2", vbTextCompare) = 0 Then .Col = 37:       .Value = fpGwatam(5).Value
                        .Col = 14:  If StrComp(Trim(.Text), "ȭ2", vbTextCompare) = 0 Then .Col = 38:       .Value = fpGwatam(6).Value
                        .Col = 15:  If StrComp(Trim(.Text), "��2", vbTextCompare) = 0 Then .Col = 39:       .Value = fpGwatam(7).Value
                        .Col = 16:  If StrComp(Trim(.Text), "��2", vbTextCompare) = 0 Then .Col = 40:       .Value = fpGwatam(8).Value
                        .Col = 41:  .Value = 0
                        .Col = 42:  .Value = 0
                        .Col = 43:  .Value = 0
                    
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
                If fpSort(nj).Value = ni Then
                    nC = nC + 1
                    
                    Select Case nj
                        Case 0                      '<< �����ȣ
                            .SortKey(nC) = 2
                            .SortKeyOrder(nC) = SortKeyOrderAscending
                            
                            sSort = sSort & "0," & CInt(Trim(fpSort(0).Value)) & "/"
                        Case 1                      '<< ����
                            .SortKey(nC) = 3
                            .SortKeyOrder(nC) = SortKeyOrderAscending
                            
                            sSort = sSort & "1," & CInt(Trim(fpSort(1).Value)) & "/"
                        Case 2                      '<< �迭
                            .SortKey(nC) = 4
                            .SortKeyOrder(nC) = SortKeyOrderDescending
                            
                            sSort = sSort & "2," & CInt(Trim(fpSort(2).Value)) & "/"
                    End Select
                    
                End If
            Next nj
        Next ni
        
        .Sort -1, -1, -1, -1, SortByRow
        
        sR = insert_AMT_ini_File("SORT", sSort)
        
        sDivs() = Split(sR, "/", -1, vbTextCompare)
        For ni = 0 To UBound(sDivs) - 1 Step 1
            sDivC = Split(sDivs(ni), ",", -1, vbTextCompare)
            
            fpSort(CInt(sDivC(0))).Value = CInt(sDivC(1))
        Next ni
    
    End With
End Sub




'>> �л� ������ ���
Private Sub cmdSave_Click()
    Dim sTmp        As String
    Dim ni          As Integer
    
    Dim bRet        As Boolean
    Dim nCnt        As Long
    
    On Error GoTo ErrStmt
    
    '## ����� �ݾ����� ini ���Ͽ� ���
    '> base
        sTmp = ""
        For ni = 1 To 8 Step 1
            sTmp = sTmp & Trim(CStr(fpBase(ni).Value)) & "/"
        Next ni
        sTmp = insert_AMT_ini_File("BASE", sTmp)
    '> satam
        sTmp = ""
        For ni = 1 To 11 Step 1
            sTmp = sTmp & Trim(CStr(fpSatam(ni).Value)) & "/"
        Next ni
        sTmp = insert_AMT_ini_File("SATAM", sTmp)
    '> gwatam
        sTmp = ""
        For ni = 1 To 8 Step 1
            sTmp = sTmp & Trim(CStr(fpGwatam(ni).Value)) & "/"
        Next ni
        sTmp = insert_AMT_ini_File("GWATAM", sTmp)
    
    bRet = False
    
    '>> ����üũ
    With sprTamgu
        If .MaxRows = 0 Then Exit Sub
        
        nCnt = 0
        For ni = 1 To .MaxRows Step 1
            .Row = ni
            .Col = 8
            If .Value = 1 Then
                .Col = 6
                If Trim(.Text) = "" Then
                    MsgBox "������°� �����ϴ�.", vbExclamation + vbOKOnly, "��ϱ� ���"
                    Exit Sub
                End If
                
                .Col = 7
                If .Value = 0 Then
                    MsgBox "��ϱ��� 0 �Դϴ�.", vbExclamation + vbOKOnly, "��ϱ� ���"
                    Exit Sub
                End If
                
                .Col = 7
                If Trim(.Text) = "" Then
                    MsgBox "�ݾ��� �����ϴ�.", vbExclamation + vbOKOnly, "��ϱ� ���"
                    Exit Sub
                End If
                
                nCnt = nCnt + 1
            End If
        Next ni
        
        If nCnt = 0 Then
            MsgBox "���� 1�� �̻��Ͻʽÿ�.", vbExclamation + vbOKOnly, "��ϱ� ���"
            Exit Sub
        End If
    End With
    
    cmdSave.Enabled = False
        bRet = Save_Amt_Data
        
    cmdSave.Enabled = True
    
    If bRet = True Then
        MsgBox "��ϱ� ����Ͽ����ϴ�.", vbInformation + vbOKOnly, "��ϱ� ���"
    Else
        MsgBox "��ϱ� ��Ͻ� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "��ϱ� ���"
    End If
    
    Exit Sub
ErrStmt:
    MsgBox "��ϱ� ��Ͻ� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "��ϱ� ���"
    On Error GoTo 0
End Sub

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
    
    Dim nRec        As Long                                 '<< ó���ؾ� �� ��
    Dim nTot        As Long                                 '<< ó���� ��
    
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
        sprTamgu.Col = 8
        
        If sprTamgu.Value = 1 Then
        
            nRec = nRec + 1
            
            '>> ���� �Ķ���Ͱ� ���� ������ �޸𸮿��� ������.
            For ni = 0 To DBCmd.Parameters.Count - 1 Step 1
                DBCmd.Parameters.Delete (0)
            Next ni
        
            sStr = ""
            sStr = sStr & "  Update CLSTD01TB"
            sStr = sStr & "     SET CY_ACNT    = ?, "
            sStr = sStr & "         TOT_AMT    = ?, "
            
            sStr = sStr & "         BASE_AMT1  = ?, "
            sStr = sStr & "         BASE_AMT2  = ?, "
            sStr = sStr & "         BASE_AMT3  = ?, "
            sStr = sStr & "         BASE_AMT4  = ?, "
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
            sStr = sStr & "         TAMGU_AMT11= ? "
            
            sStr = sStr & "   WHERE SCHNO = ? "
            sStr = sStr & "     AND ACID  = ? "
        
            '>> ���¹�ȣ
                sprTamgu.Col = 6
                sTmp = Trim(sprTamgu.Text)
                nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
                    Set DBParam = DBCmd.CreateParameter("CY_ACNT", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
            '>> ��ü�ݾ�
                sprTamgu.Col = 7
                If Trim(sprTamgu.Text) = "" Then
                    nTmp = 0
                Else
                    nTmp = CLng(sprTamgu.Value)
                End If
                    Set DBParam = DBCmd.CreateParameter("TOT_AMT", adDouble, adParamInput, , nTmp):   DBCmd.Parameters.Append DBParam
            
            '>> ��ϱ�, �����, �α����, ����ȸ����, ���1 ~ ���4
                For ni = 25 To 32 Step 1
                    sprTamgu.Col = ni
                    If Trim(sprTamgu.Text) = "" Then
                        nTmp = 0
                    Else
                        nTmp = CLng(sprTamgu.Value)
                    End If
                        sTmp = "BASE_AMT" & Trim(CStr(ni - 23))
                        Set DBParam = DBCmd.CreateParameter(sTmp, adDouble, adParamInput, , nTmp):   DBCmd.Parameters.Append DBParam
                Next ni
            
            '>> ���ñݾ� 1 ~ 11
                For ni = 33 To 43 Step 1
                    sprTamgu.Col = ni
                    If Trim(sprTamgu.Text) = "" Then
                        nTmp = 0
                    Else
                        nTmp = CLng(sprTamgu.Value)
                    End If
                        sTmp = "TAMGU_AMT" & Trim(CStr(ni - 31))
                        Set DBParam = DBCmd.CreateParameter(sTmp, adDouble, adParamInput, , nTmp):   DBCmd.Parameters.Append DBParam
                Next ni
            
            '>> �л��ڵ�
                sprTamgu.Col = 1
                sTmp = Trim(sprTamgu.Text)
                nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
                    Set DBParam = DBCmd.CreateParameter("SCHHO", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
            '>> �п��ڵ� �з�
                sTmp = Trim(basModule.SchCD)
                nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
                    Set DBParam = DBCmd.CreateParameter("ACID", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
            
            
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
