VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form STD200 
   Caption         =   "���л��� >> ���� & ����Ŭ���� ��� �� ��ȸ"
   ClientHeight    =   10110
   ClientLeft      =   5415
   ClientTop       =   4290
   ClientWidth     =   15210
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   10110
   ScaleWidth      =   15210
   Begin FPSpread.vaSpread sprStdData 
      Height          =   2685
      Left            =   3990
      TabIndex        =   119
      Top             =   10020
      Visible         =   0   'False
      Width           =   22035
      _Version        =   393216
      _ExtentX        =   38867
      _ExtentY        =   4736
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
      SpreadDesigner  =   "STD200.frx":0000
   End
   Begin VB.Frame fraHak 
      BackColor       =   &H00808080&
      BorderStyle     =   0  '����
      Height          =   4515
      Left            =   12570
      TabIndex        =   143
      Top             =   5640
      Width           =   5475
      Begin VB.Frame Frame16 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  '����
         Height          =   4455
         Left            =   30
         TabIndex        =   144
         Top             =   30
         Width           =   5415
         Begin VB.TextBox txtFHak 
            Height          =   345
            Left            =   870
            TabIndex        =   81
            Text            =   "txtFHak"
            Top             =   120
            Width           =   2055
         End
         Begin FPSpread.vaSpread sprHak 
            Height          =   3645
            Left            =   210
            TabIndex        =   82
            Top             =   600
            Width           =   5025
            _Version        =   393216
            _ExtentX        =   8864
            _ExtentY        =   6429
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
            MaxCols         =   3
            SpreadDesigner  =   "STD200.frx":01D4
         End
         Begin VB.Label lblHakClose 
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
            Left            =   4440
            TabIndex        =   146
            Top             =   150
            Width           =   1035
         End
         Begin VB.Label Label32 
            Alignment       =   1  '������ ����
            BackStyle       =   0  '����
            Caption         =   "�б�"
            Height          =   210
            Left            =   -240
            TabIndex        =   145
            Top             =   180
            Width           =   975
         End
      End
   End
   Begin VB.Frame fraAddr 
      BackColor       =   &H00808080&
      BorderStyle     =   0  '����
      Height          =   4335
      Left            =   15870
      TabIndex        =   137
      Top             =   1740
      Width           =   7455
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  '����
         Height          =   4275
         Left            =   30
         TabIndex        =   138
         Top             =   30
         Width           =   7395
         Begin VB.TextBox txtFAddr 
            Height          =   345
            Left            =   900
            TabIndex        =   79
            Text            =   "txtFAddr"
            Top             =   90
            Width           =   2505
         End
         Begin FPSpread.vaSpread sprZip 
            Height          =   3645
            Left            =   120
            TabIndex        =   80
            Top             =   480
            Width           =   7155
            _Version        =   393216
            _ExtentX        =   12621
            _ExtentY        =   6429
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
            MaxCols         =   3
            SpreadDesigner  =   "STD200.frx":19E1
         End
         Begin VB.Label Label29 
            Alignment       =   1  '������ ����
            BackStyle       =   0  '����
            Caption         =   "�� �̸�"
            Height          =   210
            Left            =   -180
            TabIndex        =   140
            Top             =   150
            Width           =   975
         End
         Begin VB.Label lblZipClose 
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
            Left            =   6420
            TabIndex        =   139
            Top             =   150
            Width           =   1035
         End
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00C6AD84&
      BorderStyle     =   0  '����
      Caption         =   "Frame10"
      Height          =   10065
      Left            =   0
      TabIndex        =   93
      Top             =   0
      Width           =   8355
      Begin VB.Frame Frame9 
         BackColor       =   &H00F7EFE7&
         BorderStyle     =   0  '����
         Caption         =   "Frame9"
         Height          =   10005
         Left            =   30
         TabIndex        =   94
         Top             =   30
         Width           =   8295
         Begin VB.Frame Frame7 
            BackColor       =   &H0082C8E8&
            BorderStyle     =   0  '����
            Caption         =   "Frame13"
            Height          =   1605
            Left            =   30
            TabIndex        =   122
            Top             =   4980
            Width           =   8235
            Begin VB.Frame Frame8 
               BackColor       =   &H00F7EFE7&
               BorderStyle     =   0  '����
               Caption         =   ">> ��ȸŽ�� ���ð���"
               Height          =   1545
               Left            =   30
               TabIndex        =   123
               Top             =   30
               Width           =   8175
               Begin VB.CommandButton cmdPZip 
                  Caption         =   "��"
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
                  Left            =   1950
                  TabIndex        =   51
                  Top             =   870
                  Width           =   495
               End
               Begin VB.TextBox txtPJob 
                  Height          =   300
                  Left            =   5700
                  TabIndex        =   43
                  Text            =   "txtPJob"
                  Top             =   255
                  Width           =   1905
               End
               Begin VB.ComboBox cboPrtRel 
                  Height          =   300
                  Left            =   2820
                  Style           =   2  '��Ӵٿ� ���
                  TabIndex        =   42
                  Top             =   255
                  Width           =   1725
               End
               Begin VB.TextBox txtPrtNM 
                  Height          =   270
                  IMEMode         =   10  '�ѱ� 
                  Left            =   1080
                  TabIndex        =   41
                  Text            =   "txtPrtNM"
                  Top             =   270
                  Width           =   1725
               End
               Begin VB.TextBox txtPCel 
                  Height          =   270
                  IMEMode         =   10  '�ѱ� 
                  Index           =   2
                  Left            =   4830
                  MaxLength       =   4
                  TabIndex        =   49
                  Text            =   "txtP"
                  Top             =   570
                  Width           =   615
               End
               Begin VB.TextBox txtPCel 
                  Height          =   270
                  IMEMode         =   10  '�ѱ� 
                  Index           =   1
                  Left            =   4200
                  MaxLength       =   4
                  TabIndex        =   48
                  Text            =   "txtP"
                  Top             =   570
                  Width           =   615
               End
               Begin VB.TextBox txtPTel 
                  Height          =   270
                  IMEMode         =   10  '�ѱ� 
                  Index           =   2
                  Left            =   2340
                  MaxLength       =   4
                  TabIndex        =   46
                  Text            =   "9999"
                  Top             =   570
                  Width           =   615
               End
               Begin VB.TextBox txtPTel 
                  Height          =   270
                  IMEMode         =   10  '�ѱ� 
                  Index           =   1
                  Left            =   1710
                  MaxLength       =   4
                  TabIndex        =   45
                  Text            =   "9999"
                  Top             =   570
                  Width           =   615
               End
               Begin VB.TextBox txtPTel 
                  Height          =   270
                  IMEMode         =   10  '�ѱ� 
                  Index           =   0
                  Left            =   1080
                  MaxLength       =   4
                  TabIndex        =   44
                  Text            =   "9999"
                  Top             =   570
                  Width           =   615
               End
               Begin VB.TextBox txtPCel 
                  Height          =   270
                  IMEMode         =   10  '�ѱ� 
                  Index           =   0
                  Left            =   3570
                  MaxLength       =   4
                  TabIndex        =   47
                  Text            =   "txtP"
                  Top             =   570
                  Width           =   615
               End
               Begin VB.TextBox txtPAddr2 
                  Height          =   300
                  Left            =   3510
                  TabIndex        =   53
                  Text            =   "txtPAddr2"
                  Top             =   1170
                  Width           =   4605
               End
               Begin VB.TextBox txtPAddr1 
                  Height          =   300
                  Left            =   1080
                  TabIndex        =   52
                  Text            =   "txtPAddr1"
                  Top             =   1170
                  Width           =   2415
               End
               Begin EditLib.fpMask fpPZipCD 
                  Height          =   255
                  Left            =   1080
                  TabIndex        =   50
                  Top             =   870
                  Width           =   855
                  _Version        =   196608
                  _ExtentX        =   1508
                  _ExtentY        =   450
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
               Begin VB.Label Label16 
                  Alignment       =   1  '������ ����
                  BackStyle       =   0  '����
                  Caption         =   "����"
                  Height          =   210
                  Left            =   4710
                  TabIndex        =   131
                  Top             =   300
                  Width           =   975
               End
               Begin VB.Label Label15 
                  Alignment       =   1  '������ ����
                  BackStyle       =   0  '����
                  Caption         =   "�θ��"
                  Height          =   210
                  Left            =   0
                  TabIndex        =   129
                  Top             =   330
                  Width           =   975
               End
               Begin VB.Label Label14 
                  Alignment       =   1  '������ ����
                  BackStyle       =   0  '����
                  Caption         =   "TEL"
                  ForeColor       =   &H00000000&
                  Height          =   210
                  Left            =   -30
                  TabIndex        =   128
                  Top             =   600
                  Width           =   975
               End
               Begin VB.Label Label5 
                  Alignment       =   1  '������ ����
                  BackStyle       =   0  '����
                  Caption         =   "�ڵ���"
                  ForeColor       =   &H00000000&
                  Height          =   210
                  Left            =   2580
                  TabIndex        =   127
                  Top             =   600
                  Width           =   975
               End
               Begin VB.Label Label47 
                  Alignment       =   1  '������ ����
                  BackStyle       =   0  '����
                  Caption         =   "�����ּ�"
                  Height          =   210
                  Left            =   0
                  TabIndex        =   126
                  Top             =   1230
                  Width           =   975
               End
               Begin VB.Label Label46 
                  Alignment       =   1  '������ ����
                  BackStyle       =   0  '����
                  Caption         =   "������ȣ"
                  Height          =   210
                  Left            =   0
                  TabIndex        =   125
                  Top             =   900
                  Width           =   975
               End
               Begin VB.Label Label12 
                  BackStyle       =   0  '����
                  Caption         =   ">> �θ�"
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
                  TabIndex        =   124
                  Top             =   90
                  Width           =   2625
               End
            End
         End
         Begin VB.Frame Frame11 
            BackColor       =   &H00C6AD84&
            BorderStyle     =   0  '����
            Caption         =   "Frame11"
            Height          =   4875
            Left            =   30
            TabIndex        =   107
            Top             =   30
            Width           =   8235
            Begin VB.Frame Frame3 
               BackColor       =   &H00F7EFE7&
               BorderStyle     =   0  '����
               Caption         =   ">> �⺻�׸�"
               Height          =   4815
               Left            =   30
               TabIndex        =   108
               Top             =   30
               Width           =   8175
               Begin VB.Frame Frame14 
                  BackColor       =   &H00F7EFE7&
                  BorderStyle     =   0  '����
                  Height          =   435
                  Left            =   3030
                  TabIndex        =   148
                  Top             =   870
                  Width           =   2535
                  Begin VB.OptionButton optMU 
                     BackColor       =   &H00F7EFE7&
                     Caption         =   "������"
                     Height          =   285
                     Index           =   2
                     Left            =   1110
                     TabIndex        =   18
                     Top             =   90
                     Width           =   1215
                  End
                  Begin VB.OptionButton optMU 
                     BackColor       =   &H00F7EFE7&
                     Caption         =   "������"
                     Height          =   285
                     Index           =   1
                     Left            =   30
                     TabIndex        =   17
                     Top             =   90
                     Width           =   885
                  End
               End
               Begin VB.ComboBox cboKaeyol 
                  Height          =   300
                  Left            =   3450
                  Style           =   2  '��Ӵٿ� ���
                  TabIndex        =   15
                  Top             =   585
                  Width           =   1125
               End
               Begin VB.CommandButton cmdHak 
                  Caption         =   "��"
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
                  Left            =   1770
                  TabIndex        =   36
                  Top             =   4140
                  Width           =   495
               End
               Begin VB.Frame Frame1 
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   0  '����
                  Height          =   2205
                  Left            =   6180
                  TabIndex        =   142
                  Top             =   330
                  Width           =   1875
                  Begin VB.Image Photo 
                     Height          =   2145
                     Left            =   30
                     Stretch         =   -1  'True
                     Top             =   30
                     Width           =   1785
                  End
               End
               Begin VB.CommandButton cmdZip 
                  Caption         =   "��"
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
                  Left            =   1740
                  TabIndex        =   32
                  Top             =   3390
                  Width           =   495
               End
               Begin VB.TextBox txtHakNM 
                  Height          =   300
                  Left            =   870
                  TabIndex        =   37
                  Text            =   "txtHakNM"
                  Top             =   4440
                  Width           =   2295
               End
               Begin VB.TextBox txtHakCD 
                  Height          =   300
                  Left            =   870
                  TabIndex        =   35
                  Text            =   "txtHakCD"
                  Top             =   4110
                  Width           =   885
               End
               Begin VB.TextBox txtAddr2 
                  Height          =   300
                  Left            =   3270
                  TabIndex        =   34
                  Text            =   "txtAddr2"
                  Top             =   3690
                  Width           =   4605
               End
               Begin VB.TextBox txtAddr1 
                  Height          =   300
                  Left            =   870
                  TabIndex        =   33
                  Text            =   "txtAddr1"
                  Top             =   3690
                  Width           =   2415
               End
               Begin VB.ComboBox cboSex 
                  Height          =   300
                  Left            =   2250
                  Style           =   2  '��Ӵٿ� ���
                  TabIndex        =   22
                  Top             =   1980
                  Width           =   705
               End
               Begin VB.ComboBox cboGrd 
                  Height          =   300
                  Left            =   870
                  Style           =   2  '��Ӵٿ� ���
                  TabIndex        =   16
                  Top             =   930
                  Width           =   1245
               End
               Begin VB.TextBox txtEmail 
                  Height          =   315
                  Left            =   870
                  TabIndex        =   30
                  Text            =   "txtEmail"
                  Top             =   3000
                  Width           =   4605
               End
               Begin VB.TextBox txtAcc_No 
                  Height          =   315
                  IMEMode         =   10  '�ѱ� 
                  Left            =   4350
                  TabIndex        =   38
                  Text            =   "txtAcc_No"
                  Top             =   4110
                  Width           =   1815
               End
               Begin VB.TextBox txtCel 
                  Height          =   315
                  IMEMode         =   10  '�ѱ� 
                  Index           =   2
                  Left            =   2130
                  MaxLength       =   4
                  TabIndex        =   29
                  Text            =   "txtCel"
                  Top             =   2640
                  Width           =   615
               End
               Begin VB.TextBox txtCel 
                  Height          =   315
                  IMEMode         =   10  '�ѱ� 
                  Index           =   1
                  Left            =   1500
                  MaxLength       =   4
                  TabIndex        =   28
                  Text            =   "txtCel"
                  Top             =   2640
                  Width           =   615
               End
               Begin VB.TextBox txtTel 
                  Height          =   315
                  IMEMode         =   10  '�ѱ� 
                  Index           =   2
                  Left            =   2130
                  MaxLength       =   4
                  TabIndex        =   26
                  Text            =   "9999"
                  Top             =   2340
                  Width           =   615
               End
               Begin VB.TextBox txtTel 
                  Height          =   315
                  IMEMode         =   10  '�ѱ� 
                  Index           =   1
                  Left            =   1500
                  MaxLength       =   4
                  TabIndex        =   25
                  Text            =   "9999"
                  Top             =   2340
                  Width           =   615
               End
               Begin VB.TextBox txtSu_No 
                  Height          =   315
                  IMEMode         =   10  '�ѱ� 
                  Left            =   870
                  TabIndex        =   20
                  Text            =   "txtSu_No"
                  Top             =   1620
                  Width           =   1065
               End
               Begin VB.TextBox txtStdNM 
                  Height          =   315
                  IMEMode         =   10  '�ѱ� 
                  Left            =   870
                  TabIndex        =   21
                  Text            =   "txtStdNM"
                  Top             =   1980
                  Width           =   1365
               End
               Begin VB.TextBox txtOrd_No 
                  BackColor       =   &H00C0FFFF&
                  Enabled         =   0   'False
                  Height          =   315
                  Left            =   870
                  TabIndex        =   19
                  Text            =   "txtOrd_No"
                  Top             =   1260
                  Width           =   1065
               End
               Begin VB.TextBox txtTel 
                  Height          =   315
                  IMEMode         =   10  '�ѱ� 
                  Index           =   0
                  Left            =   870
                  MaxLength       =   4
                  TabIndex        =   24
                  Text            =   "9999"
                  Top             =   2340
                  Width           =   615
               End
               Begin VB.TextBox txtCel 
                  Height          =   315
                  IMEMode         =   10  '�ѱ� 
                  Index           =   0
                  Left            =   870
                  MaxLength       =   4
                  TabIndex        =   27
                  Text            =   "txtCel"
                  Top             =   2655
                  Width           =   615
               End
               Begin VB.TextBox txtRegDate 
                  Enabled         =   0   'False
                  Height          =   315
                  IMEMode         =   10  '�ѱ� 
                  Left            =   6390
                  TabIndex        =   40
                  Text            =   "txtRegDate"
                  Top             =   4440
                  Width           =   1455
               End
               Begin VB.CommandButton cmdNew 
                  Caption         =   "�ű� (&S)"
                  Height          =   315
                  Left            =   1410
                  TabIndex        =   12
                  Top             =   60
                  Width           =   1125
               End
               Begin EditLib.fpMask fpBirth 
                  Height          =   315
                  Left            =   4140
                  TabIndex        =   23
                  Top             =   1980
                  Width           =   1185
                  _Version        =   196608
                  _ExtentX        =   2090
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
               Begin VB.Frame Frame2 
                  BackColor       =   &H00F7EFE7&
                  BorderStyle     =   0  '����
                  Height          =   435
                  Left            =   210
                  TabIndex        =   120
                  Top             =   510
                  Width           =   2535
                  Begin VB.OptionButton optGaeyul 
                     BackColor       =   &H00F7EFE7&
                     Caption         =   "����"
                     Height          =   285
                     Index           =   1
                     Left            =   30
                     TabIndex        =   13
                     Top             =   90
                     Width           =   885
                  End
                  Begin VB.OptionButton optGaeyul 
                     BackColor       =   &H00F7EFE7&
                     Caption         =   "����Ŭ����"
                     Height          =   285
                     Index           =   2
                     Left            =   1110
                     TabIndex        =   14
                     Top             =   90
                     Width           =   1215
                  End
               End
               Begin EditLib.fpLongInteger fpAmnt 
                  Height          =   315
                  Left            =   4350
                  TabIndex        =   39
                  Top             =   4440
                  Width           =   1815
                  _Version        =   196608
                  _ExtentX        =   3201
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
                  MaxValue        =   "99999999"
                  MinValue        =   "0"
                  NegFormat       =   1
                  NegToggle       =   0   'False
                  Separator       =   ","
                  UseSeparator    =   -1  'True
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
               Begin EditLib.fpMask fpZip 
                  Height          =   255
                  Left            =   870
                  TabIndex        =   31
                  Top             =   3390
                  Width           =   855
                  _Version        =   196608
                  _ExtentX        =   1508
                  _ExtentY        =   450
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
               Begin VB.TextBox txtPhoto 
                  Alignment       =   1  '������ ����
                  BackColor       =   &H00F7EFE7&
                  BorderStyle     =   0  '����
                  Enabled         =   0   'False
                  Height          =   300
                  Left            =   2910
                  TabIndex        =   147
                  Text            =   "txtPhoto"
                  Top             =   60
                  Width           =   5085
               End
               Begin VB.Label Label22 
                  Alignment       =   1  '������ ����
                  BackStyle       =   0  '����
                  Caption         =   "�б�"
                  Height          =   210
                  Left            =   -210
                  TabIndex        =   136
                  Top             =   4140
                  Width           =   975
               End
               Begin VB.Label Label21 
                  Alignment       =   1  '������ ����
                  BackStyle       =   0  '����
                  Caption         =   "�����ּ�"
                  Height          =   210
                  Left            =   -180
                  TabIndex        =   135
                  Top             =   3750
                  Width           =   975
               End
               Begin VB.Label Label20 
                  Alignment       =   1  '������ ����
                  BackStyle       =   0  '����
                  Caption         =   "������ȣ"
                  Height          =   210
                  Left            =   -180
                  TabIndex        =   134
                  Top             =   3420
                  Width           =   975
               End
               Begin VB.Label Label18 
                  Alignment       =   1  '������ ����
                  BackStyle       =   0  '����
                  Caption         =   "����"
                  ForeColor       =   &H00000000&
                  Height          =   210
                  Left            =   3300
                  TabIndex        =   133
                  Top             =   4155
                  Width           =   975
               End
               Begin VB.Label Label17 
                  Alignment       =   1  '������ ����
                  BackStyle       =   0  '����
                  Caption         =   "�г�"
                  Height          =   210
                  Left            =   -180
                  TabIndex        =   132
                  Top             =   990
                  Width           =   975
               End
               Begin VB.Label Label51 
                  Alignment       =   1  '������ ����
                  BackStyle       =   0  '����
                  Caption         =   "�̸���"
                  Height          =   210
                  Left            =   -180
                  TabIndex        =   130
                  Top             =   3030
                  Width           =   975
               End
               Begin VB.Label Label3 
                  Alignment       =   1  '������ ����
                  BackStyle       =   0  '����
                  Caption         =   "�������"
                  Height          =   210
                  Left            =   3090
                  TabIndex        =   117
                  Top             =   2025
                  Width           =   975
               End
               Begin VB.Label Label2 
                  Alignment       =   1  '������ ����
                  BackStyle       =   0  '����
                  Caption         =   "�л���"
                  Height          =   210
                  Left            =   -180
                  TabIndex        =   116
                  Top             =   2040
                  Width           =   975
               End
               Begin VB.Label Label1 
                  Alignment       =   1  '������ ����
                  BackStyle       =   0  '����
                  Caption         =   "�����ȣ"
                  Height          =   210
                  Left            =   -180
                  TabIndex        =   115
                  Top             =   1680
                  Width           =   975
               End
               Begin VB.Label Label4 
                  Alignment       =   1  '������ ����
                  BackStyle       =   0  '����
                  Caption         =   "No"
                  Height          =   210
                  Left            =   -180
                  TabIndex        =   114
                  Top             =   1320
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
                  TabIndex        =   113
                  Top             =   60
                  Width           =   2625
               End
               Begin VB.Label Label28 
                  Alignment       =   1  '������ ����
                  BackStyle       =   0  '����
                  Caption         =   "�迭"
                  Height          =   210
                  Left            =   2700
                  TabIndex        =   112
                  Top             =   630
                  Width           =   705
               End
               Begin VB.Label Label39 
                  Alignment       =   1  '������ ����
                  BackStyle       =   0  '����
                  Caption         =   "TEL"
                  ForeColor       =   &H00000000&
                  Height          =   210
                  Left            =   -180
                  TabIndex        =   111
                  Top             =   2400
                  Width           =   975
               End
               Begin VB.Label Label41 
                  Alignment       =   1  '������ ����
                  BackStyle       =   0  '����
                  Caption         =   "�ڵ���"
                  ForeColor       =   &H00000000&
                  Height          =   210
                  Left            =   -180
                  TabIndex        =   110
                  Top             =   2700
                  Width           =   975
               End
               Begin VB.Label Label42 
                  BackStyle       =   0  '����
                  Caption         =   "�������"
                  ForeColor       =   &H00C000C0&
                  Height          =   315
                  Left            =   6390
                  TabIndex        =   109
                  Top             =   4200
                  Width           =   1365
               End
            End
         End
         Begin VB.Frame Frame12 
            BackColor       =   &H00C6AD84&
            BorderStyle     =   0  '����
            Caption         =   "Frame12"
            Height          =   885
            Left            =   30
            TabIndex        =   101
            Top             =   6660
            Width           =   8235
            Begin VB.Frame Frame4 
               BackColor       =   &H00F7EFE7&
               BorderStyle     =   0  '����
               Caption         =   ">> ����"
               Height          =   825
               Left            =   30
               TabIndex        =   102
               Top             =   30
               Width           =   8175
               Begin VB.ComboBox cboMu_Type 
                  Height          =   300
                  Left            =   3180
                  Style           =   2  '��Ӵٿ� ���
                  TabIndex        =   54
                  Top             =   60
                  Width           =   1845
               End
               Begin VB.ComboBox cboPTS1 
                  Height          =   300
                  Left            =   3150
                  Style           =   2  '��Ӵٿ� ���
                  TabIndex        =   56
                  Top             =   405
                  Width           =   795
               End
               Begin VB.TextBox txtEng 
                  Height          =   270
                  IMEMode         =   10  '�ѱ� 
                  Left            =   6450
                  TabIndex        =   58
                  Text            =   "txtEng"
                  Top             =   420
                  Width           =   1095
               End
               Begin VB.TextBox txtMat 
                  Height          =   270
                  IMEMode         =   10  '�ѱ� 
                  Left            =   4620
                  TabIndex        =   57
                  Text            =   "txtMat"
                  Top             =   420
                  Width           =   1095
               End
               Begin VB.TextBox txtKor 
                  Height          =   270
                  IMEMode         =   10  '�ѱ� 
                  Left            =   960
                  TabIndex        =   55
                  Text            =   "txtKor"
                  Top             =   420
                  Width           =   1095
               End
               Begin VB.Label Label33 
                  Alignment       =   1  '������ ����
                  BackStyle       =   0  '����
                  Caption         =   "���"
                  Height          =   345
                  Left            =   2640
                  TabIndex        =   155
                  Top             =   90
                  Width           =   435
               End
               Begin VB.Label Label19 
                  Alignment       =   1  '������ ����
                  BackStyle       =   0  '����
                  Caption         =   "��������"
                  Height          =   210
                  Left            =   2130
                  TabIndex        =   141
                  Top             =   450
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
                  Left            =   3600
                  TabIndex        =   106
                  Top             =   450
                  Width           =   975
               End
               Begin VB.Label Label7 
                  Alignment       =   1  '������ ����
                  BackStyle       =   0  '����
                  Caption         =   "�ܱ���"
                  Height          =   210
                  Left            =   5430
                  TabIndex        =   105
                  Top             =   450
                  Width           =   1005
               End
               Begin VB.Label Label6 
                  Alignment       =   1  '������ ����
                  BackStyle       =   0  '����
                  Caption         =   "���"
                  Height          =   210
                  Left            =   -120
                  TabIndex        =   104
                  Top             =   450
                  Width           =   975
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
                  Left            =   90
                  TabIndex        =   103
                  Top             =   60
                  Width           =   2625
               End
            End
         End
         Begin VB.Frame Frame13 
            BackColor       =   &H0082C8E8&
            BorderStyle     =   0  '����
            Caption         =   "Frame13"
            Height          =   705
            Left            =   30
            TabIndex        =   98
            Top             =   7620
            Width           =   8235
            Begin VB.Frame fraSEL1 
               BackColor       =   &H00F7EFE7&
               BorderStyle     =   0  '����
               Caption         =   ">> ��ȸŽ�� ���ð���"
               Height          =   645
               Left            =   30
               TabIndex        =   99
               Top             =   30
               Width           =   8175
               Begin VB.CheckBox chkSatam 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "��ȸ��ȭ"
                  Height          =   345
                  Index           =   10
                  Left            =   6090
                  TabIndex        =   68
                  Top             =   300
                  Width           =   1245
               End
               Begin VB.CheckBox chkSatam 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "����"
                  Height          =   345
                  Index           =   9
                  Left            =   4770
                  TabIndex        =   67
                  Top             =   300
                  Width           =   1245
               End
               Begin VB.CheckBox chkSatam 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "������ġ"
                  Height          =   345
                  Index           =   8
                  Left            =   3720
                  TabIndex        =   66
                  Top             =   300
                  Width           =   1245
               End
               Begin VB.CheckBox chkSatam 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "�����"
                  Height          =   345
                  Index           =   7
                  Left            =   2520
                  TabIndex        =   65
                  Top             =   300
                  Width           =   1245
               End
               Begin VB.CheckBox chkSatam 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "���ƽþƻ�"
                  Height          =   345
                  Index           =   6
                  Left            =   1260
                  TabIndex        =   64
                  Top             =   330
                  Width           =   1245
               End
               Begin VB.CheckBox chkSatam 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "��������"
                  Height          =   345
                  Index           =   5
                  Left            =   6090
                  TabIndex        =   63
                  Top             =   60
                  Width           =   1245
               End
               Begin VB.CheckBox chkSatam 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "�ѱ�����"
                  Height          =   345
                  Index           =   4
                  Left            =   4770
                  TabIndex        =   62
                  Top             =   60
                  Width           =   1245
               End
               Begin VB.CheckBox chkSatam 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "�ѱ���"
                  Height          =   345
                  Index           =   3
                  Left            =   3720
                  TabIndex        =   61
                  Top             =   60
                  Width           =   1245
               End
               Begin VB.CheckBox chkSatam 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "�����ͻ��"
                  Height          =   345
                  Index           =   2
                  Left            =   2520
                  TabIndex        =   60
                  Top             =   60
                  Width           =   1245
               End
               Begin VB.CheckBox chkSatam 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "��Ȱ������"
                  Height          =   345
                  Index           =   1
                  Left            =   1260
                  TabIndex        =   59
                  Top             =   60
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
                  TabIndex        =   100
                  Top             =   90
                  Width           =   2625
               End
            End
         End
         Begin VB.Frame Frame15 
            BackColor       =   &H0082C8E8&
            BorderStyle     =   0  '����
            Caption         =   "Frame15"
            Height          =   675
            Left            =   30
            TabIndex        =   95
            Top             =   8370
            Width           =   8235
            Begin VB.Frame fraSEL3 
               BackColor       =   &H00F7EFE7&
               BorderStyle     =   0  '����
               Caption         =   ">> ����Ž�� ���ð���"
               Height          =   615
               Left            =   30
               TabIndex        =   96
               Top             =   30
               Width           =   8175
               Begin VB.CheckBox chkGwatam 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "��������2"
                  Height          =   345
                  Index           =   8
                  Left            =   5340
                  TabIndex        =   76
                  Top             =   300
                  Width           =   1245
               End
               Begin VB.CheckBox chkGwatam 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "ȭ��2"
                  Height          =   345
                  Index           =   7
                  Left            =   3960
                  TabIndex        =   75
                  Top             =   300
                  Width           =   1245
               End
               Begin VB.CheckBox chkGwatam 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "����2"
                  Height          =   345
                  Index           =   6
                  Left            =   2640
                  TabIndex        =   74
                  Top             =   300
                  Width           =   1245
               End
               Begin VB.CheckBox chkGwatam 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "����2"
                  Height          =   345
                  Index           =   5
                  Left            =   1260
                  TabIndex        =   73
                  Top             =   300
                  Width           =   1245
               End
               Begin VB.CheckBox chkGwatam 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "��������1"
                  Height          =   345
                  Index           =   4
                  Left            =   5340
                  TabIndex        =   72
                  Top             =   30
                  Width           =   1245
               End
               Begin VB.CheckBox chkGwatam 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "����1"
                  Height          =   345
                  Index           =   2
                  Left            =   2640
                  TabIndex        =   71
                  Top             =   30
                  Width           =   1245
               End
               Begin VB.CheckBox chkGwatam 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "ȭ��1"
                  Height          =   345
                  Index           =   3
                  Left            =   3960
                  TabIndex        =   70
                  Top             =   30
                  Width           =   1245
               End
               Begin VB.CheckBox chkGwatam 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "����1"
                  Height          =   345
                  Index           =   1
                  Left            =   1260
                  TabIndex        =   69
                  Top             =   30
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
                  TabIndex        =   97
                  Top             =   90
                  Width           =   2625
               End
            End
         End
         Begin VB.CommandButton cmdStdin 
            Caption         =   "�л���� �� �����ϱ� (&S)"
            Height          =   450
            Left            =   2430
            TabIndex        =   77
            Top             =   9150
            Width           =   2655
         End
         Begin VB.CommandButton cmdStdDel 
            Caption         =   "�л������ϱ�"
            Height          =   450
            Left            =   5880
            TabIndex        =   78
            Top             =   9150
            Width           =   1365
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
            TabIndex        =   118
            Top             =   9690
            Width           =   4365
         End
      End
   End
   Begin VB.Frame Frame18 
      BackColor       =   &H0082C8E8&
      BorderStyle     =   0  '����
      Caption         =   "Frame18"
      Height          =   9465
      Left            =   8400
      TabIndex        =   86
      Top             =   30
      Width           =   6615
      Begin VB.Frame Frame19 
         BackColor       =   &H00D2EAF5&
         BorderStyle     =   0  '����
         Caption         =   "Frame19"
         Height          =   9405
         Left            =   30
         TabIndex        =   87
         Top             =   30
         Width           =   6555
         Begin VB.TextBox Text1 
            Height          =   1035
            Left            =   780
            TabIndex        =   154
            Text            =   "Text1"
            Top             =   2670
            Visible         =   0   'False
            Width           =   4185
         End
         Begin VB.ComboBox cbo_gbn 
            Height          =   300
            Left            =   5625
            TabIndex        =   153
            Top             =   390
            Width           =   900
         End
         Begin VB.ComboBox cboMU 
            Height          =   300
            Left            =   3390
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   7
            Top             =   690
            Width           =   1365
         End
         Begin VB.TextBox txtSu_No2 
            Height          =   285
            IMEMode         =   10  '�ѱ� 
            Left            =   2100
            TabIndex        =   4
            Text            =   "txtSu_No2"
            Top             =   690
            Width           =   735
         End
         Begin VB.TextBox txtSu_No1 
            Height          =   285
            IMEMode         =   10  '�ѱ� 
            Left            =   900
            TabIndex        =   3
            Text            =   "txtSu_No1"
            Top             =   690
            Width           =   735
         End
         Begin VB.TextBox txtStdNM_F 
            Height          =   285
            IMEMode         =   10  '�ѱ� 
            Left            =   900
            TabIndex        =   8
            Text            =   "txtStdNM_F"
            Top             =   1035
            Width           =   735
         End
         Begin VB.CommandButton cmdFind 
            Caption         =   "��ȸ�ϱ�(&F)"
            Height          =   390
            Left            =   4500
            TabIndex        =   10
            Top             =   960
            Width           =   1305
         End
         Begin VB.ComboBox cboKaeyol_F 
            Height          =   300
            Left            =   3390
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   2
            Top             =   390
            Width           =   1365
         End
         Begin EditLib.fpMask fpBirth_F 
            Height          =   285
            Left            =   2520
            TabIndex        =   9
            Top             =   1035
            Width           =   1215
            _Version        =   196608
            _ExtentX        =   2143
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
         Begin FPSpread.vaSpread sprSTD_F 
            Height          =   7965
            Left            =   30
            TabIndex        =   11
            Top             =   1440
            Width           =   6495
            _Version        =   393216
            _ExtentX        =   11456
            _ExtentY        =   14049
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
            SpreadDesigner  =   "STD200.frx":3211
         End
         Begin VB.Frame Frame6 
            BackColor       =   &H00D2EAF5&
            BorderStyle     =   0  '����
            Height          =   435
            Left            =   120
            TabIndex        =   121
            Top             =   300
            Width           =   2535
            Begin VB.OptionButton optGaeyul_F 
               BackColor       =   &H00D2EAF5&
               Caption         =   "����Ŭ����"
               Height          =   285
               Index           =   2
               Left            =   1110
               TabIndex        =   1
               Top             =   90
               Width           =   1215
            End
            Begin VB.OptionButton optGaeyul_F 
               BackColor       =   &H00D2EAF5&
               Caption         =   "����"
               Height          =   285
               Index           =   1
               Left            =   30
               TabIndex        =   0
               Top             =   90
               Width           =   885
            End
         End
         Begin InetCtlsObjects.Inet Inet1 
            Left            =   0
            Top             =   0
            _ExtentX        =   1005
            _ExtentY        =   1005
            _Version        =   393216
         End
         Begin VB.Frame Frame17 
            BackColor       =   &H00D2EAF5&
            BorderStyle     =   0  '����
            Height          =   435
            Left            =   2970
            TabIndex        =   149
            Top             =   -90
            Visible         =   0   'False
            Width           =   2535
            Begin VB.OptionButton optMU_F 
               BackColor       =   &H00D2EAF5&
               Caption         =   "������"
               Enabled         =   0   'False
               Height          =   285
               Index           =   1
               Left            =   30
               TabIndex        =   5
               Top             =   90
               Width           =   885
            End
            Begin VB.OptionButton optMU_F 
               BackColor       =   &H00D2EAF5&
               Caption         =   "������"
               Enabled         =   0   'False
               Height          =   285
               Index           =   2
               Left            =   1110
               TabIndex        =   6
               Top             =   90
               Width           =   1215
            End
         End
         Begin VB.Label �гⱸ�� 
            Caption         =   "�гⱸ��"
            Height          =   300
            Left            =   4800
            TabIndex        =   152
            Top             =   450
            Width           =   720
         End
         Begin VB.Label Label30 
            Alignment       =   1  '������ ����
            BackStyle       =   0  '����
            Caption         =   "�迭"
            Height          =   210
            Left            =   0
            TabIndex        =   151
            Top             =   -165
            Width           =   1035
         End
         Begin VB.Label Label23 
            Alignment       =   1  '������ ����
            BackStyle       =   0  '����
            Caption         =   "����"
            Height          =   210
            Left            =   2310
            TabIndex        =   150
            Top             =   735
            Width           =   1035
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
            TabIndex        =   92
            Top             =   90
            Width           =   2625
         End
         Begin VB.Label Label25 
            BackStyle       =   0  '����
            Caption         =   "�����ȣ             ����"
            Height          =   210
            Left            =   180
            TabIndex        =   91
            Top             =   750
            Width           =   2025
         End
         Begin VB.Label Label26 
            Alignment       =   1  '������ ����
            BackStyle       =   0  '����
            Caption         =   "�л���"
            Height          =   210
            Left            =   -150
            TabIndex        =   90
            Top             =   1095
            Width           =   975
         End
         Begin VB.Label Label27 
            Alignment       =   1  '������ ����
            BackStyle       =   0  '����
            Caption         =   "�������"
            Height          =   210
            Left            =   1530
            TabIndex        =   89
            Top             =   1095
            Width           =   975
         End
         Begin VB.Label Label31 
            Alignment       =   1  '������ ����
            BackStyle       =   0  '����
            Caption         =   "�迭"
            Height          =   210
            Left            =   2310
            TabIndex        =   88
            Top             =   435
            Width           =   1035
         End
         Begin VB.Image imgExcel 
            Height          =   420
            Left            =   6030
            Picture         =   "STD200.frx":4C56
            Stretch         =   -1  'True
            Top             =   930
            Width           =   390
         End
      End
   End
   Begin VB.Frame fraGwamok 
      BackColor       =   &H0082C8E8&
      BorderStyle     =   0  '����
      Caption         =   "����"
      Height          =   4275
      Left            =   2100
      TabIndex        =   83
      Top             =   9840
      Width           =   8865
      Begin VB.Frame Frame23 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  '����
         Height          =   4215
         Left            =   30
         TabIndex        =   84
         Top             =   30
         Width           =   8805
         Begin VB.CommandButton cmdClose 
            Caption         =   "�ݱ�"
            Height          =   330
            Left            =   8160
            TabIndex        =   85
            Top             =   3840
            Width           =   585
         End
         Begin VB.Image Image1 
            Height          =   4080
            Left            =   30
            Picture         =   "STD200.frx":5097
            Top             =   60
            Width           =   8730
         End
      End
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   150
      Top             =   9720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   132
      ImageHeight     =   150
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "STD200.frx":C761
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgFile 
      Left            =   0
      Top             =   10230
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "STD200"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################################################
'   �� �� ��  �� : �뼺�п� ���л���, �ݹ��� & �ð�ǥ ���α׷�
'   ����ý��۸� :
'   ��   ��   �� : STD200
'   �� ��  �� �� : ���� & ��������Ŭ����
'
'   ��   ��   �� : 2009/12/10
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
    ORD_NO            As String
    ACACD             As String
    EXMROUND          As String
    EMAIL             As String
    USERNM            As String
    SU_NO             As String
    HOPE_ACACD        As String
    SEX               As String
    KEYOL             As String
    Birth             As String
    SEL1              As String
    SEL2              As String
    SEL3              As String
    SEL4              As String
    PTS_SEL           As String
    PTS1              As String
    PTS2              As String
    GRADE_KOR         As String
    GRADE_MAT         As String
    GRADE_ENG         As String
    ZIPCODE           As String
    ADDR2             As String
    ADDR              As String
    TEL1              As String
    TEL2              As String
    TEL3              As String
    CEL1              As String
    CEL2              As String
    CEL3              As String
    HAKCD             As String
    GYEAR             As String
    D_UNIVCD          As String
    D_MAJORCD         As String
    FILENM            As String
    PRTNM             As String
    PRTREL            As String
    PZIPCODE          As String
    PADDR2            As String
    PADDR             As String
    PJOB              As String
    PTEL1             As String
    PTEL2             As String
    PTEL3             As String
    JTEL1             As String
    JTEL2             As String
    JTEL3             As String
    REG_DATE          As String
    BIGO              As String
    ACC_NO            As String
    AMNT              As String
    SEL5              As String
    MOD_REG_DATE      As String
    RECSMS            As String
    GRADE_TAM1        As String
    GRADE_TAM2        As String
    GRADE_TAM1_SELECT As String
    GRADE_TAM2_SELECT As String
    AGREE_DM          As String
    AGREE_DS          As String
End Type
Private uExcel_StaData          As tExcel_StdData


Private sSavePath       As String       '<< image ���
Private smSavePath      As String       '<< image ��� (����Ŭ����)

Private Sub Form_Terminate()
    Unload Me
End Sub










'>> ���� & ����Ŭ���п� ���� �迭ó��
Private Sub optGaeyul_Click(Index As Integer)
    
    Select Case Index
        Case 1      '< ����
            cboKaeyol.Clear
            cboKaeyol.AddItem "�ι�" & Space(50) & "1"
            cboKaeyol.AddItem "�ڿ�" & Space(50) & "2"
            cboKaeyol.ListIndex = 0
            
            optGaeyul_F(1).value = True
            
        Case 2      '< ����Ŭ����
            cboKaeyol.Clear
            cboKaeyol.AddItem "�ι�" & Space(50) & "11"
            cboKaeyol.AddItem "�ڿ�" & Space(50) & "12"
            cboKaeyol.ListIndex = 0
            
            optGaeyul_F(2).value = True
            
    End Select
    
End Sub


Private Sub optGaeyul_F_Click(Index As Integer)
    Select Case Index
        Case 1      '< ����
            cboKaeyol_F.Clear
            cboKaeyol_F.AddItem "��ü" & Space(50) & "X"
            cboKaeyol_F.AddItem "�ι�" & Space(50) & "1"
            cboKaeyol_F.AddItem "�ڿ�" & Space(50) & "2"
            cboKaeyol_F.ListIndex = 0
            
        Case 2      '< ����Ŭ����
            cboKaeyol_F.Clear
            cboKaeyol_F.AddItem "��ü" & Space(50) & "X"
            cboKaeyol_F.AddItem "�ι�" & Space(50) & "11"
            cboKaeyol_F.AddItem "�ڿ�" & Space(50) & "12"
            cboKaeyol_F.ListIndex = 0
            
    End Select
    
End Sub


Private Sub Form_Load()
    
    Me.Move 0, 0, 15255, 10620
    fraGwamok.Visible = False
    
    sSavePath = App.Path & "\PHOTO"
    If Dir(sSavePath, vbDirectory) = "" Then
        Call MkDir(sSavePath)
    End If
    smSavePath = App.Path & "\MPHOTO"
    If Dir(smSavePath, vbDirectory) = "" Then
        Call MkDir(smSavePath)
    End If
    
    With sprSTD_F
    .ShadowColor = basModule.ShadowColor1
    .ShadowDark = basModule.ShadowDark1
    .ShadowText = basModule.ShadowText1
    .GridColor = basModule.GridColor1
    .GrayAreaBackColor = basModule.GrayAreaBackColor1
    End With
    
    optGaeyul(1).value = True
    
    With cboGrd
        .Clear
        .AddItem "���" & Space(30) & "4"
        .AddItem "3�г�" & Space(30) & "3"
        .AddItem "2�г�" & Space(30) & "2"
        .AddItem "1�г�" & Space(30) & "1"
        
        .ListIndex = 0
    End With
    
    With cboSex
        .Clear
        .AddItem "��" & Space(30) & "1"
        .AddItem "��" & Space(30) & "2"
        
        .ListIndex = 0
    End With
    
    
    With cboPTS1
        .Clear
        .AddItem "����" & Space(50) & "X"
        .AddItem "��" & Space(30) & "1"
        .AddItem "��" & Space(30) & "2"
        
        .ListIndex = 0
    End With
    
    With cboPrtRel
        .Clear
        .AddItem "��" & Space(30) & "1"
        .AddItem "��" & Space(30) & "2"
        
        .ListIndex = 0
    End With
    
    With cboMU
        .Clear
        .AddItem "��ü" & Space(30) & "X"
        .AddItem "������" & Space(30) & "1"
        .AddItem "������" & Space(30) & "2"
        
        .ListIndex = 0
    End With
    
    Select Case Trim(SchCD)
        Case "S"
            With cbo_gbn
                .Enabled = True
                .Clear
                .AddItem "��3"
                .AddItem "���"
            End With
            
            With �гⱸ��
                .Visible = True
            End With
        
        Case Else
            With cbo_gbn
                .Enabled = True
                .Visible = False
            End With
            
            With �гⱸ��
                .Visible = False
            End With
    End Select
    
    '���
    With cboMu_type
        .Clear
        
        .AddItem "2013 ����" & Space(30) & "1"   '����
        .AddItem "6�� �򰡿�" & Space(30) & "2"
        .AddItem "9�� �򰡿�" & Space(30) & "3"
        .AddItem "��2 �뼺���ǰ���" & Space(30) & "4"
        .AddItem "��2 ����û���ǰ���" & Space(30) & "5"
        If basModule.SchCD = "N" Then .AddItem "���ŵ��" & Space(30) & "9"
        .AddItem "����" & Space(30) & "X"
        
        .Enabled = True
        
        .ListIndex = .ListCount - 1
        
    End With
    
    
    optMU(1).value = True
    optMU(2).value = False
    
    fraGwamok.Visible = False
    fraAddr.Visible = False:        fraAddr.Tag = "S"
    fraHak.Visible = False
    
    cmdNew_Click
    
End Sub


'���
Private Sub Set_Mu_type(ByVal val As Integer)
    Select Case val
        Case "1"
            cboMu_type.ListIndex = 0 '2013 ����
        Case "2"
            cboMu_type.ListIndex = 1 '6�� ����
        Case "3"
            cboMu_type.ListIndex = 2 '9�� ����
        Case "4"
            cboMu_type.ListIndex = 3 '9�� ����
        Case "5"
            cboMu_type.ListIndex = 4 '9�� ����
        Case "9"
            cboMu_type.ListIndex = 5 '���ŵ��
        Case Else
            cboMu_type.ListIndex = cboMu_type.ListCount - 1
    End Select
    
End Sub


'----------- ��ȸȭ�� --------------------------------

Private Sub cmdGwamokView_Click()
    fraGwamok.Left = 60
    fraGwamok.Top = 3390
    fraGwamok.ZOrder 0
    
    fraGwamok.Visible = True
End Sub

Private Sub cmdClose_Click()
    fraGwamok.Visible = False
End Sub

Private Sub lblHakClose_Click()
    fraHak.Visible = False
End Sub

Private Sub lblZipClose_Click()
    fraAddr.Visible = False
End Sub

Private Sub cmdZip_Click()
    fraAddr.Tag = "S"
    fraAddr.Top = cmdZip.Top + 150
    fraAddr.Left = cmdZip.Left + 90
    fraAddr.Visible = True
    
    txtFAddr.Text = ""
    txtFAddr.SetFocus
    
End Sub

Private Sub cmdPZip_Click()
    fraAddr.Tag = "P"
    fraAddr.Top = cmdPZip.Top + 1000
    fraAddr.Left = cmdPZip.Left + 90
    fraAddr.Visible = True
    
    txtFAddr.Text = ""
    txtFAddr.SetFocus
    
End Sub

Private Sub cmdHak_Click()
    fraHak.Top = cmdHak.Top + 130
    fraHak.Left = cmdHak.Left + 90
    fraHak.Visible = True
    
    txtFHak.Text = ""
    txtFHak.SetFocus

End Sub







'>> �ű�
Private Sub cmdNew_Click()
    Dim ni      As Integer
    
    '======== 1 =================
    
    Set Photo.Picture = imgList.ListImages.Item(1).Picture      '<< �⺻����
    
    txtOrd_No.Text = ""
    txtSu_No.Text = ""
    txtStdNM.Text = ""
    
    fpBirth.Text = ""
    
    txtTel(0).Text = ""
    txtTel(1).Text = ""
    txtTel(2).Text = ""
    
    txtCel(0).Text = ""
    txtCel(1).Text = ""
    txtCel(2).Text = ""
    
    txtEmail.Text = ""
    
    fpZip.Text = ""
    txtAddr1.Text = ""
    txtAddr2.Text = ""
    
    txtHakCD.Text = ""
    txtHakNM.Text = ""
    txtAcc_No.Text = ""
    fpAmnt.value = 0
    
    txtRegDate.Text = ""
    txtPhoto.Text = ""
    
    '======== 2 =================
    
    txtPrtNM.Text = ""
    txtPJob.Text = ""
    
    txtPTel(0).Text = ""
    txtPTel(1).Text = ""
    txtPTel(2).Text = ""
    
    txtPCel(0).Text = ""
    txtPCel(1).Text = ""
    txtPCel(2).Text = ""
    
    fpPZipCD.Text = ""
    txtPAddr1.Text = ""
    txtPAddr2.Text = ""
    
    '======== 3 =================
    txtKor.Text = ""
    txtMat.Text = "":       cboPTS1.ListIndex = 0
    txtEng.Text = ""
    cboMu_type.ListIndex = cboMu_type.ListCount - 1

    '��Ž
    For ni = 1 To SATAM_COUNT Step 1
        chkSatam(ni).value = 0
    Next ni
    
    '��Ž
    For ni = 1 To 8 Step 1
        chkGwatam(ni).value = 0
    Next ni
    
    '======== ��ȸâ =================
    
    txtSu_No1.Text = ""
    txtSu_No2.Text = ""
    txtStdNM_F.Text = ""
    fpBirth_F.Text = ""
    
    sprZip.MaxRows = 0
    sprHak.MaxRows = 0
    sprSTD_F.MaxRows = 0
    
End Sub




'=����ϱ� ========================================================================================
Private Sub cmdStdin_Click()
    Dim bRet        As Boolean
    
    If Trim(txtStdNM.Text) = "" Then
        MsgBox "�л����� ����Ͻʽÿ�.", vbExclamation + vbOKOnly, "�л����"
        Exit Sub
    End If
    
    If Trim(fpBirth.UnFmtText) = "" Then
        MsgBox "��������� ����Ͻʽÿ�.", vbExclamation + vbOKOnly, "�л����"
        Exit Sub
    End If
    
    On Error GoTo ErrStmt
    
    cmdStdin.Enabled = False
    
        bRet = Save_Stdin
            
    cmdStdin.Enabled = True
    
    If bRet = True Then
        MsgBox "�л� ����Ͽ����ϴ�.", vbInformation + vbOKOnly, "�л����"
        
    Else
        MsgBox "�л� ��Ͻ� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "�л����"
    
    End If
    
    Exit Sub
ErrStmt:
    MsgBox "�л���Ͻ� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "�л����"
    On Error GoTo 0
    
    cmdStdin.Enabled = True
    
End Sub



'>> �л����
Private Function Save_Stdin() As Boolean
    Dim bRet        As Boolean
    
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    Dim sStr        As String
    
    Dim ni          As Long
    
    Dim nLength     As Byte
    Dim sTmp        As String
    Dim nTmp        As Double
    Dim nExe        As Integer
    
    Dim sOrd_No     As String
    Dim sExmRoundX  As String
    
    bRet = False
    
    On Error Resume Next
    
    basDataBase.DBConn.BeginTrans
    
    Set DBCmd = New ADODB.Command
    Set DBRec = New ADODB.Recordset
    Set DBParam = New ADODB.Parameter
    
    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
    
    '    ������
    '    1 �����ϴ��߿� �������� �׳� ���� On Error Resume Next...
    '    2 ������ �ɰ�� nExe���� 1�̹Ƿ� ����� ���� �ʴ´�.
    '    3. nExe���� 1�� �ƴҰ�� update�����̹Ƿ� Insert�Ѵ�.
    '    ��ġ�ڴ�.. �̷�����..
    
    
    
    '>>> UPDATE
    sStr = ""
    sStr = sStr & " UPDATE HWSIN01TB_WINTER "
    sStr = sStr & "    SET "
    
'    sStr = sStr & "        EXMROUND = "
'        Select Case Trim(SchCD)
'        Case "N"
'            sStr = sStr & "           'NR081126E1',"
'        Case "K"
'            sStr = sStr & "           'KN081126E1',"
'        Case "S"
'            sStr = sStr & "           'SP081126E1',"
'        Case "P"
'            sStr = sStr & "           'MK081126E1',"
'        Case "M"
'            'sStr = sStr & "           'NR081126E1',"
'
'        Case "W"
'            sStr = sStr & "           'KN081126E1',"
'        Case "Q"
'            sStr = sStr & "           'KN081126E1',"
'
'        Case "J"
'            sStr = sStr & "           'YJ081126E1',"
'        Case "B"
'            sStr = sStr & "           'BS081126E1',"
'
'        Case Else
'            sStr = sStr & "           'BS081126E1',"
'    End Select

    sStr = sStr & "        BIGO        = '" & Trim(Right(cboGrd.Text, 2)) & "',"          '�г�
    
    If Trim(Right(cboGrd.Text, 1)) = "1" Then
        sStr = sStr & "        KEYOL       = '',"                                               '�迭       1�г⸸
    Else
        sStr = sStr & "        KEYOL       = '" & Trim(Right(cboKaeyol.Text, 2)) & "',"       '�迭
    End If
    
    sStr = sStr & "        USERNM      = '" & Trim(txtStdNM.Text) & "',"                  '�л���
    sStr = sStr & "        SU_NO       = '" & Trim(txtSu_No.Text) & "',"                  '�����ȣ
                                       
    sStr = sStr & "        SEX         = '" & Trim(Right(cboSex.Text, 2)) & "',"          '��/��
    sStr = sStr & "        Birth       = '" & Trim(fpBirth.UnFmtText) & "',"              '�������
                                       
    sStr = sStr & "        TEL1        = '" & Trim(txtTel(0).Text) & "',"
    sStr = sStr & "        TEL2        = '" & Trim(txtTel(1).Text) & "',"
    sStr = sStr & "        TEL3        = '" & Trim(txtTel(2).Text) & "',"
                                       
    sStr = sStr & "        CEL1        = '" & Trim(txtCel(0).Text) & "',"
    sStr = sStr & "        CEL2        = '" & Trim(txtCel(1).Text) & "',"
    sStr = sStr & "        CEL3        = '" & Trim(txtCel(2).Text) & "',"
                                       
    sStr = sStr & "        EMAIL       = '" & Trim(txtEmail.Text) & "',"                  '�̸���
                                       
    '�ּ�
    sStr = sStr & "        ZIPCODE     = '" & Trim(fpZip.UnFmtText) & "',"
    sStr = sStr & "        ADDR2       = '" & Trim(txtAddr1.Text) & "',"
    sStr = sStr & "        ADDR        = '" & Trim(txtAddr2.Text) & "',"

    '>> ��Ž���� ����
        sTmp = ""
        For ni = 1 To SATAM_COUNT Step 1
            If chkSatam(ni).value = 1 Then
                sTmp = sTmp & Format(ni, "0") & "|"
            End If
        Next ni
    sStr = sStr & "        SEL1        = '" & Trim(sTmp) & "',"


    '>> ��Ž���� ����
        sTmp = ""
        For ni = 1 To 8 Step 1
            If chkGwatam(ni).value = 1 Then
                sTmp = sTmp & Format(ni, "0") & "|"
            End If
        Next ni
    sStr = sStr & "        SEL4        = '" & Trim(sTmp) & "',"
    
    '>> �����κ�
    sStr = sStr & "        GRADE_KOR   = '" & Trim(txtKor.Text) & "',"
    
    If Trim(Right(cboPTS1.Text, 1)) = "X" Then
        sTmp = ""
    Else
        sTmp = Trim(Right(cboPTS1.Text, 1))
    End If
    sStr = sStr & "        PTS1        = '" & Trim(sTmp) & "',"
    sStr = sStr & "        GRADE_MAT   = '" & Trim(txtMat.Text) & "',"
    sStr = sStr & "        GRADE_ENG   = '" & Trim(txtEng.Text) & "',"
    
    sStr = sStr & "        HAKCD       = '" & Trim(txtHakCD.Text) & "',"                  '�б�
    sStr = sStr & "        FILENM      = '" & Trim(txtPhoto.Text) & "',"                  '�������
    
    sStr = sStr & "        PRTNM       = '" & Trim(txtPrtNM.Text) & "',"                  '�кθ�
    sStr = sStr & "        PRTREL      = '" & Trim(Right(cboPrtRel.Text, 1)) & "',"       '����
    '>> �ּ�
    sStr = sStr & "        PZIPCODE    = '" & Trim(fpPZipCD.UnFmtText) & "',"
    sStr = sStr & "        PADDR2      = '" & Trim(txtPAddr1.Text) & "',"
    sStr = sStr & "        PADDR       = '" & Trim(txtPAddr2.Text) & "',"
                                       
    sStr = sStr & "        PJOB        = '" & Trim(txtPJob.Text) & "',"                   '����
    
    '��ȭ��ȣ
    sStr = sStr & "        PTEL1       = '" & Trim(txtPTel(0).Text) & "',"
    sStr = sStr & "        PTEL2       = '" & Trim(txtPTel(1).Text) & "',"
    sStr = sStr & "        PTEL3       = '" & Trim(txtPTel(2).Text) & "',"
    '�޴���
    sStr = sStr & "        JTEL1       = '" & Trim(txtPCel(0).Text) & "',"
    sStr = sStr & "        JTEL2       = '" & Trim(txtPCel(1).Text) & "',"
    sStr = sStr & "        JTEL3       = '" & Trim(txtPCel(2).Text) & "',"
    
    sStr = sStr & "        ACC_NO      = '" & Trim(txtAcc_No.Text) & "',"
    sStr = sStr & "        AMNT        = '" & Trim(fpAmnt.value) & "',"
        
    '���
    sTmp = ""
    If Trim(Right(cboMu_type.Text, 30)) <> "X" Then sTmp = Trim(Right(cboMu_type.Text, 30))
    sStr = sStr & "        ETC1        = '" & sTmp & "'"

    sStr = sStr & "  WHERE ORD_NO      = '" & Trim(txtOrd_No.Text) & "'"
    sStr = sStr & "    AND ACACD       = '" & Trim(basModule.SchCD) & "'"
    
    nExe = 0

    DBCmd.CommandText = sStr
    DBCmd.CommandType = adCmdText
    DBCmd.CommandTimeout = 30

    DBCmd.Execute nExe, , -1

    Do While basDataBase.DBConn.State And adStateExecuting
        DoEvents
    Loop

    If nExe = 1 Then
        bRet = True     '����
        
    ElseIf nExe = 0 Then
        
        On Error GoTo ErrStmt

            sStr = ""
            sStr = " SELECT ORD_NO_SEQ.NEXTVAL AS ORDNO FROM DUAL"
            
            DBCmd.CommandText = sStr
            DBCmd.CommandType = adCmdText
            DBCmd.CommandTimeout = 30

            DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic���·� ���ԵǸ� record count�� �� �� ����.
            Do While DBRec.State And adStateExecuting
                DoEvents
            Loop
            
            If DBRec.RecordCount = 1 Then
                sOrd_No = DBRec.Fields("ORDNO")
                txtOrd_No.Text = sOrd_No
            Else
                txtOrd_No.Text = ""
                GoTo ErrStmt
            End If
            
            If optMU(1).value = True Then
                sExmRoundX = "1"
            ElseIf optMU(2).value = True Then
                sExmRoundX = "2"
            End If
            
            
    '>>> INSERT
            sStr = ""
            sStr = sStr & " INSERT INTO HWSIN01TB_WINTER ("
            sStr = sStr & "     ORD_NO      , ACACD       , EXMROUND    ,"
            sStr = sStr & "     BIGO        , KEYOL       ,"
            sStr = sStr & "     USERNM      , SU_NO       ,"
            sStr = sStr & "     SEX         ,"
            sStr = sStr & "     Birth       ,"
            sStr = sStr & "     TEL1        , TEL2        , TEL3        ,"
            sStr = sStr & "     CEL1        , CEL2        , CEL3        ,"
            sStr = sStr & "     EMAIL       ,"
            sStr = sStr & "     ZIPCODE     , ADDR2       , ADDR        ,"
            sStr = sStr & " "
            sStr = sStr & "     SEL1        , SEL4        ,"
            sStr = sStr & " "
            sStr = sStr & "     GRADE_KOR   ,"
            sStr = sStr & "     PTS1        ,"
            sStr = sStr & "     GRADE_MAT   ,"
            sStr = sStr & "     GRADE_ENG   ,"
            sStr = sStr & " "
            sStr = sStr & "     HAKCD       ,"
            sStr = sStr & "     FILENM      ,"
            sStr = sStr & " "
            sStr = sStr & "     PRTNM       , PRTREL      ,"
            sStr = sStr & "     PZIPCODE    , PADDR2      , PADDR       , PJOB        ,"
            sStr = sStr & "     PTEL1       , PTEL2       , PTEL3       ,"
            sStr = sStr & "     JTEL1       , JTEL2       , JTEL3       ,"
            sStr = sStr & "     REG_DATE    ,"
            sStr = sStr & " "
            sStr = sStr & "     ACC_NO      ,"
            sStr = sStr & "     AMNT        ,"
            sStr = sStr & "     AGREE_DM    ,"
            sStr = sStr & "     AGREE_DS    , ETC1"
            sStr = sStr & " ) VALUES ("
            sStr = sStr & "                 " & sOrd_No & ","
            
            sStr = sStr & "                 '" & Trim(basModule.SchCD) & "',"
            
            Select Case Trim(SchCD)
                Case "N"
                    sStr = sStr & "         'NR081126E" & sExmRoundX & "',"
                Case "K"
                    sStr = sStr & "         'KN081126E" & sExmRoundX & "',"
                Case "S"
                    sStr = sStr & "         'SP081126E" & sExmRoundX & "',"
                Case "P"
                    sStr = sStr & "         'MK081126E" & sExmRoundX & "',"
                Case "M"
                    sStr = sStr & "         'NR081126E1'," '<---------- ���� �ּ��Ǿ��־ insert�� value������ �����ʴ�.. �� �ּ�?!!
                    
                Case "W"
                    sStr = sStr & "         'KN081126E" & sExmRoundX & "',"
                Case "Q"
                    sStr = sStr & "         'KN081126E" & sExmRoundX & "',"
                    
                Case "J"
                    sStr = sStr & "         'YJ081126E" & sExmRoundX & "',"
                Case "B"
                    sStr = sStr & "         'BS081126E" & sExmRoundX & "',"
                
                Case Else
                    sStr = sStr & "         'BS081126E" & sExmRoundX & "',"
                    
            End Select
            sStr = sStr & "'" & Trim(Right(cboGrd.Text, 2)) & "',"          '�г�
            
            If Trim(Right(cboGrd.Text, 1)) = "1" Then
                sStr = sStr & "'',"                         '�迭
            Else
                sStr = sStr & "'" & Trim(Right(cboKaeyol.Text, 2)) & "',"       '�迭
            End If
            
            sStr = sStr & "'" & Trim(txtStdNM.Text) & "',"                  '�л���
            sStr = sStr & "'" & Trim(txtSu_No.Text) & "',"                  '�����ȣ
            
            sStr = sStr & "'" & Trim(Right(cboSex.Text, 2)) & "',"          '��/��
            sStr = sStr & "'" & Trim(fpBirth.UnFmtText) & "',"              '�������
            
            '��ȭ��ȣ
            sStr = sStr & "'" & Trim(txtTel(0).Text) & "',"
            sStr = sStr & "'" & Trim(txtTel(1).Text) & "',"
            sStr = sStr & "'" & Trim(txtTel(2).Text) & "',"
            '�޴���
            sStr = sStr & "'" & Trim(txtCel(0).Text) & "',"
            sStr = sStr & "'" & Trim(txtCel(1).Text) & "',"
            sStr = sStr & "'" & Trim(txtCel(2).Text) & "',"
            sStr = sStr & "'" & Trim(txtEmail.Text) & "',"                  '�̸���
            
            '�ּ�
            sStr = sStr & "'" & Trim(fpZip.UnFmtText) & "',"
            sStr = sStr & "'" & Trim(txtAddr1.Text) & "',"
            sStr = sStr & "'" & Trim(txtAddr2.Text) & "',"
            
            '>> ��Ž���� ����
                sTmp = ""
                For ni = 1 To SATAM_COUNT Step 1
                    If chkSatam(ni).value = 1 Then
                        sTmp = sTmp & Format(ni, "0") & "|"
                    End If
                Next ni
            sStr = sStr & "'" & Trim(sTmp) & "',"
        
        
            '>> ��Ž���� ����
                sTmp = ""
                For ni = 1 To 8 Step 1
                    If chkGwatam(ni).value = 1 Then
                        sTmp = sTmp & Format(ni, "0") & "|"
                    End If
                Next ni
            sStr = sStr & "'" & Trim(sTmp) & "',"
            
            '>> �����κ�
            sStr = sStr & "'" & Trim(txtKor.Text) & "',"
            If Trim(Right(cboPTS1.Text, 1)) = "X" Then
                sStr = sStr & "'', "
            Else
                sStr = sStr & "'" & Trim(Right(cboPTS1.Text, 1)) & "',"
            End If
            sStr = sStr & "'" & Trim(txtMat.Text) & "',"
            sStr = sStr & "'" & Trim(txtEng.Text) & "',"
            
            sStr = sStr & "'" & Trim(txtHakCD.Text) & "',"                  '�б�
            If Trim(txtPhoto.Text) = "" Then                                '�������
                sStr = sStr & "'',"
            Else
                sStr = sStr & "'" & Trim(txtPhoto.Text) & "',"
            End If
            
            sStr = sStr & "'" & Trim(txtPrtNM.Text) & "',"                  '�кθ�
            sStr = sStr & "'" & Trim(Right(cboPrtRel.Text, 1)) & "',"       '����
            '>> �ּ�
            sStr = sStr & "'" & Trim(fpPZipCD.UnFmtText) & "',"
            sStr = sStr & "'" & Trim(txtPAddr1.Text) & "',"
            sStr = sStr & "'" & Trim(txtPAddr2.Text) & "',"
            
            sStr = sStr & "'" & Trim(txtPJob.Text) & "',"                   '����
            
            '��ȭ��ȣ
            sStr = sStr & "'" & Trim(txtPTel(0).Text) & "',"
            sStr = sStr & "'" & Trim(txtPTel(1).Text) & "',"
            sStr = sStr & "'" & Trim(txtPTel(2).Text) & "',"
            '�޴���
            sStr = sStr & "'" & Trim(txtPCel(0).Text) & "',"
            sStr = sStr & "'" & Trim(txtPCel(1).Text) & "',"
            sStr = sStr & "'" & Trim(txtPCel(2).Text) & "',"
            
            sStr = sStr & " SYSDATE, "
            
            sStr = sStr & "'" & Trim(txtAcc_No.Text) & "',"
            sStr = sStr & "'" & Trim(fpAmnt.value) & "',"
            
            sStr = sStr & "'Y','Y' , "
            
            
            sTmp = ""
            If Trim(Right(cboMu_type.Text, 30)) <> "X" Then sTmp = Trim(Right(cboMu_type.Text, 30))
            sStr = sStr & "'" & sTmp & "'"                                   '���
            
            sStr = sStr & " )"
    
    
            nExe = 0
            
            'Text1.Text = sStr

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
            
            Set DBRec = Nothing

    End If
    



    basDataBase.DBConn.CommitTrans

    Save_Stdin = bRet

    Set DBCmd = Nothing
    Set DBParam = Nothing
    
    Exit Function
    
ErrStmt:
    
    basDataBase.DBConn.RollbackTrans
    
    Set DBCmd = Nothing
    Set DBParam = Nothing
    
    Save_Stdin = bRet
    
End Function













'= ���� =========================================================================================
Private Sub cmdStdDel_Click()
    
    Dim bRet        As Boolean
    Dim sTmp        As String
    
    '>> üũ����
    If Trim(txtOrd_No.Text) = "" Then
        MsgBox "No�� �����ϴ�.", vbExclamation + vbOKOnly, "�л�����"
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
        MsgBox "�л� ������ ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "�л�����"
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
    
    Dim sStr        As String
    Dim nLength     As Byte
    Dim sTmp        As String
    Dim nTmp        As Double
    
    Dim nExe        As Long
    
    bRet = False
    
    On Error GoTo ErrStmt
    
    basDataBase.DBConn.BeginTrans
    
    Set DBCmd = New ADODB.Command
    Set DBParam = New ADODB.Parameter
    
    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
    
    
    sStr = ""
    sStr = sStr & " DELETE FROM HWSIN01TB_WINTER "
    sStr = sStr & "  WHERE ORD_NO = " & Trim(txtOrd_No.Text)
    
    nExe = 0

    DBCmd.CommandText = sStr
    DBCmd.CommandType = adCmdText
    DBCmd.CommandTimeout = 30

    DBCmd.Execute nExe, , -1

    Do While basDataBase.DBConn.State And adStateExecuting
        DoEvents
    Loop

    If nExe = 1 Then
        bRet = True     '����
    End If
    
    Set DBCmd = Nothing
    Set DBParam = Nothing
    
    basDataBase.DBConn.CommitTrans
    
    Delete_StdOut = bRet
    Exit Function
    
ErrStmt:
    basDataBase.DBConn.RollbackTrans
    
    Set DBCmd = Nothing
    Set DBParam = Nothing
    
    Delete_StdOut = bRet
End Function









'= ��ȸ�ϱ� ========================================================================================


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
    Dim nTmp        As Long
    
    Dim sGbn        As String
    Dim sKaeyol     As String
    
    On Error GoTo ErrStmt
    
    cmdFind.Enabled = False
    
    sprSTD_F.MaxRows = 0
    
    sStr = ""
    sStr = sStr & " SELECT ORD_NO, SU_NO, USERNM, DECODE(LENGTH(KEYOL),1,'����',2,'����C','����') AS WNT,"
    sStr = sStr & "        DECODE(KEYOL, '1', '�ι�', '11', '�ι�',"
    sStr = sStr & "                      '2', '�ڿ�', '12', '�ڿ�') AS KEYOL,"
    sStr = sStr & "        SUBSTR(Birth,1,4)||'-'||SUBSTR(Birth,5,2) ||'-'|| SUBSTR(Birth,7,2) AS Birth,"
    sStr = sStr & "        GRADE_KOR KOR, GRADE_MAT MAT, GRADE_ENG ENG"
    sStr = sStr & "   From HWSIN01TB_WINTER"
    sStr = sStr & "  WHERE ACACD = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "    AND EXMROUND LIKE "
        Select Case Trim(SchCD)
        Case "N"
            sStr = sStr & "           'NR081126%' "
        Case "K"
            sStr = sStr & "           'KN081126%' "
        Case "S"
            sStr = sStr & "           'SP081126%' "
        Case "P"
            sStr = sStr & "           'MK081126%' "
        Case "M"
            sStr = sStr & "           'NR081126E1' "   '<< ��ȸ�� �������� �ּ��Ǿ��ִ°� Ǯ����.. ����.. // ������ 2012.10.25

        Case "W"
            sStr = sStr & "           'KN081126%' "
        Case "Q"
            sStr = sStr & "           'KN081126%' "

        Case "J"
            sStr = sStr & "           'YJ081126%' "
        Case "B"
            sStr = sStr & "           'BS081126%' "

        Case Else
            sStr = sStr & "           'BS081126%' "
    End Select
    
    If Trim(Right(cboMU.Text, 1)) = "1" Then
        sStr = sStr & " AND SUBSTR(EXMROUND, LENGTH(EXMROUND))  = '1' "
    ElseIf Trim(Right(cboMU.Text, 1)) = "2" Then
        sStr = sStr & " AND SUBSTR(EXMROUND, LENGTH(EXMROUND))  = '2' "
    End If
    
    If optGaeyul_F(1).value = True Then
        sStr = sStr & " AND ( KEYOL IN ('1','2') or KEYOL IS NULL )"
    ElseIf optGaeyul_F(2).value = True Then
        sStr = sStr & " AND KEYOL IN ('11','12')"
    End If
    
    If Trim(txtSu_No1.Text) > " " And Trim(txtSu_No2.Text) > " " Then
        sStr = sStr & " AND SU_NO BETWEEN '" & Trim(txtSu_No1.Text) & "' " & _
                                     "AND '" & Trim(txtSu_No2.Text) & "'"
    End If
    If Trim(txtStdNM_F.Text) > " " Then
        sStr = sStr & " AND USERNM LIKE '%" & Trim(txtStdNM_F.Text) & "%'"
    End If
    If Trim(fpBirth_F.UnFmtText) > " " Then
        sStr = sStr & " AND Birth LIKE '%" & Trim(fpBirth_F.UnFmtText) & "%'"
    End If
    
    If Trim(Right(cboKaeyol_F.Text, 2)) <> "X" Then
        sStr = sStr & " AND KEYOL = '" & Trim(Right(cboKaeyol_F.Text, 3)) & "'"
    End If
    
    If Trim(cbo_gbn.Text = "��3") Then
        sStr = sStr & " AND BIGO = 3"
    ElseIf Trim(cbo_gbn.Text = "���") Then
        sStr = sStr & " AND BIGO = 4"
    End If
    
'    Text1.Text = sStr
    
    Set DBCmd = New ADODB.Command
    Set DBRec = New ADODB.Recordset
    Set DBParam = New ADODB.Parameter
    
    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
    DBCmd.CommandText = sStr
    DBCmd.CommandType = adCmdText
    DBCmd.CommandTimeout = 30
    'Text1.Text = sStr


    
    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic���·� ���ԵǸ� record count�� �� �� ����.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop
    
    With DBRec
        If .RecordCount > 0 Then
            .MoveFirst
            
            For nRec = 1 To .RecordCount Step 1
                sprSTD_F.MaxRows = sprSTD_F.MaxRows + 1
                sprSTD_F.Row = sprSTD_F.MaxRows
                
                
                sprSTD_F.Col = 1
                    sTmp = " ":   If IsNull(.Fields("ORD_NO")) = False Then sTmp = Trim(.Fields("ORD_NO"))
                        Call basFunction.Set_SprType_Text(sprSTD_F, "CENTER", "LEFT", LenB(sTmp), sTmp)
                        
                sprSTD_F.Col = sprSTD_F.Col + 1
                    sTmp = " ":  If IsNull(.Fields("SU_NO")) = False Then sTmp = Trim(.Fields("SU_NO"))
                        Call basFunction.Set_SprType_Text(sprSTD_F, "CENTER", "LEFT", LenB(sTmp), sTmp)
                        
                sprSTD_F.Col = sprSTD_F.Col + 1
                    sTmp = " ":  If IsNull(.Fields("USERNM")) = False Then sTmp = Trim(.Fields("USERNM"))
                        Call basFunction.Set_SprType_Text(sprSTD_F, "CENTER", "LEFT", LenB(sTmp), sTmp)
                
                sprSTD_F.Col = sprSTD_F.Col + 1
                    sTmp = " ":  If IsNull(.Fields("WNT")) = False Then sTmp = Trim(.Fields("WNT"))
                        Call basFunction.Set_SprType_Text(sprSTD_F, "CENTER", "LEFT", LenB(sTmp), sTmp)
                        
                sprSTD_F.Col = sprSTD_F.Col + 1
                    sTmp = " ":  If IsNull(.Fields("KEYOL")) = False Then sTmp = Trim(.Fields("KEYOL"))
                        Call basFunction.Set_SprType_Text(sprSTD_F, "CENTER", "LEFT", LenB(sTmp), sTmp)
                        
                sprSTD_F.Col = sprSTD_F.Col + 1
                    sTmp = " ":  If IsNull(.Fields("Birth")) = False Then sTmp = Trim(.Fields("Birth"))
                        Call basFunction.Set_SprType_Text(sprSTD_F, "CENTER", "LEFT", LenB(sTmp), sTmp)
                    
                        
                sprSTD_F.Col = sprSTD_F.Col + 1
                    sTmp = " ":  If IsNull(.Fields("KOR")) = False Then sTmp = Trim(.Fields("KOR"))
                        Call basFunction.Set_SprType_Text(sprSTD_F, "CENTER", "LEFT", LenB(sTmp), sTmp)
                sprSTD_F.Col = sprSTD_F.Col + 1
                    sTmp = " ":  If IsNull(.Fields("MAT")) = False Then sTmp = Trim(.Fields("MAT"))
                        Call basFunction.Set_SprType_Text(sprSTD_F, "CENTER", "LEFT", LenB(sTmp), sTmp)
                sprSTD_F.Col = sprSTD_F.Col + 1
                    sTmp = " ":  If IsNull(.Fields("ENG")) = False Then sTmp = Trim(.Fields("ENG"))
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
    
End Sub

'>> �����л� �����ֱ�
Private Sub Show_Select_STD(ByVal aOrdNO As String)
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
    sStr = sStr & "  SELECT "
        sStr = sStr & "     ORD_NO      , ACACD       , EXMROUND    ,"
        sStr = sStr & "     BIGO        , KEYOL       ,"
        sStr = sStr & "     USERNM      , SU_NO       ,"
        sStr = sStr & "     SEX         ,"
        sStr = sStr & "     Birth       ,"
        sStr = sStr & "     TEL1        , TEL2        , TEL3        ,"
        sStr = sStr & "     CEL1        , CEL2        , CEL3        ,"
        sStr = sStr & "     EMAIL       ,"
        sStr = sStr & "     ZIPCODE     , ADDR2       , ADDR        ,"
        sStr = sStr & " "
        sStr = sStr & "     SEL1        , SEL4        ,"
        sStr = sStr & " "
        sStr = sStr & "     GRADE_KOR AS KOR,"
        sStr = sStr & "     PTS1        ,"
        sStr = sStr & "     GRADE_MAT AS MAT,"
        sStr = sStr & "     GRADE_ENG AS ENG,"
        sStr = sStr & " "
        sStr = sStr & "     HAKCD       , GET_SCHOOLNM(HAKCD) AS HAKNM, "
        sStr = sStr & "     FILENM      ,"
        sStr = sStr & " "
        sStr = sStr & "     PRTNM       , PRTREL      ,"
        sStr = sStr & "     PZIPCODE    , PADDR2      , PADDR       , PJOB        ,"
        sStr = sStr & "     PTEL1       , PTEL2       , PTEL3       ,"
        sStr = sStr & "     JTEL1       , JTEL2       , JTEL3       ,"
        sStr = sStr & "     TO_CHAR(REG_DATE,'YYYY-MM-DD') AS REGDATE,"
        sStr = sStr & " "
        sStr = sStr & "     ACC_NO      ,"
        sStr = sStr & "     AMNT        , ETC1       "
    sStr = sStr & "    From HWSIN01TB_WINTER "
    sStr = sStr & "   WHERE EXMROUND LIKE "
    Select Case Trim(SchCD)
        Case "N"
            sStr = sStr & "            'NR081126%' "
        Case "K"
            sStr = sStr & "            'KN081126%' "
        Case "S"
            sStr = sStr & "            'SP081126%' "
        Case "P"
            sStr = sStr & "            'MK081126%' "
        Case "M"
            sStr = sStr & "            'NR081126%' "
            
        Case "W"
            sStr = sStr & "            'KN081126%' "
        Case "Q"
            sStr = sStr & "            'KN081126%' "
            
        Case "J"
            sStr = sStr & "            'YJ081126%' "
        Case "B"
            sStr = sStr & "            'BS081126%' "
        
        Case Else
            sStr = sStr & "            'BS081126%' "
    End Select
            
    sStr = sStr & "     AND ORD_NO = '" & Trim(aOrdNO) & "'"
    sStr = sStr & "     AND ACACD  = '" & Trim(basModule.SchCD) & "'"
    
    Set DBCmd = New ADODB.Command
    Set DBRec = New ADODB.Recordset
    Set DBParam = New ADODB.Parameter
    
    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
    DBCmd.CommandText = sStr
    DBCmd.CommandType = adCmdText
    DBCmd.CommandTimeout = 30
    
 
 
        
    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic���·� ���ԵǸ� record count�� �� �� ����.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop
    
    With DBRec
        If .RecordCount <> 1 Then
            MsgBox "��ȸ�� �л��� �����ϴ�.", vbExclamation + vbOKOnly, "�л���ȸ"
        Else
            .MoveFirst
            
            If IsNull(.Fields("KEYOL")) = True Then
                optGaeyul(1).value = True
                cboKaeyol.ListIndex = 0
            Else
                If Trim(.Fields("KEYOL")) = "1" Then
                    optGaeyul(1).value = True
                    cboKaeyol.ListIndex = 0
                    
                ElseIf Trim(.Fields("KEYOL")) = "11" Then
                    optGaeyul(2).value = True
                    cboKaeyol.ListIndex = 0
                    
                ElseIf Trim(.Fields("KEYOL")) = "2" Then
                    optGaeyul(1).value = True
                    cboKaeyol.ListIndex = 1
                    
                ElseIf Trim(.Fields("KEYOL")) = "12" Then
                    optGaeyul(2).value = True
                    cboKaeyol.ListIndex = 1
                
                End If
            End If
            
            Select Case Trim(.Fields("BIGO"))
                Case "1"
                    cboGrd.ListIndex = 3
                Case "2"
                    cboGrd.ListIndex = 2
                Case "3"
                    cboGrd.ListIndex = 1
                Case "4"
                    cboGrd.ListIndex = 0
            End Select
            txtOrd_No.Text = "":        If IsNull(.Fields("ORD_NO")) = False Then txtOrd_No.Text = Trim(.Fields("ORD_NO"))
            txtSu_No.Text = "":         If IsNull(.Fields("SU_NO")) = False Then txtSu_No.Text = Trim(.Fields("SU_NO"))
            txtStdNM.Text = "":         If IsNull(.Fields("USERNM")) = False Then txtStdNM.Text = Trim(.Fields("USERNM"))
            Select Case Trim(.Fields("SEX"))
                Case "1"
                    cboSex.ListIndex = 0
                Case "2"
                    cboSex.ListIndex = 1
            End Select
            
            If IsNull(.Fields("EXMROUND")) = False Then
                If Right(Trim(.Fields("EXMROUND")), 1) = "1" Then
                    optMU(1).value = True
                ElseIf Right(Trim(.Fields("EXMROUND")), 1) = "2" Then
                    optMU(2).value = True
                End If
            End If
            
            fpBirth.Text = "":          If IsNull(.Fields("Birth")) = False Then fpBirth.Text = Trim(.Fields("Birth"))
            
            txtTel(0).Text = "":        If IsNull(.Fields("TEL1")) = False Then txtTel(0).Text = Trim(.Fields("TEL1"))
            txtTel(1).Text = "":        If IsNull(.Fields("TEL2")) = False Then txtTel(1).Text = Trim(.Fields("TEL2"))
            txtTel(2).Text = "":        If IsNull(.Fields("TEL3")) = False Then txtTel(2).Text = Trim(.Fields("TEL3"))
            
            txtCel(0).Text = "":        If IsNull(.Fields("CEL1")) = False Then txtCel(0).Text = Trim(.Fields("CEL1"))
            txtCel(1).Text = "":        If IsNull(.Fields("CEL2")) = False Then txtCel(1).Text = Trim(.Fields("CEL2"))
            txtCel(2).Text = "":        If IsNull(.Fields("CEL3")) = False Then txtCel(2).Text = Trim(.Fields("CEL3"))
            
            txtEmail.Text = "":         If IsNull(.Fields("EMAIL")) = False Then txtEmail.Text = Trim(.Fields("EMAIL"))
            
            fpZip.Text = "":            If IsNull(.Fields("ZIPCODE")) = False Then fpZip.Text = Trim(.Fields("ZIPCODE"))
            txtAddr1.Text = "":         If IsNull(.Fields("ADDR2")) = False Then txtAddr1.Text = Trim(.Fields("ADDR2"))
            txtAddr2.Text = "":         If IsNull(.Fields("ADDR")) = False Then txtAddr2.Text = Trim(.Fields("ADDR"))
            
            txtHakCD.Text = "":         If IsNull(.Fields("HAKCD")) = False Then txtHakCD.Text = Trim(.Fields("HAKCD"))
            txtHakNM.Text = "":         If IsNull(.Fields("HAKNM")) = False Then txtHakNM.Text = Trim(.Fields("HAKNM"))
            
            txtPhoto.Text = "":         If IsNull(.Fields("FILENM")) = False Then txtPhoto.Text = Trim(.Fields("FILENM"))
            
            txtAcc_No.Text = "":        If IsNull(.Fields("ACC_NO")) = False Then txtAcc_No.Text = Trim(.Fields("ACC_NO"))
            fpAmnt.value = 0:           If IsNull(.Fields("AMNT")) = False Then fpAmnt = CDbl(.Fields("AMNT"))
            
            txtRegDate.Text = "":       If IsNull(.Fields("REGDATE")) = False Then txtRegDate.Text = Trim(.Fields("REGDATE"))
            
            
            txtPrtNM.Text = "":         If IsNull(.Fields("PRTNM")) = False Then txtPrtNM.Text = Trim(.Fields("PRTNM"))
            Select Case Trim(.Fields("PRTREL"))
                Case "1"
                    cboPrtRel.ListIndex = 0
                Case "2"
                    cboPrtRel.ListIndex = 1
            End Select
            txtPJob.Text = "":          If IsNull(.Fields("PJOB")) = False Then txtPJob.Text = Trim(.Fields("PJOB"))
            
            txtPTel(0).Text = "":       If IsNull(.Fields("PTEL1")) = False Then txtPTel(0).Text = Trim(.Fields("PTEL1"))
            txtPTel(1).Text = "":       If IsNull(.Fields("PTEL2")) = False Then txtPTel(1).Text = Trim(.Fields("PTEL2"))
            txtPTel(2).Text = "":       If IsNull(.Fields("PTEL3")) = False Then txtPTel(2).Text = Trim(.Fields("PTEL3"))
            
            txtPCel(0).Text = "":       If IsNull(.Fields("JTEL1")) = False Then txtPCel(0).Text = Trim(.Fields("JTEL1"))
            txtPCel(1).Text = "":       If IsNull(.Fields("JTEL2")) = False Then txtPCel(1).Text = Trim(.Fields("JTEL2"))
            txtPCel(2).Text = "":       If IsNull(.Fields("JTEL3")) = False Then txtPCel(2).Text = Trim(.Fields("JTEL3"))
            
            fpPZipCD.Text = "":         If IsNull(.Fields("PZIPCODE")) = False Then fpPZipCD.Text = Trim(.Fields("PZIPCODE"))
            txtPAddr1.Text = "":        If IsNull(.Fields("PADDR2")) = False Then txtPAddr1.Text = Trim(.Fields("PADDR2"))
            txtPAddr2.Text = "":        If IsNull(.Fields("PADDR")) = False Then txtPAddr2.Text = Trim(.Fields("PADDR"))
            
            txtKor.Text = "":           If IsNull(.Fields("KOR")) = False Then txtKor.Text = Trim(.Fields("KOR"))
            txtMat.Text = "":           If IsNull(.Fields("MAT")) = False Then txtMat.Text = Trim(.Fields("MAT"))
            txtEng.Text = "":           If IsNull(.Fields("ENG")) = False Then txtEng.Text = Trim(.Fields("ENG"))
            
            
            If IsNull(.Fields("PTS1")) = True Then
                cboPTS1.ListIndex = 0
            Else
                Select Case Trim(.Fields("PTS1"))
                    Case "1"
                        cboPTS1.ListIndex = 1
                    Case "2"
                        cboPTS1.ListIndex = 2
                End Select
            End If
            
        '���ɵ��
        If IsNull(.Fields("ETC1")) = True Then
            cboMu_type.ListIndex = cboMu_type.ListCount - 1
        Else
            Call Set_Mu_type(.Fields("ETC1"))
        End If
            
        '## ���ð���
            '>> ��Ž
            For ni = 1 To SATAM_COUNT Step 1
                chkSatam(ni).value = 0
            Next ni
            If IsNull(.Fields("SEL1")) = False Then
                sTmp = Trim(.Fields("SEL1"))
                sDiv = Split(sTmp, "|", -1, vbTextCompare)
                
                For ni = 0 To UBound(sDiv) - 1 Step 1
                    chkSatam(CInt(sDiv(ni))).value = 1
                Next ni
            End If
            
            '>> ��Ž
            For ni = 1 To 8 Step 1
                chkGwatam(ni).value = 0
            Next ni
            If IsNull(.Fields("SEL4")) = False Then
                sTmp = Trim(.Fields("SEL4"))
                sDiv = Split(sTmp, "|", -1, vbTextCompare)
                
                For ni = 0 To UBound(sDiv) - 1 Step 1
                    chkGwatam(CInt(sDiv(ni))).value = 1
                Next ni
            End If
            
            
            If Trim(txtPhoto.Text) = "" Then
                Set Photo.Picture = imgList.ListImages.Item(1).Picture
                
            ElseIf Trim(txtPhoto.Text) > " " Then
                
                '2010.12.20 �뷮��,����,������ ��쿡�� ���� ������ �����ȣ�� ����, �� �ܿ��� �ֹ� ��ȣ�� ����. ���ѿ�
                Select Case Trim(SchCD)
                    Case "N" '�뷮��
                        Call Get_STD_image(txtSu_No.Text, txtPhoto.Text)               '<< �̹��� �ڷ� ��������
                        
                        If optGaeyul(1).value = True Then
                            Set Photo.Picture = CheckJPG(sSavePath & "\" & txtSu_No.Text & ".jpg")
                        ElseIf optGaeyul(2).value = True Then
                            Set Photo.Picture = CheckJPG(smSavePath & "\" & txtSu_No.Text & ".jpg")
                        End If
                    Case "S" '����
                        Call Get_STD_image(txtSu_No.Text, txtPhoto.Text)               '<< �̹��� �ڷ� ��������
                        
                        If optGaeyul(1).value = True Then
                            Set Photo.Picture = CheckJPG(sSavePath & "\" & txtSu_No.Text & ".jpg")
                        ElseIf optGaeyul(2).value = True Then
                            Set Photo.Picture = CheckJPG(smSavePath & "\" & txtSu_No.Text & ".jpg")
                        End If
                    Case "J" '��ȭ
                        Call Get_STD_image(txtSu_No.Text, txtPhoto.Text)               '<< �̹��� �ڷ� ��������
                        
                        If optGaeyul(1).value = True Then
                            Set Photo.Picture = CheckJPG(sSavePath & "\" & txtSu_No.Text & ".jpg")
                        ElseIf optGaeyul(2).value = True Then
                            Set Photo.Picture = CheckJPG(smSavePath & "\" & txtSu_No.Text & ".jpg")
                        End If
                    Case Else '�� ��
                        Call Get_STD_image(txtOrd_No.Text, txtPhoto.Text)               '<< �̹��� �ڷ� ��������
                        
                        If optGaeyul(1).value = True Then
                            Set Photo.Picture = CheckJPG(sSavePath & "\" & txtOrd_No.Text & ".jpg")
                        ElseIf optGaeyul(2).value = True Then
                            Set Photo.Picture = CheckJPG(smSavePath & "\" & txtOrd_No.Text & ".jpg")
                        End If
                End Select
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

'>> �̹��� �������� üũ : üũ�� �̻��� �ִ� ��쿣 default ���� ������.
Public Function CheckJPG(fileName As String) As Picture

    Dim header(2)     As Byte
    Dim tailer(2)     As Byte
    Dim f             As Integer
    Dim MaxSize       As Long


    On Error Resume Next

    f = FreeFile()
    Open fileName For Binary As #f

        On Error GoTo 0
        If Err <> 0 Then
            Set CheckJPG = imgList.ListImages.Item(1).Picture
            Exit Function
        End If

        On Error Resume Next
        MaxSize = LOF(f)                                        '<< ������ ����Ʈ ũ�⸦ ���մϴ�.
        Get #f, 1, header()
        Get #f, MaxSize - 1, tailer()
    Close f

    ' Must start with hex FF D8  and end data hex FF D9
'    If (header(0) = 255 And header(1) = 216) And _
'       (tailer(0) = 255 And tailer(1) >= 209) Then
'        Set CheckJPG = LoadPicture(fileName)
'    Else
'        Set CheckJPG = imgList.ListImages.Item(1).Picture       '<< no-image
'    End If
    
    Set CheckJPG = LoadPicture(fileName)

End Function

'## ������ �̹��� ��������
Private Sub Get_STD_image(ByVal aOrdNO As String, ByVal aPhoto As String)
    
    Dim bData()     As Byte
    Dim f           As Integer
    Dim nRec        As Long

    Dim sLocalFile  As String
    Dim sSourceUrl  As String

    On Error Resume Next

    f = FreeFile()
    
    If optGaeyul(1).value = True Then
        sLocalFile = sSavePath & "\" & aOrdNO & ".jpg"                      '<< unique key : �л��ڵ�
    ElseIf optGaeyul(2).value = True Then
        sLocalFile = smSavePath & "\" & aOrdNO & ".jpg"                      '<< unique key : �л��ڵ�
    End If
    
    If Dir(sLocalFile, vbNormal) = "" Then                                                '<< �л� �̹��� ���� �͸� ����
        sSourceUrl = "http://www.dshw.co.kr" & aPhoto                   '<< ������ �̹��� ���
        
        bData() = Inet1.OpenURL(sSourceUrl, icByteArray)
        
        If UBound(bData) > 0 Then
            Open sLocalFile For Binary Access Write As #f
            Put #f, , bData()
        
            DoEvents
            Close #f
        End If
        
    End If
        
End Sub









'= ����ó�� ========================================================================================
Private Sub imgExcel_Click()
    
    On Error GoTo ErrStmt
    
    imgExcel.Enabled = False
        Call Get_Excel_Data
        
    imgExcel.Enabled = True
    
    Exit Sub
ErrStmt:
    MsgBox "�����ڷ� �������� �� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "�л� �����ڷ� ��������"
    On Error GoTo 0
    
End Sub






'## ��ü�л� ������ �ޱ�
Private Sub Get_Excel_Data()
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
    
    sStr = ""
    sStr = sStr & "  SELECT "
        
        '2010.12.20 �뷮�� ���� ���� ��� ���� ���� No�� �����ȣ�� ����
        Select Case Trim(SchCD)
            Case "N"
                sStr = sStr & "     SU_NO AS NO     , "
            Case "J"
                sStr = sStr & "     SU_NO AS NO     , "
            Case "S"
                sStr = sStr & "     SU_NO AS NO     , "
            Case Else
                sStr = sStr & "     ORD_NO AS NO     , "
        End Select
        
        sStr = sStr & "     DECODE(BIGO,'1','1 �г�','2', '2 �г�','3', '3 �г�', '4', '���') AS �г�, "
        sStr = sStr & "     DECODE(LENGTH(KEYOL),1,'����','����Ŭ����') AS ����,"
        sStr = sStr & "     DECODE(KEYOL,'1','�ι�','2','�ڿ�','11','�ι�','12','�ڿ�') AS �迭,"
        sStr = sStr & "     DECODE(SUBSTR(EXMROUND,LENGTH(EXMROUND)),'1','������','2','������') AS ����,"
        sStr = sStr & "     USERNM AS �л���, SU_NO AS �����ȣ      ,"
        sStr = sStr & "     DECODE(SEX,'1','��','2','��') AS ����,"
        sStr = sStr & "     SUBSTR(Birth,1,4)||'-'||SUBSTR(Birth,5,2) ||'-'||SUBSTR(Birth,7,2) AS �������,"
        sStr = sStr & "     TEL1||'-'||TEL2||'-'||TEL3 AS �л���ȭ,"
        sStr = sStr & "     CEL1||'-'||CEL2||'-'||CEL3 AS �л��ڵ���,"
        sStr = sStr & "     EMAIL  �̸���     ,"
        sStr = sStr & "     ZIPCODE ������ȣ    , ADDR2  �����ּ�      , ADDR  ���ּ�      ,"
        
        sStr = sStr & "     /* ��Ž, ��Ž �и� */"
        sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'" & constSatamCodes(0) & "|') > 0 THEN          /* ��Ž-" & constSatams(0) & " */"
        sStr = sStr & "             '" & constSatams(0) & "'"
        sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'51|') > 0 THEN     /* ��Ž-����1 */"
        sStr = sStr & "             '��1'"
        sStr = sStr & "         ELSE"
        sStr = sStr & "             ' '"
        sStr = sStr & "         END END AS Ž��1,"
        sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'" & constSatamCodes(1) & "|') > 0 THEN          /* ��Ž-" & constSatams(1) & " */"
        sStr = sStr & "             '" & constSatams(1) & "'"
        sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'52|') > 0 THEN     /* ��Ž-ȭ��1 */"
        sStr = sStr & "             'ȭ1'"
        sStr = sStr & "         ELSE"
        sStr = sStr & "             ' '"
        sStr = sStr & "         END END AS Ž��2,"
        sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'" & constSatamCodes(2) & "|') > 0 THEN          /* ��Ž-" & constSatams(2) & " */"
        sStr = sStr & "             '" & constSatams(2) & "'"
        sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'53|') > 0 THEN     /* ��Ž-��������1 */"
        sStr = sStr & "             '��1'"
        sStr = sStr & "         ELSE"
        sStr = sStr & "             ' '"
        sStr = sStr & "         END END AS Ž��3,"
        sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'" & constSatamCodes(3) & "|') > 0 THEN          /* ��Ž-" & constSatams(3) & " */"
        sStr = sStr & "             '" & constSatams(3) & "'"
        sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'54|') > 0 THEN     /* ��Ž-��������1 */"
        sStr = sStr & "             '��1'"
        sStr = sStr & "         ELSE"
        sStr = sStr & "             ' '"
        sStr = sStr & "         END END AS Ž��4,"
        sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'" & constSatamCodes(4) & "|') > 0 THEN          /* ��Ž-" & constSatams(4) & " */"
        sStr = sStr & "             '" & constSatams(4) & "'"
        sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'55|') > 0 THEN     /* ��Ž-����2 */"
        sStr = sStr & "             '��2'"
        sStr = sStr & "         ELSE"
        sStr = sStr & "             ' '"
        sStr = sStr & "         END END AS Ž��5,"
        sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'" & constSatamCodes(5) & "|') > 0 THEN          /* ��Ž-" & constSatams(5) & " */"
        sStr = sStr & "             '" & constSatams(5) & "'"
        sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'56|') > 0 THEN     /* ��Ž-ȭ��2 */"
        sStr = sStr & "             'ȭ2'"
        sStr = sStr & "         ELSE"
        sStr = sStr & "             ' '"
        sStr = sStr & "         END END AS Ž��6,"
        sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'" & constSatamCodes(6) & "|') > 0 THEN          /* ��Ž-" & constSatams(6) & " */"
        sStr = sStr & "             '" & constSatams(6) & "'"
        sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'57|') > 0 THEN     /* ��Ž-��������2 */"
        sStr = sStr & "             '��2'"
        sStr = sStr & "         ELSE"
        sStr = sStr & "             ' '"
        sStr = sStr & "         END END AS Ž��7,"
        sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'" & constSatamCodes(7) & "|') > 0 THEN          /* ��Ž-" & constSatams(7) & " */"
        sStr = sStr & "             '" & constSatams(7) & "'"
        sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'58|') > 0 THEN     /* ��Ž-��������2 */"
        sStr = sStr & "             '��2'"
        sStr = sStr & "         ELSE"
        sStr = sStr & "             ' '"
        sStr = sStr & "         END END AS Ž��8,"
        sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'" & constSatamCodes(8) & "|') > 0 THEN          /* ��Ž-" & constSatams(8) & " */"
        sStr = sStr & "             '" & constSatams(8) & "'"
        sStr = sStr & "         ELSE"
        sStr = sStr & "             ' '"
        sStr = sStr & "         END AS Ž��9,"
        sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'" & constSatamCodes(9) & "|') > 0 THEN          /* ��Ž-" & constSatams(9) & " */"
        sStr = sStr & "             '" & constSatams(9) & "'"
        sStr = sStr & "         ELSE"
        sStr = sStr & "             ' '"
        sStr = sStr & "         END AS Ž��10,"
        sStr = sStr & " '' AS Ž��11,"
    
        sStr = sStr & " "
        sStr = sStr & "     GRADE_KOR AS ���,"
        sStr = sStr & "     DECODE(PTS1,'1','��������','2','��������','') AS ��������,"
        sStr = sStr & "     GRADE_MAT AS ����,"
        sStr = sStr & "     GRADE_ENG AS �ܱ���,"
        sStr = sStr & " "
        sStr = sStr & "     HAKCD �б��ڵ�      , GET_SCHOOLNM(HAKCD) AS �����б�, "
        sStr = sStr & "     FILENM �̹������     ,"
        sStr = sStr & " "
        sStr = sStr & "     PRTNM �θ��      , DECODE(PRTREL,'1','��','2','��') AS ����,"
        sStr = sStr & "     PZIPCODE P������ȣ   , PADDR2 P�����ּ�     , PADDR P���ּ�      , PJOB ����       ,"
        
        sStr = sStr & "     PTEL1||'-'||PTEL2||'-'||PTEL3 AS �θ���ȭ,"
        sStr = sStr & "     JTEL1||'-'||JTEL2||'-'||JTEL3 AS �θ��ڵ���,"
        
        sStr = sStr & "     ACC_NO   ����   ,"
        sStr = sStr & "     AMNT     �ݾ�  ,"
        sStr = sStr & "     TO_CHAR(REG_DATE,'YYYY-MM-DD') AS �������"
        
    sStr = sStr & "    From HWSIN01TB_WINTER "
    sStr = sStr & "   WHERE EXMROUND LIKE "
    Select Case Trim(SchCD)
        Case "N"
            sStr = sStr & "            'NR081126E%' "
        Case "K"
            sStr = sStr & "            'KN081126E%' "
        Case "S"
            sStr = sStr & "            'SP081126E%' "
        Case "P"
            sStr = sStr & "            'MK081126E%' "
        Case "M"
            sStr = sStr & "            'NR081126E%' "
            
        Case "W"
            sStr = sStr & "            'KN081126E%' "
        Case "Q"
            sStr = sStr & "            'KN081126E%' "
            
        Case "J"
            sStr = sStr & "            'YJ081126E%' "
        Case "B"
            sStr = sStr & "            'BS081126E%' "
        
        Case Else
            sStr = sStr & "            'BS081126E%' "
    End Select
    
    sStr = sStr & "     AND ACACD  = '" & Trim(basModule.SchCD) & "'"
    
    Set DBCmd = New ADODB.Command
    Set DBRec = New ADODB.Recordset
    Set DBParam = New ADODB.Parameter
    
    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
    DBCmd.CommandText = sStr
    DBCmd.CommandType = adCmdText
    DBCmd.CommandTimeout = 30
    
    
    
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





















'=========================================================================================

'>> �����ּ�
Private Sub txtFAddr_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    
    Dim sStr        As String
    Dim sTmp        As String
    Dim nRec        As Long
    
    On Error GoTo ErrStmt
    
    If KeyCode = vbKeyReturn Then
        
        sStr = ""
        sStr = sStr & " SELECT ZIPCODE, SIDO||' '||GUGUN||' '||DONG AS ADDR, BUNJI"
        sStr = sStr & "   From HWEXM03TB"
        sStr = sStr & "  WHERE DONG LIKE '%" & Trim(txtFAddr.Text) & "%'"
        
        Set DBCmd = New ADODB.Command
        Set DBRec = New ADODB.Recordset
        Set DBParam = New ADODB.Parameter
        
        DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
        DBCmd.CommandText = sStr
        DBCmd.CommandType = adCmdText
        DBCmd.CommandTimeout = 30
        
        


        DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic���·� ���ԵǸ� record count�� �� �� ����.
        Do While DBRec.State And adStateExecuting
            DoEvents
        Loop
        
        sprZip.MaxRows = 0
        With DBRec
            If .RecordCount > 0 Then
                
                .MoveFirst
                For nRec = 1 To .RecordCount Step 1
                
                    sprZip.MaxRows = sprZip.MaxRows + 1
                    sprZip.Row = sprZip.MaxRows
                    
                    sprZip.Col = 1:     sTmp = ""
                        If IsNull(.Fields("ZIPCODE")) = False Then sTmp = Trim(.Fields("ZIPCODE"))
                        sprZip.Text = sTmp
                        
                    sprZip.Col = 2:     sTmp = ""
                        If IsNull(.Fields("ADDR")) = False Then sTmp = Trim(.Fields("ADDR"))
                        sprZip.Text = sTmp
                        
                    sprZip.Col = 3:     sTmp = ""
                        If IsNull(.Fields("BUNJI")) = False Then sTmp = Trim(.Fields("BUNJI"))
                        sprZip.Text = sTmp
                        
                    .MoveNext
                Next nRec
            End If
        End With
    End If
    
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    Exit Sub
ErrStmt:
    
    MsgBox "�ּ� ��ȸ�� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "�����ּ� �˻�"
    On Error GoTo 0
    
End Sub

Private Sub sprZip_Click(ByVal Col As Long, ByVal Row As Long)
    If Row < 1 Then Exit Sub
    
    With sprZip
        If .MaxRows < 1 Then Exit Sub
        
        sprZip.Enabled = False
        
            If .Tag = "" Then .Tag = "1"
            
            .Row = CLng(.Tag):  .Row2 = .Row
            .Col = 1:           .Col2 = .MaxCols
            .BlockMode = True
                .BackColor = basModule.BackColor1
                .BackColorStyle = BackColorStyleUnderGrid
            .BlockMode = False
            
            DoEvents
            
            .Row = Row:         .Row2 = .Row
            .Col = 1:           .Col2 = .MaxCols
            .BlockMode = True
                .BackColor = basModule.SelectColor1
                .BackColorStyle = BackColorStyleUnderGrid
            .BlockMode = False
            
            .Tag = Trim(CStr(Row))
            
            
            Select Case fraAddr.Tag
                Case "S"
                    sprZip.Col = 1:     fpZip.Text = sprZip.Text
                    sprZip.Col = 2:     txtAddr1.Text = sprZip.Text
                Case "P"
                    sprZip.Col = 1:     fpPZipCD.Text = sprZip.Text
                    sprZip.Col = 2:     txtPAddr1.Text = sprZip.Text
            End Select
            
        sprZip.Enabled = True
        sprZip.SetFocus
        sprZip.SetActiveCell Col, Row
        
    End With
    
End Sub

Private Sub sprZip_DblClick(ByVal Col As Long, ByVal Row As Long)
    fraAddr.Visible = False
    
    Select Case fraAddr.Tag
        Case "S"
            txtAddr2.SetFocus
        Case "P"
            txtPAddr2.SetFocus
    End Select
    
End Sub


'>> �б�
Private Sub txtFHak_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    
    Dim sStr        As String
    Dim sTmp        As String
    Dim nRec        As Long
    
    On Error GoTo ErrStmt
    
    If KeyCode = vbKeyReturn Then
        
        sStr = ""
        sStr = sStr & " SELECT JIYK, HAKCD, HAKNM"
        sStr = sStr & "   From HWEXM02TB"
        sStr = sStr & "  WHERE HAKNM LIKE '%" & Trim(txtFHak.Text) & "%'"
        
        Set DBCmd = New ADODB.Command
        Set DBRec = New ADODB.Recordset
        Set DBParam = New ADODB.Parameter
        
        DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
        DBCmd.CommandText = sStr
        DBCmd.CommandType = adCmdText
        DBCmd.CommandTimeout = 30
        
        


        DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic���·� ���ԵǸ� record count�� �� �� ����.
        Do While DBRec.State And adStateExecuting
            DoEvents
        Loop
        
        sprHak.MaxRows = 0
        With DBRec
            If .RecordCount > 0 Then
                
                .MoveFirst
                For nRec = 1 To .RecordCount Step 1
                
                    sprHak.MaxRows = sprHak.MaxRows + 1
                    sprHak.Row = sprHak.MaxRows
                    
                    sprHak.Col = 1:     sTmp = ""
                        If IsNull(.Fields("JIYK")) = False Then sTmp = Trim(.Fields("JIYK"))
                        sprHak.Text = sTmp
                        
                    sprHak.Col = 2:     sTmp = ""
                        If IsNull(.Fields("HAKCD")) = False Then sTmp = Trim(.Fields("HAKCD"))
                        sprHak.Text = sTmp
                        
                    sprHak.Col = 3:     sTmp = ""
                        If IsNull(.Fields("HAKNM")) = False Then sTmp = Trim(.Fields("HAKNM"))
                        sprHak.Text = sTmp
                        
                    .MoveNext
                Next nRec
            End If
        End With
    End If
    
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    Exit Sub
ErrStmt:
    
    MsgBox "�б� ��ȸ�� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "�б� �˻�"
    On Error GoTo 0
    
End Sub

Private Sub sprHak_Click(ByVal Col As Long, ByVal Row As Long)
    If Row < 1 Then Exit Sub
    
    With sprHak
        If .MaxRows < 1 Then Exit Sub
        
        sprHak.Enabled = False
        
            If .Tag = "" Then .Tag = "1"
            
            .Row = CLng(.Tag):  .Row2 = .Row
            .Col = 1:           .Col2 = .MaxCols
            .BlockMode = True
                .BackColor = basModule.BackColor1
                .BackColorStyle = BackColorStyleUnderGrid
            .BlockMode = False
            
            DoEvents
            
            .Row = Row:         .Row2 = .Row
            .Col = 1:           .Col2 = .MaxCols
            .BlockMode = True
                .BackColor = basModule.SelectColor1
                .BackColorStyle = BackColorStyleUnderGrid
            .BlockMode = False
            
            .Tag = Trim(CStr(Row))
            
            sprHak.Col = 2:     txtHakCD.Text = sprHak.Text
            sprHak.Col = 3:     txtHakNM.Text = sprHak.Text
            
        sprHak.Enabled = True
        sprHak.SetFocus
        sprHak.SetActiveCell Col, Row
        
    End With
End Sub



Private Sub sprHak_DblClick(ByVal Col As Long, ByVal Row As Long)
    fraHak.Visible = False
    txtHakCD.SetFocus
    
End Sub



























































'## ���� ���ε�
Private Sub Photo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    Dim sFileLocation   As String
'    Dim sSchNO          As String
'    Dim sOrdNO          As String
'    Dim sExmID          As String
'    Dim simageFile      As String
'
'    Dim bRet            As String
'
'    Dim sDiv()          As String
'    Dim nS              As Long
'    Dim sLocalFile      As String
'
'
'    If Button <> vbRightButton Then
'        Exit Sub
'    End If
'
'    If �л�����.Tag = "" Then
'        MsgBox "�л��� ��ȸ�Ͻʽÿ�.", vbExclamation + vbOKOnly, "���� ���ε�"
'        Exit Sub
'    End If
'    If UBound(uSTD) < 1 Then
'        MsgBox "�л��� ��ȸ�Ͻʽÿ�.", vbExclamation + vbOKOnly, "���� ���ε�"
'        Exit Sub
'    End If
'
'    '�����ȣ.tag
'
'    With uSTD(VScroll1.Value)
'        sOrdNO = .ORD_NO
'
'        sFileLocation = .IMAGE_DIR
'        simageFile = .IMAGE_FILE
'
'        bRet = ""
'        If Trim(sOrdNO) = "" Then        '< �̹����� ���� ��쿣 ������ ����
'            bRet = Make_image_Path(sSchNO, sExmID, simageFile)
'
'            If bRet = "" Then
'                MsgBox "��� ������ ������ �ֽ��ϴ�." & vbCrLf & _
'                       "�����ڿ��� �����Ͻʽÿ�.", vbExclamation + vbOKOnly, "���� ���ε�"
'                Exit Sub
'            Else
'                sFileLocation = bRet
'            End If
'        End If
'    End With
'
'    '<< ���� ����� >>
'    If Trim(txtPage) > " " Then
'        sDiv = Split(txtPage.Text, "/", -1, vbTextCompare)
'
'        nS = CLng(sDiv(0))
'        sLocalFile = sSavePath & "\" & uSTD(nS).IMAGE_FILE & ".jpg"       '<< unique key : �п�+�����ȣ
'        If Dir(sLocalFile) > " " Then
'            Kill sLocalFile
'        End If
'    End If
'
'    '���� �ֱ�
''    Load INT900
''    Call INT900.Save_Photo(sFileLocation, sSchNO)
''    INT900.Show
    
End Sub


'## �̹��� ���� ��� ������ ����
Private Function Make_image_Path(ByVal aOrd_No As String, ByVal aExmID As String, ByVal aimageFile As String) As String
'    Dim sFilePath       As String
'
'    Dim sStr            As String
'    Dim DBCmd           As ADODB.Command
'    Dim DBParam         As ADODB.Parameter
'
'    Dim ni              As Long
'    Dim sLocalFile      As String
'    Dim nExe            As Integer
'    Dim f               As Integer
'    Dim MaxSize         As Long
'
'    sFilePath = ""
'    Select Case Trim(basModule.SchCD)
'        Case "N"
'            sFilePath = "/NDOC/dshw/noryangjin/register/ETC/"
'        Case "K", "W", "Q"
'            sFilePath = "/NDOC/dshw/kangnam/register/ETC/"
'        Case "S"
'            sFilePath = "/NDOC/dshw/songpa/register/ETC/"
'        Case "P"
'            sFilePath = "/NDOC/dshw/msongpa/register/ETC/"
'        Case "M"
'            sFilePath = "/NDOC/dshw/mkangnam/register/ETC/"
'        Case "J"
'            sFilePath = "/NDOC/dshw/mgwanghwa/register/ETC/"
'        Case "B"
'            sFilePath = "/NDOC2/PUB/DS/MPHOTO/"
'    End Select
'
'    sFilePath = sFilePath & aOrd_No & ".jpg"
'
'    On Error GoTo ErrStmt
'
'    basDataBase.DBConn.BeginTrans
'
'    Set DBCmd = New ADODB.Command
'    Set DBParam = New ADODB.Parameter
'
'    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
'


'
'    '<< UPDATE
'    sStr = ""
'    sStr = sStr & " Update HWSIN01TB_WINTER"
'    sStr = sStr & "    SET FILENM = '" & sFilePath & "'"
'    sStr = sStr & "  WHERE ORD_NO = '" & Trim(aOrd_No) & "'"
'
'    DBCmd.CommandText = sStr
'    DBCmd.CommandType = adCmdText
'    DBCmd.CommandTimeout = 30
'
'    DBCmd.Execute nExe, , -1
'
'    Do While basDataBase.DBConn.State And adStateExecuting
'        DoEvents
'    Loop
'
'    If nExe = 1 Then
'        basDataBase.DBConn.CommitTrans
'
'        Set DBCmd = Nothing
'        Set DBParam = Nothing
'
'        f = FreeFile()
'        sLocalFile = sSavePath & "\" & aimageFile & ".jpg"               '<< unique key : �п�+�����ȣ
'        If Dir(sLocalFile) > " " Then
'            Open sLocalFile For Binary As #f
'                On Error Resume Next
'                MaxSize = LOF(f)
'            Close f
'
'            Kill sLocalFile
'
'        End If
'
'        Make_image_Path = sFilePath
'    Else
'        basDataBase.DBConn.RollbackTrans
'
'        Set DBCmd = Nothing
'        Set DBParam = Nothing
'
'        Make_image_Path = ""
'    End If
'
'    Exit Function
'
'ErrStmt:
'    basDataBase.DBConn.RollbackTrans
'
'    Set DBCmd = Nothing
'    Set DBParam = Nothing
'
'    Make_image_Path = ""
End Function





