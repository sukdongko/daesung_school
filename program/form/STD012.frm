VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form STD012 
   Caption         =   "���л��� >> �л���ü ��ȸ (������)"
   ClientHeight    =   9690
   ClientLeft      =   6045
   ClientTop       =   1770
   ClientWidth     =   15450
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9690
   ScaleWidth      =   15450
   Begin VB.Frame Frame10 
      BackColor       =   &H00C6AD84&
      BorderStyle     =   0  '����
      Caption         =   "Frame10"
      Height          =   10245
      Left            =   7380
      TabIndex        =   17
      Top             =   14190
      Visible         =   0   'False
      Width           =   8925
      Begin VB.Frame Frame9 
         BackColor       =   &H00F7EFE7&
         BorderStyle     =   0  '����
         Caption         =   "Frame9"
         Height          =   9435
         Left            =   60
         TabIndex        =   18
         Top             =   240
         Width           =   8295
         Begin VB.CommandButton cmdStdDel 
            Caption         =   "�л������ϱ�"
            Height          =   450
            Left            =   6090
            TabIndex        =   108
            Top             =   8820
            Width           =   1815
         End
         Begin VB.CommandButton cmdStdin 
            Caption         =   "�л���� �� �����ϱ� (&S)"
            Height          =   450
            Left            =   3150
            TabIndex        =   107
            Top             =   8820
            Width           =   2595
         End
         Begin VB.Frame Frame17 
            BackColor       =   &H0082C8E8&
            BorderStyle     =   0  '����
            Caption         =   "Frame17"
            Height          =   825
            Left            =   30
            TabIndex        =   100
            Top             =   7800
            Width           =   8235
            Begin VB.Frame Frame8 
               BackColor       =   &H00F7EFE7&
               BorderStyle     =   0  '����
               Caption         =   ">> ��� ���ð���"
               Height          =   765
               Left            =   30
               TabIndex        =   101
               Top             =   30
               Width           =   8175
               Begin VB.CheckBox chkNonsul 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "���"
                  Height          =   345
                  Index           =   1
                  Left            =   240
                  TabIndex        =   105
                  Top             =   360
                  Width           =   1245
               End
               Begin VB.CheckBox chkNonsul 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "����"
                  Height          =   345
                  Index           =   2
                  Left            =   1590
                  TabIndex        =   104
                  Top             =   360
                  Width           =   1245
               End
               Begin VB.CheckBox chkNonsul 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "��ȸŽ��"
                  Height          =   345
                  Index           =   3
                  Left            =   2940
                  TabIndex        =   103
                  Top             =   360
                  Width           =   1245
               End
               Begin VB.CheckBox chkNonsul 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "����Ž��"
                  Height          =   345
                  Index           =   4
                  Left            =   4290
                  TabIndex        =   102
                  Top             =   360
                  Width           =   1245
               End
               Begin VB.Label Label15 
                  BackStyle       =   0  '����
                  Caption         =   ">> ��� ���ð���"
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
                  TabIndex        =   106
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
            TabIndex        =   93
            Top             =   6930
            Width           =   8235
            Begin VB.Frame Frame7 
               BackColor       =   &H00F7EFE7&
               BorderStyle     =   0  '����
               Caption         =   ">> �������� ���ð���"
               Height          =   765
               Left            =   30
               TabIndex        =   94
               Top             =   30
               Width           =   8175
               Begin VB.CheckBox chkMath 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "������"
                  Height          =   345
                  Index           =   1
                  Left            =   240
                  TabIndex        =   98
                  Top             =   390
                  Width           =   1245
               End
               Begin VB.CheckBox chkMath 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "�̻����"
                  Height          =   345
                  Index           =   2
                  Left            =   1590
                  TabIndex        =   97
                  Top             =   390
                  Width           =   1245
               End
               Begin VB.CheckBox chkMath 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "Ȯ�����"
                  Height          =   345
                  Index           =   3
                  Left            =   2940
                  TabIndex        =   96
                  Top             =   390
                  Width           =   1245
               End
               Begin VB.CheckBox chkMath 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "��������"
                  Height          =   345
                  Index           =   4
                  Left            =   4290
                  TabIndex        =   95
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
                  TabIndex        =   99
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
            Left            =   30
            TabIndex        =   82
            Top             =   5670
            Width           =   8235
            Begin VB.Frame Frame6 
               BackColor       =   &H00F7EFE7&
               BorderStyle     =   0  '����
               Caption         =   ">> ����Ž�� ���ð���"
               Height          =   1155
               Left            =   30
               TabIndex        =   83
               Top             =   30
               Width           =   8175
               Begin VB.CheckBox chkGwatam 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "����1"
                  Height          =   345
                  Index           =   1
                  Left            =   240
                  TabIndex        =   91
                  Top             =   360
                  Width           =   1245
               End
               Begin VB.CheckBox chkGwatam 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "ȭ��1"
                  Height          =   345
                  Index           =   2
                  Left            =   1620
                  TabIndex        =   90
                  Top             =   360
                  Width           =   1245
               End
               Begin VB.CheckBox chkGwatam 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "����1"
                  Height          =   345
                  Index           =   3
                  Left            =   2970
                  TabIndex        =   89
                  Top             =   360
                  Width           =   1245
               End
               Begin VB.CheckBox chkGwatam 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "��������1"
                  Height          =   345
                  Index           =   4
                  Left            =   4320
                  TabIndex        =   88
                  Top             =   360
                  Width           =   1245
               End
               Begin VB.CheckBox chkGwatam 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "����2"
                  Height          =   345
                  Index           =   5
                  Left            =   240
                  TabIndex        =   87
                  Top             =   780
                  Width           =   1245
               End
               Begin VB.CheckBox chkGwatam 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "ȭ��2"
                  Height          =   345
                  Index           =   6
                  Left            =   1620
                  TabIndex        =   86
                  Top             =   780
                  Width           =   1245
               End
               Begin VB.CheckBox chkGwatam 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "����2"
                  Height          =   345
                  Index           =   7
                  Left            =   2970
                  TabIndex        =   85
                  Top             =   780
                  Width           =   1245
               End
               Begin VB.CheckBox chkGwatam 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "��������2"
                  Height          =   345
                  Index           =   8
                  Left            =   4320
                  TabIndex        =   84
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
                  TabIndex        =   92
                  Top             =   90
                  Width           =   2625
               End
            End
         End
         Begin VB.Frame Frame14 
            BackColor       =   &H0082C8E8&
            BorderStyle     =   0  '����
            Caption         =   "Frame14"
            Height          =   825
            Left            =   30
            TabIndex        =   81
            Top             =   4800
            Width           =   8235
            Begin VB.Frame fraSEL2 
               BackColor       =   &H00F7EFE7&
               BorderStyle     =   0  '����
               Caption         =   ">> ��2 �ܱ��� ���ð���"
               Height          =   765
               Left            =   30
               TabIndex        =   137
               Top             =   30
               Width           =   8175
               Begin VB.CheckBox chkEng2 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "�ѹ�"
                  Height          =   345
                  Index           =   6
                  Left            =   7140
                  TabIndex        =   148
                  Top             =   240
                  Width           =   1245
               End
               Begin VB.CheckBox chkEng2 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "�߱���"
                  Height          =   345
                  Index           =   5
                  Left            =   5820
                  TabIndex        =   147
                  Top             =   240
                  Width           =   1245
               End
               Begin VB.CheckBox chkEng2 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "�Ҿ�"
                  Height          =   345
                  Index           =   4
                  Left            =   4320
                  TabIndex        =   146
                  Top             =   240
                  Width           =   1245
               End
               Begin VB.CheckBox chkEng2 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "�����ĳľ�"
                  Height          =   345
                  Index           =   3
                  Left            =   2970
                  TabIndex        =   145
                  Top             =   240
                  Width           =   1245
               End
               Begin VB.CheckBox chkEng2 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "�Ͼ�"
                  Height          =   345
                  Index           =   2
                  Left            =   1620
                  TabIndex        =   144
                  Top             =   240
                  Width           =   1245
               End
               Begin VB.CheckBox chkEng2 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "����"
                  Height          =   345
                  Index           =   1
                  Left            =   240
                  TabIndex        =   143
                  Top             =   240
                  Width           =   1245
               End
               Begin VB.CheckBox chkEng2 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "���"
                  Height          =   345
                  Index           =   7
                  Left            =   240
                  TabIndex        =   142
                  Top             =   510
                  Width           =   1245
               End
               Begin VB.CheckBox chkEng2 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "����"
                  Height          =   345
                  Index           =   8
                  Left            =   1620
                  TabIndex        =   141
                  Top             =   510
                  Width           =   1245
               End
               Begin VB.CheckBox chkEng2 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "����"
                  Height          =   345
                  Index           =   9
                  Left            =   2970
                  TabIndex        =   140
                  Top             =   510
                  Width           =   1245
               End
               Begin VB.CheckBox chkEng2 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "�����"
                  Height          =   345
                  Index           =   10
                  Left            =   4320
                  TabIndex        =   139
                  Top             =   510
                  Width           =   1245
               End
               Begin VB.CheckBox chkEng2 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "��������"
                  Height          =   345
                  Index           =   11
                  Left            =   5820
                  TabIndex        =   138
                  Top             =   510
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
                  TabIndex        =   149
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
            TabIndex        =   67
            Top             =   3540
            Width           =   8235
            Begin VB.Frame Frame2 
               BackColor       =   &H00F7EFE7&
               BorderStyle     =   0  '����
               Caption         =   ">> ��ȸŽ�� ���ð���"
               Height          =   1155
               Left            =   30
               TabIndex        =   68
               Top             =   30
               Width           =   8175
               Begin VB.CheckBox chkSatam 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "����"
                  Height          =   345
                  Index           =   1
                  Left            =   240
                  TabIndex        =   79
                  Top             =   330
                  Width           =   1245
               End
               Begin VB.CheckBox chkSatam 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "����"
                  Height          =   345
                  Index           =   2
                  Left            =   1620
                  TabIndex        =   78
                  Top             =   330
                  Width           =   1245
               End
               Begin VB.CheckBox chkSatam 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "����"
                  Height          =   345
                  Index           =   3
                  Left            =   2970
                  TabIndex        =   77
                  Top             =   330
                  Width           =   1245
               End
               Begin VB.CheckBox chkSatam 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "�ѱ�������"
                  Height          =   345
                  Index           =   4
                  Left            =   4320
                  TabIndex        =   76
                  Top             =   330
                  Width           =   1245
               End
               Begin VB.CheckBox chkSatam 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "�����"
                  Height          =   345
                  Index           =   5
                  Left            =   5790
                  TabIndex        =   75
                  Top             =   330
                  Width           =   1245
               End
               Begin VB.CheckBox chkSatam 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "��������"
                  Height          =   345
                  Index           =   6
                  Left            =   240
                  TabIndex        =   74
                  Top             =   780
                  Width           =   1245
               End
               Begin VB.CheckBox chkSatam 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "�ѱ�����"
                  Height          =   345
                  Index           =   7
                  Left            =   1620
                  TabIndex        =   73
                  Top             =   750
                  Width           =   1245
               End
               Begin VB.CheckBox chkSatam 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "��ġ"
                  Height          =   345
                  Index           =   8
                  Left            =   2970
                  TabIndex        =   72
                  Top             =   750
                  Width           =   1245
               End
               Begin VB.CheckBox chkSatam 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "��ȸ��ȭ"
                  Height          =   345
                  Index           =   9
                  Left            =   4320
                  TabIndex        =   71
                  Top             =   750
                  Width           =   1245
               End
               Begin VB.CheckBox chkSatam 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "����"
                  Height          =   345
                  Index           =   10
                  Left            =   5790
                  TabIndex        =   70
                  Top             =   750
                  Width           =   1245
               End
               Begin VB.CheckBox chkSatam 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "��������"
                  Height          =   345
                  Index           =   11
                  Left            =   7110
                  TabIndex        =   69
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
                  TabIndex        =   80
                  Top             =   90
                  Width           =   2625
               End
            End
         End
         Begin VB.Frame Frame12 
            BackColor       =   &H00C6AD84&
            BorderStyle     =   0  '����
            Caption         =   "Frame12"
            Height          =   825
            Left            =   30
            TabIndex        =   58
            Top             =   2670
            Width           =   8235
            Begin VB.Frame Frame4 
               BackColor       =   &H00F7EFE7&
               BorderStyle     =   0  '����
               Caption         =   ">> ����"
               Height          =   765
               Left            =   30
               TabIndex        =   59
               Top             =   30
               Width           =   8175
               Begin EditLib.fpLongInteger fpK_Num 
                  Height          =   345
                  Left            =   1140
                  TabIndex        =   60
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
                  Left            =   2820
                  TabIndex        =   61
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
                  Left            =   4590
                  TabIndex        =   62
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
                  TabIndex        =   66
                  Top             =   30
                  Width           =   2625
               End
               Begin VB.Label Label6 
                  Alignment       =   1  '������ ����
                  BackStyle       =   0  '����
                  Caption         =   "����"
                  Height          =   210
                  Left            =   0
                  TabIndex        =   65
                  Top             =   360
                  Width           =   975
               End
               Begin VB.Label Label7 
                  Alignment       =   1  '������ ����
                  BackStyle       =   0  '����
                  Caption         =   "����"
                  Height          =   210
                  Left            =   1680
                  TabIndex        =   64
                  Top             =   360
                  Width           =   975
               End
               Begin VB.Label Label8 
                  Alignment       =   1  '������ ����
                  BackStyle       =   0  '����
                  Caption         =   "����"
                  Height          =   210
                  Left            =   3450
                  TabIndex        =   63
                  Top             =   360
                  Width           =   975
               End
            End
         End
         Begin VB.Frame Frame11 
            BackColor       =   &H00C6AD84&
            BorderStyle     =   0  '����
            Caption         =   "Frame11"
            Height          =   2535
            Left            =   30
            TabIndex        =   19
            Top             =   90
            Width           =   8235
            Begin VB.Frame Frame3 
               BackColor       =   &H00F7EFE7&
               BorderStyle     =   0  '����
               Caption         =   ">> �⺻�׸�"
               Height          =   2475
               Left            =   30
               TabIndex        =   20
               Top             =   30
               Width           =   8175
               Begin VB.TextBox txtPayGbn 
                  Enabled         =   0   'False
                  Height          =   270
                  IMEMode         =   10  '�ѱ� 
                  Left            =   6810
                  TabIndex        =   37
                  Text            =   "txtPayGbn"
                  Top             =   2250
                  Width           =   1275
               End
               Begin VB.TextBox txtRegDate 
                  Enabled         =   0   'False
                  Height          =   270
                  IMEMode         =   10  '�ѱ� 
                  Left            =   3900
                  TabIndex        =   36
                  Text            =   "txtRegDate"
                  Top             =   2220
                  Width           =   1395
               End
               Begin VB.TextBox txtCel 
                  Height          =   270
                  IMEMode         =   10  '�ѱ� 
                  Left            =   3900
                  TabIndex        =   35
                  Text            =   "txtCel"
                  Top             =   1875
                  Width           =   1395
               End
               Begin VB.TextBox txtOrdNo 
                  Enabled         =   0   'False
                  Height          =   270
                  IMEMode         =   10  '�ѱ� 
                  Left            =   6810
                  TabIndex        =   34
                  Text            =   "txtOrdNo"
                  Top             =   1965
                  Width           =   1275
               End
               Begin VB.TextBox txtTel 
                  Height          =   270
                  IMEMode         =   10  '�ѱ� 
                  Left            =   3900
                  TabIndex        =   33
                  Text            =   "9999-9999-9999"
                  Top             =   1560
                  Width           =   1395
               End
               Begin VB.ComboBox cboKaeyol 
                  Height          =   300
                  Left            =   3900
                  Style           =   2  '��Ӵٿ� ���
                  TabIndex        =   32
                  Top             =   352
                  Width           =   1395
               End
               Begin VB.ComboBox cboPass4 
                  Height          =   300
                  Left            =   6810
                  Style           =   2  '��Ӵٿ� ���
                  TabIndex        =   31
                  Top             =   1612
                  Width           =   1275
               End
               Begin VB.ComboBox cboPass3 
                  Height          =   300
                  Left            =   6810
                  Style           =   2  '��Ӵٿ� ���
                  TabIndex        =   30
                  Top             =   1192
                  Width           =   1275
               End
               Begin VB.ComboBox cboPass2 
                  Height          =   300
                  Left            =   6810
                  Style           =   2  '��Ӵٿ� ���
                  TabIndex        =   29
                  Top             =   772
                  Width           =   1275
               End
               Begin VB.ComboBox cboPass1 
                  Height          =   300
                  Left            =   6810
                  Style           =   2  '��Ӵٿ� ���
                  TabIndex        =   28
                  Top             =   352
                  Width           =   1275
               End
               Begin VB.ComboBox cboSel2_Sch 
                  Height          =   300
                  Left            =   3900
                  Style           =   2  '��Ӵٿ� ���
                  TabIndex        =   27
                  Top             =   1192
                  Width           =   1395
               End
               Begin VB.ComboBox cboSel1_Sch 
                  Height          =   300
                  Left            =   3900
                  Style           =   2  '��Ӵٿ� ���
                  TabIndex        =   26
                  Top             =   772
                  Width           =   1395
               End
               Begin VB.TextBox txtSchNo 
                  BackColor       =   &H00C0FFFF&
                  Height          =   345
                  Left            =   1140
                  TabIndex        =   25
                  Text            =   "txtSchNo"
                  Top             =   330
                  Width           =   1605
               End
               Begin VB.TextBox txtStdNM 
                  Height          =   345
                  IMEMode         =   10  '�ѱ� 
                  Left            =   1140
                  TabIndex        =   24
                  Text            =   "txtStdNM"
                  Top             =   1170
                  Width           =   1605
               End
               Begin VB.Frame Frame1 
                  BackColor       =   &H00F7EFE7&
                  BorderStyle     =   0  '����
                  Height          =   435
                  Left            =   1140
                  TabIndex        =   21
                  Top             =   2025
                  Width           =   1965
                  Begin VB.OptionButton optExmN 
                     BackColor       =   &H00F7EFE7&
                     Caption         =   "������"
                     Height          =   285
                     Left            =   1050
                     TabIndex        =   23
                     Top             =   90
                     Width           =   885
                  End
                  Begin VB.OptionButton optExmY 
                     BackColor       =   &H00F7EFE7&
                     Caption         =   "������"
                     Height          =   285
                     Left            =   0
                     TabIndex        =   22
                     Top             =   90
                     Width           =   885
                  End
               End
               Begin EditLib.fpMask fpExmID 
                  Height          =   345
                  Left            =   1140
                  TabIndex        =   38
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
               Begin EditLib.fpMask fpJumin 
                  Height          =   345
                  Left            =   1140
                  TabIndex        =   39
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
               Begin VB.Label Label43 
                  Alignment       =   1  '������ ����
                  BackStyle       =   0  '����
                  Caption         =   "������"
                  ForeColor       =   &H00C000C0&
                  Height          =   180
                  Left            =   5370
                  TabIndex        =   57
                  Top             =   2295
                  Width           =   1365
               End
               Begin VB.Label Label42 
                  Alignment       =   1  '������ ����
                  BackStyle       =   0  '����
                  Caption         =   "�������"
                  ForeColor       =   &H00C000C0&
                  Height          =   180
                  Left            =   2520
                  TabIndex        =   56
                  Top             =   2265
                  Width           =   1365
               End
               Begin VB.Label Label41 
                  Alignment       =   1  '������ ����
                  BackStyle       =   0  '����
                  Caption         =   "�ڵ���"
                  ForeColor       =   &H00000000&
                  Height          =   210
                  Left            =   2850
                  TabIndex        =   55
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
                  TabIndex        =   54
                  Top             =   2010
                  Width           =   1365
               End
               Begin VB.Label Label39 
                  Alignment       =   1  '������ ����
                  BackStyle       =   0  '����
                  Caption         =   "TEL"
                  ForeColor       =   &H00000000&
                  Height          =   210
                  Left            =   2850
                  TabIndex        =   53
                  Top             =   1620
                  Width           =   975
               End
               Begin VB.Label Label28 
                  Alignment       =   1  '������ ����
                  BackStyle       =   0  '����
                  Caption         =   "��  ��"
                  Height          =   210
                  Left            =   2880
                  TabIndex        =   52
                  Top             =   397
                  Width           =   975
               End
               Begin VB.Label Label21 
                  Alignment       =   1  '������ ����
                  BackStyle       =   0  '����
                  Caption         =   "4���� �հ��п�"
                  Height          =   210
                  Left            =   5280
                  TabIndex        =   51
                  Top             =   1650
                  Width           =   1455
               End
               Begin VB.Label Label20 
                  Alignment       =   1  '������ ����
                  BackStyle       =   0  '����
                  Caption         =   "3���� �հ��п�"
                  Height          =   210
                  Left            =   5280
                  TabIndex        =   50
                  Top             =   1230
                  Width           =   1455
               End
               Begin VB.Label Label19 
                  Alignment       =   1  '������ ����
                  BackStyle       =   0  '����
                  Caption         =   "2���� �հ��п�"
                  Height          =   210
                  Left            =   5280
                  TabIndex        =   49
                  Top             =   810
                  Width           =   1455
               End
               Begin VB.Label Label18 
                  Alignment       =   1  '������ ����
                  BackStyle       =   0  '����
                  Caption         =   "1���� �հ��п�"
                  Height          =   210
                  Left            =   5280
                  TabIndex        =   48
                  Top             =   390
                  Width           =   1455
               End
               Begin VB.Label Label17 
                  Alignment       =   1  '������ ����
                  BackStyle       =   0  '����
                  Caption         =   "2���� �п�"
                  Height          =   210
                  Left            =   2880
                  TabIndex        =   47
                  Top             =   1237
                  Width           =   975
               End
               Begin VB.Label Label16 
                  Alignment       =   1  '������ ����
                  BackStyle       =   0  '����
                  Caption         =   "1���� �п�"
                  Height          =   210
                  Left            =   2880
                  TabIndex        =   46
                  Top             =   817
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
                  TabIndex        =   45
                  Top             =   60
                  Width           =   2625
               End
               Begin VB.Label Label4 
                  Alignment       =   1  '������ ����
                  BackStyle       =   0  '����
                  Caption         =   "�й�"
                  Height          =   210
                  Left            =   60
                  TabIndex        =   44
                  Top             =   390
                  Width           =   975
               End
               Begin VB.Label Label1 
                  Alignment       =   1  '������ ����
                  BackStyle       =   0  '����
                  Caption         =   "�����ȣ"
                  Height          =   210
                  Left            =   60
                  TabIndex        =   43
                  Top             =   810
                  Width           =   975
               End
               Begin VB.Label Label2 
                  Alignment       =   1  '������ ����
                  BackStyle       =   0  '����
                  Caption         =   "�л���"
                  Height          =   210
                  Left            =   60
                  TabIndex        =   42
                  Top             =   1230
                  Width           =   975
               End
               Begin VB.Label Label3 
                  Alignment       =   1  '������ ����
                  BackStyle       =   0  '����
                  Caption         =   "�ֹι�ȣ"
                  Height          =   210
                  Left            =   60
                  TabIndex        =   41
                  Top             =   1650
                  Width           =   975
               End
               Begin VB.Label Label5 
                  Alignment       =   1  '������ ����
                  BackStyle       =   0  '����
                  Caption         =   "��/������"
                  Height          =   210
                  Left            =   60
                  TabIndex        =   40
                  Top             =   2160
                  Width           =   975
               End
            End
         End
      End
   End
   Begin VB.Frame fraGwamok 
      BackColor       =   &H0082C8E8&
      BorderStyle     =   0  '����
      Caption         =   "����"
      Height          =   4275
      Left            =   2100
      TabIndex        =   133
      Top             =   9840
      Width           =   8865
      Begin VB.Frame Frame23 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  '����
         Height          =   4215
         Left            =   30
         TabIndex        =   134
         Top             =   30
         Width           =   8805
         Begin VB.CommandButton cmdClose 
            Caption         =   "�ݱ�"
            Height          =   330
            Left            =   8160
            TabIndex        =   135
            Top             =   3840
            Width           =   585
         End
         Begin VB.Image Image1 
            Height          =   4080
            Left            =   30
            Picture         =   "STD012.frx":0000
            Top             =   60
            Width           =   8730
         End
      End
   End
   Begin VB.Frame Frame20 
      BackColor       =   &H0082C8E8&
      BorderStyle     =   0  '����
      Caption         =   "Frame20"
      Height          =   4365
      Left            =   60
      TabIndex        =   125
      Top             =   14190
      Visible         =   0   'False
      Width           =   6615
      Begin VB.Frame Frame21 
         BackColor       =   &H00D2EAF5&
         BorderStyle     =   0  '����
         Caption         =   "Frame21"
         Height          =   3285
         Left            =   30
         TabIndex        =   126
         Top             =   120
         Width           =   6555
         Begin VB.CommandButton cmdGwamokView 
            Caption         =   "���񺸱�"
            Height          =   315
            Left            =   4260
            TabIndex        =   129
            Top             =   870
            Width           =   885
         End
         Begin VB.CommandButton cmdExcelSave 
            Caption         =   "�����ڷ� �����ϱ�"
            Height          =   450
            Left            =   4470
            TabIndex        =   128
            Top             =   2760
            Width           =   1875
         End
         Begin VB.CommandButton cmdGetExcel 
            Caption         =   "�����ڷ� ��������"
            Height          =   390
            Left            =   4410
            TabIndex        =   127
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
            Height          =   1455
            Left            =   60
            TabIndex        =   130
            Top             =   1230
            Width           =   6405
            _Version        =   393216
            _ExtentX        =   11298
            _ExtentY        =   2566
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
            SpreadDesigner  =   "STD012.frx":76CA
         End
         Begin VB.Label Label30 
            BackStyle       =   0  '����
            Caption         =   $"STD012.frx":7C12
            Height          =   615
            Left            =   240
            TabIndex        =   132
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
            TabIndex        =   131
            Top             =   120
            Width           =   2625
         End
      End
   End
   Begin VB.Frame Frame18 
      BackColor       =   &H0082C8E8&
      BorderStyle     =   0  '����
      Caption         =   "Frame18"
      Height          =   9465
      Left            =   30
      TabIndex        =   109
      Top             =   30
      Width           =   15015
      Begin VB.Frame Frame19 
         BackColor       =   &H00D2EAF5&
         BorderStyle     =   0  '����
         Caption         =   "Frame19"
         Height          =   9405
         Left            =   30
         TabIndex        =   110
         Top             =   30
         Width           =   14955
         Begin VB.CommandButton cmdAllStdData 
            Caption         =   "������ ������ �ޱ�"
            Height          =   435
            Left            =   12300
            TabIndex        =   15
            Top             =   30
            Width           =   2625
         End
         Begin VB.ComboBox cboinGbn 
            Height          =   300
            Left            =   5160
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   9
            Top             =   555
            Width           =   885
         End
         Begin VB.ComboBox cboExmType 
            Height          =   300
            Left            =   2820
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   8
            Top             =   555
            Width           =   855
         End
         Begin VB.ComboBox cboPay 
            Height          =   300
            Left            =   5670
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   4
            Top             =   90
            Width           =   855
         End
         Begin VB.ComboBox cboPassCN 
            Height          =   300
            Left            =   13650
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   14
            Top             =   555
            Width           =   885
         End
         Begin VB.ComboBox cboKaeyol_F 
            Height          =   300
            Left            =   510
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   7
            Top             =   555
            Width           =   1485
         End
         Begin VB.CommandButton cmdFind 
            Caption         =   "��ȸ�ϱ�(&F)"
            Height          =   480
            Left            =   390
            TabIndex        =   0
            Top             =   0
            Width           =   1305
         End
         Begin VB.TextBox txtStdNM_F 
            Height          =   345
            IMEMode         =   10  '�ѱ� 
            Left            =   9750
            TabIndex        =   12
            Text            =   "txtStdNM_F"
            Top             =   540
            Width           =   825
         End
         Begin VB.ComboBox cboSel1_SCH_F 
            Height          =   300
            Left            =   7110
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   5
            Top             =   90
            Width           =   1005
         End
         Begin VB.ComboBox cboSel2_SCH_F 
            Height          =   300
            Left            =   8700
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   6
            Top             =   90
            Width           =   1185
         End
         Begin EditLib.fpLongInteger fpPayOK 
            Height          =   315
            Left            =   3420
            TabIndex        =   2
            Top             =   90
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
            Height          =   8415
            Left            =   30
            TabIndex        =   16
            Top             =   960
            Width           =   14895
            _Version        =   393216
            _ExtentX        =   26273
            _ExtentY        =   14843
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
            MaxCols         =   37
            SpreadDesigner  =   "STD012.frx":7CA9
         End
         Begin EditLib.fpMask fpExmID_F 
            Height          =   345
            Left            =   7020
            TabIndex        =   10
            Top             =   540
            Width           =   705
            _Version        =   196608
            _ExtentX        =   1244
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
            Left            =   11370
            TabIndex        =   13
            Top             =   540
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
         Begin EditLib.fpMask fpExmID_E 
            Height          =   345
            Left            =   8130
            TabIndex        =   11
            Top             =   540
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
            Left            =   4560
            TabIndex        =   3
            Top             =   90
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
            Left            =   2460
            TabIndex        =   1
            Top             =   90
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
         Begin EditLib.fpLongInteger fpSuNung 
            Height          =   315
            Left            =   11700
            TabIndex        =   150
            Top             =   90
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
         Begin EditLib.fpLongInteger fpSunHang 
            Height          =   315
            Left            =   10680
            TabIndex        =   151
            Top             =   90
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
         Begin VB.Label lblB 
            Alignment       =   1  '������ ����
            BackStyle       =   0  '����
            Caption         =   "����"
            Height          =   210
            Left            =   11220
            TabIndex        =   153
            Top             =   120
            Width           =   465
         End
         Begin VB.Label lblA 
            Alignment       =   1  '������ ����
            BackStyle       =   0  '����
            Caption         =   "���ǰ��"
            Height          =   360
            Left            =   10200
            TabIndex        =   152
            Top             =   60
            Width           =   465
         End
         Begin VB.Label Label38 
            Alignment       =   1  '������ ����
            BackStyle       =   0  '����
            Caption         =   "��ü����"
            ForeColor       =   &H00C000C0&
            Height          =   210
            Left            =   1500
            TabIndex        =   124
            Top             =   135
            Width           =   975
         End
         Begin VB.Label Label37 
            Alignment       =   1  '������ ����
            BackStyle       =   0  '����
            Caption         =   "���ͳ�/�п�����"
            Height          =   210
            Left            =   3660
            TabIndex        =   123
            Top             =   600
            Width           =   1455
         End
         Begin VB.Label Label36 
            Alignment       =   1  '������ ����
            BackStyle       =   0  '����
            Caption         =   "��/������"
            Height          =   210
            Left            =   1800
            TabIndex        =   122
            Top             =   600
            Width           =   975
         End
         Begin VB.Label Label35 
            Alignment       =   1  '������ ����
            BackStyle       =   0  '����
            Caption         =   "�̰���"
            ForeColor       =   &H000000C0&
            Height          =   210
            Left            =   3600
            TabIndex        =   121
            Top             =   135
            Width           =   975
         End
         Begin VB.Label Label34 
            Alignment       =   1  '������ ����
            BackStyle       =   0  '����
            Caption         =   "����"
            ForeColor       =   &H00FF0000&
            Height          =   210
            Left            =   3000
            TabIndex        =   120
            Top             =   135
            Width           =   435
         End
         Begin VB.Label Label33 
            Alignment       =   1  '������ ����
            BackStyle       =   0  '����
            Caption         =   "����"
            Height          =   210
            Left            =   5190
            TabIndex        =   119
            Top             =   135
            Width           =   465
         End
         Begin VB.Label Label32 
            Alignment       =   1  '������ ����
            BackStyle       =   0  '����
            Caption         =   "�հ�����"
            Height          =   210
            Left            =   12660
            TabIndex        =   118
            Top             =   600
            Width           =   975
         End
         Begin VB.Label Label31 
            Alignment       =   1  '������ ����
            BackStyle       =   0  '����
            Caption         =   "�� ��"
            Height          =   210
            Left            =   -90
            TabIndex        =   117
            Top             =   600
            Width           =   525
         End
         Begin VB.Label Label27 
            Alignment       =   1  '������ ����
            BackStyle       =   0  '����
            Caption         =   "�������"
            Height          =   210
            Left            =   10380
            TabIndex        =   116
            Top             =   600
            Width           =   975
         End
         Begin VB.Label Label26 
            Alignment       =   1  '������ ����
            BackStyle       =   0  '����
            Caption         =   "�л���"
            Height          =   210
            Left            =   8700
            TabIndex        =   115
            Top             =   600
            Width           =   975
         End
         Begin VB.Label Label25 
            BackStyle       =   0  '����
            Caption         =   "�����ȣ             ����"
            Height          =   210
            Left            =   6240
            TabIndex        =   114
            Top             =   600
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
            TabIndex        =   113
            Top             =   90
            Width           =   2625
         End
         Begin VB.Label Label23 
            Alignment       =   1  '������ ����
            BackStyle       =   0  '����
            Caption         =   "1�����п�"
            Height          =   210
            Left            =   6420
            TabIndex        =   112
            Top             =   135
            Width           =   975
         End
         Begin VB.Label Label22 
            Alignment       =   1  '������ ����
            BackStyle       =   0  '����
            Caption         =   "2�����п�"
            Height          =   210
            Left            =   7950
            TabIndex        =   111
            Top             =   135
            Width           =   975
         End
      End
   End
   Begin FPSpread.vaSpread sprStdData 
      Height          =   165
      Left            =   90
      TabIndex        =   136
      Top             =   5820
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
      SpreadDesigner  =   "STD012.frx":855C
   End
End
Attribute VB_Name = "STD012"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################################################
'   �� �� ��  �� : �뼺�п� ���л���, �ݹ��� & �ð�ǥ ���α׷�
'   ����ý��۸� :
'   ��   ��   �� : STD011
'   �� ��  �� �� : �л���ü ��ȸ
'
'   ��   ��   �� : 2007/12/13
'   ��   ��   �� : ���ϱ�
' --------------------------------------------------------------------------------------------------------------
' --------------------------------------------------------------------------------------------------------------
'                 ��     ��     ��     ��
' --------------------------------------------------------------------------------------------------------------
'   1. ������ : �л����, �л����, �л����� ���� �޼���� ������.
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
    SATAM11     As String
    
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

Private sini_Path       As String
Private sChasuTimes     As String

Private Sub Form_Terminate()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim sChasuT     As String
    Dim sTmp        As String
    Dim sData       As String * 255
    Dim nRtn        As Long
    
    Me.Move 0, 0, 15255, 9980
    fraGwamok.Visible = False
    
    sini_Path = App.Path & "\DAESUNG.INI"       '<< ini file
    
    sChasuT = "CHASU"
        sTmp = ""
        nRtn = basModule.GetPrivateProfileString(sChasuT, "TIMES", "", sData, 255, sini_Path)
        If nRtn > 0 Then
            sChasuTimes = Left(sData, nRtn)
        Else
            sTmp = "2008030409"
            nRtn = basModule.WritePrivateProfileString(sChasuT, "TIMES", sTmp, sini_Path)
            sChasuTimes = sTmp
        End If
        
    
    With sprSTD_F
        .ShadowColor = basModule.ShadowColor1
        .ShadowDark = basModule.ShadowDark1
        .ShadowText = basModule.ShadowText1
        .GridColor = basModule.GridColor1
        .GrayAreaBackColor = basModule.GrayAreaBackColor1
    End With
    
    With sprExcel_STD_Data
        .ShadowColor = basModule.ShadowColor1
        .ShadowDark = basModule.ShadowDark1
        .ShadowText = basModule.ShadowText1
        .GridColor = basModule.GridColor1
        .GrayAreaBackColor = basModule.GrayAreaBackColor1
    End With
    
    With cboKaeyol
        .Clear
        .AddItem "��ü" & Space(30) & "ALL"
        .AddItem "�ι�" & Space(30) & "01"
        .AddItem "�ڿ�" & Space(30) & "02"
        
    '<< �迭 >> : 2008.01.09
        If Trim(basModule.schcd) = "N" Then             '< �뷮��
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
        If Trim(basModule.schcd) = "K" Or Trim(basModule.schcd) = "W" Or Trim(basModule.schcd) = "Q" Then           '< ���� 2008.03.24
            
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
        If Trim(basModule.schcd) = "S" Then             '< ����
'            .AddItem "��ü��" & Space(30) & "03"
'
'            .AddItem "�ι�����" & Space(30) & "05"
'            .AddItem "�ڿ�����" & Space(30) & "06"
'
'            .AddItem "�ż��ι�" & Space(30) & "11"
'            .AddItem "�ż��ڿ�" & Space(30) & "12"
            
            .AddItem "�ι������̾�" & Space(30) & "18"
            .AddItem "�ڿ������̾�" & Space(30) & "19"

        End If
    '<< �迭 >> : 2008.02.15
        If Trim(basModule.schcd) = "P" Then             '< ����
            .AddItem "Ư���ι�" & Space(30) & "03"
            .AddItem "Ư���ڿ�" & Space(30) & "04"
        End If
        
        If Trim(basModule.schcd) = "J" Then             '< ����
            .AddItem "�ż��ι�" & Space(30) & "03"
            .AddItem "�ż��ڿ�" & Space(30) & "04"
            
            .AddItem "�ι������̾�" & Space(30) & "18"
            .AddItem "�ڿ������̾�" & Space(30) & "19"

        End If
        
    '<< �迭 >> : 2009.01.09
        If Trim(basModule.schcd) = "B" Then             '< �λ�
            .AddItem "�����ι�" & Space(30) & "05"
            .AddItem "�����ڿ�" & Space(30) & "06"
            
            .AddItem "��.����ι�" & Space(30) & "07"
            .AddItem "��.����ڿ�" & Space(30) & "08"
            
            .AddItem "��ȭ�ι�" & Space(30) & "09"
            .AddItem "��ȭ�ڿ�" & Space(30) & "10"
        End If
        
        .ListIndex = 0
    End With
    
    With cboKaeyol_F
        .Clear
        .AddItem "��ü" & Space(30) & "ALL"
        .AddItem "�ι�" & Space(30) & "01"
        .AddItem "�ڿ�" & Space(30) & "02"
        
    '<< �迭 >> : 2008.01.09
        If Trim(basModule.schcd) = "N" Then             '< �뷮��
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
        If Trim(basModule.schcd) = "K" Or Trim(basModule.schcd) = "W" Or Trim(basModule.schcd) = "Q" Then           '< ���� 2008.03.24
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
        If Trim(basModule.schcd) = "S" Then             '< ����
''            .AddItem "��ü��" & Space(30) & "03"
''
''            .AddItem "�ι�����" & Space(30) & "05"
''            .AddItem "�ڿ�����" & Space(30) & "06"
''
''            .AddItem "�ż��ι�" & Space(30) & "11"
''            .AddItem "�ż��ڿ�" & Space(30) & "12"
            
            .AddItem "�ι������̾�" & Space(30) & "18"
            .AddItem "�ڿ������̾�" & Space(30) & "19"

        End If
    '<< �迭 >> : 2008.02.15
        If Trim(basModule.schcd) = "P" Then             '< ����
            .AddItem "Ư���ι�" & Space(30) & "03"
            .AddItem "Ư���ڿ�" & Space(30) & "04"
        End If
        
        If Trim(basModule.schcd) = "J" Then             '< ����
            .AddItem "�ż��ι�" & Space(30) & "11"
            .AddItem "�ż��ڿ�" & Space(30) & "12"
            
            .AddItem "�ι������̾�" & Space(30) & "18"
            .AddItem "�ڿ������̾�" & Space(30) & "19"
        End If
    
    '<< �迭 >> : 2009.01.09
        If Trim(basModule.schcd) = "B" Then             '< �λ�
            .AddItem "�����ι�" & Space(30) & "05"
            .AddItem "�����ڿ�" & Space(30) & "06"
            
            .AddItem "��.����ι�" & Space(30) & "07"
            .AddItem "��.����ڿ�" & Space(30) & "08"
            
            .AddItem "��ȭ�ι�" & Space(30) & "09"
            .AddItem "��ȭ�ڿ�" & Space(30) & "10"
        End If
        
        .ListIndex = 0
    End With
    
    With cboSel1_Sch
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
        
        .ListIndex = 0
    End With
    
    With cboSel1_SCH_F
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
        
        .ListIndex = 0
    End With
    
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
        
        .ListIndex = 0
    End With
    
    With cboSel2_SCH_F
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
        
        .ListIndex = 0
    End With
        
    With cboPass1
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
        
        .ListIndex = 0
    End With
    
    With cboPass2
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
        
        .ListIndex = 0
    End With
    
    With cboPass3
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
        
        .ListIndex = 0
    End With
    
    With cboPass4
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
        
        .ListIndex = 0
    End With
    
    
    With cboPassCN
        .Clear
        .AddItem "��ü" & Space(30) & "ALL"
        .AddItem "1��" & Space(30) & "1"
        .AddItem "2��" & Space(30) & "2"
        .AddItem "3��" & Space(30) & "3"
        .AddItem "4��" & Space(30) & "4"
        
        .ListIndex = 0
    End With
    
    With cboPay
        .Clear
        .AddItem "��ü" & Space(30) & "ALL"
        .AddItem "����" & Space(30) & "OK"
        .AddItem "�̰���" & Space(30) & "NOT"
        
        .ListIndex = 1
    End With
    
    With cboExmType
        .Clear
        .AddItem "��ü" & Space(30) & "ALL"
        .AddItem "������" & Space(30) & "1"
        .AddItem "������" & Space(30) & "0"
        
        .ListIndex = 0
    End With
    
    With cboinGbn
        .Clear
        .AddItem "��ü" & Space(30) & "ALL"
        .AddItem "���ͳ�" & Space(30) & "INT"
        .AddItem "�п�" & Space(30) & "HAK"
        
        .ListIndex = 0
    End With
    
    Call init_Form
    
    
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

Private Sub init_Form()
    Dim ni      As Integer
    
    txtSchNo.Text = ""
    fpExmID.Text = ""
    txtStdNM.Text = ""
    
    optExmY.value = True
    optExmN.value = False
    
    fpK_Num.value = 0
    fpE_Num.value = 0
    fpM_Num.value = 0
    
    For ni = 1 To 11 Step 1
        chkSatam(ni).value = 0
    Next ni
    
    For ni = 1 To 11 Step 1
        chkEng2(ni).value = 0
    Next ni
    
    For ni = 1 To 8 Step 1
        chkGwatam(ni).value = 0
    Next ni
    
    For ni = 1 To 4 Step 1
        chkMath(ni).value = 0
        chkNonsul(ni).value = 0
    Next ni
    
    '>> ��ȸ�κ�
    fpExmID_F.Text = ""
    fpExmID_E.Text = ""
    
    txtStdNM_F.Text = ""
    fpBirth_ymd_F.Text = ""
    sprSTD_F.MaxRows = 0
    
    sprExcel_STD_Data.MaxRows = 0
    
    
    fpPayOK.value = 0
    fpPayNot.value = 0
    fpPayTot.value = 0
    
    txtOrdNo.Text = ""
    txtTel.Text = ""
    txtCel.Text = ""
    
    txtRegDate.Text = ""
    txtPayGbn.Text = ""
    
    fpSunHang.value = 0
    fpSuNung.value = 0
    
    If Trim(basModule.schcd) = "P" Then
        fpSunHang.Visible = True
        fpSuNung.Visible = True
        
        lblA.Visible = True
        lblB.Visible = True
    Else
        fpSunHang.Visible = False
        fpSuNung.Visible = False
        
        lblA.Visible = False
        lblB.Visible = False
    End If
    
End Sub



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
    
    Dim nJumsu      As Integer
    Dim nj          As Integer
    Dim sJTmp       As String
    
    nJumsu = 42     ' 44�� �׸�
    
    On Error GoTo ErrStmt
    
    cmdFind.Enabled = False
    
    sprSTD_F.MaxRows = 0
    fpPayOK.value = 0
    fpPayNot.value = 0
    fpPayTot.value = 0
    
    fpSunHang.value = 0
    fpSuNung.value = 0
    
    sStr = ""
    sStr = sStr & "  SELECT "
    Select Case Trim(Right(cboPassCN, 30))
        Case "ALL"      ' /* �հݻ��� ��� ��ȸ */
            sStr = sStr & "         SCHNO, "
        Case Else
            sStr = sStr & "         A.SCHNO, "
    End Select
    sStr = sStr & "         EXMID, STDNM, SEL1_SCH , SEL2_SCH, Birth_ymd,"
    
    '<< �迭 >> : 2008.01.09
    If Trim(basModule.schcd) = "N" Then
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
        sStr = sStr & "            ) AS GAEYUL,"
    '<< �迭 >> : 2008.01.10
    ElseIf Trim(basModule.schcd) = "K" Or Trim(basModule.schcd) = "W" Or Trim(basModule.schcd) = "Q" Then       '< 2008.03.24
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
        sStr = sStr & "            ) AS GAEYUL,"
    '<< �迭 >> : 2008.02.15
    ElseIf Trim(basModule.schcd) = "S" Then
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
        sStr = sStr & "                   '22','�����Ư���ڿ�'"
        
        sStr = sStr & "            ) AS GAEYUL,"
    '<< �迭 >> : 2008.02.15
    ElseIf Trim(basModule.schcd) = "P" Then
        sStr = sStr & "     DECODE(KAEYOL,'01','�ι�',"
        sStr = sStr & "                   '02','�ڿ�',"
        sStr = sStr & "                   '03','Ư���ι�',"
        sStr = sStr & "                   '04','Ư���ڿ�'"
        sStr = sStr & "            ) AS GAEYUL,"
    
    ElseIf Trim(basModule.schcd) = "J" Then
        sStr = sStr & "     DECODE(KAEYOL,'01','�ι�',"
        sStr = sStr & "                   '02','�ڿ�',"
        sStr = sStr & "                   '11','�ż��ι�',"
        sStr = sStr & "                   '12','�ż��ڿ�',"
        
        sStr = sStr & "                   '18','�ι������̾�',"
        sStr = sStr & "                   '19','�ڿ������̾�',"
        sStr = sStr & "                   '21','�����Ư���ι�',"
        sStr = sStr & "                   '22','�����Ư���ڿ�'"
        sStr = sStr & "            ) AS GAEYUL,"
        
    Else
        sStr = sStr & "     DECODE(KAEYOL,'01','�ι�',"
        sStr = sStr & "                   '02','�ڿ�'"
        sStr = sStr & "            ) AS GAEYUL,"
    End If
    
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
    sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'" & constSatamCodes(6) & "|') > 0 THEN          /* ��Ž-������� */"
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
    sStr = sStr & " '' AS SEL11,"

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
    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'42|') > 0 THEN '42'"
    
    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL4,'81|') > 0 THEN '81'"
    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL4,'82|') > 0 THEN '82'"
    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL4,'83|') > 0 THEN '83'"
    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL4,'84|') > 0 THEN '84'"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             '00'"
    sStr = sStr & "         END END END END END END END END END END END END END END END END SEL_X2,"
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
    sStr = sStr & "         CASE WHEN INSTR(SEL5,'93|') > 0 THEN         /* �ܱ��� */"          '< ����
    sStr = sStr & "             '93'"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             '00'"
    sStr = sStr & "         END SEL_N3,"
    sStr = sStr & "         CASE WHEN INSTR(SEL5,'94|') > 0 THEN         /*  */"                '< ����
    sStr = sStr & "             '94'"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             '00'"
    sStr = sStr & "         END SEL_N4, "
    sStr = sStr & "         PAYOK, PAYNOT, "
    
' 1������ ���� ó�� : ������ �ڵ��մϴ�. 2007.12.26 ############################################################################################
' ���������� ������ �־�� ��.
    sStr = sStr & "         GET_INTERNET_TCNT_STD_CHASU('" & Trim(basModule.schcd) & "') AS PAYTOT, "       '< ������ ���� ����
    sStr = sStr & "         GET_SUNHANG_P_STD_CHASU('" & Trim(basModule.schcd) & "') AS SUNHANG, "          '< ������ ���� ����
    sStr = sStr & "         GET_SUNUNG_P_STD_CHASU('" & Trim(basModule.schcd) & "') AS SUNUNG, "            '< ������ ���� ����
'###############################################################################################################################################
    'sStr = sStr & "         GET_INTERNET_TOT_STD_INWON('" & Trim(basModule.SchCD) & "') AS PAYTOT, "        '< ��ü���� �ϴ� �Լ�
    
    sStr = sStr & "         K_NUM, M_NUM, E_NUM, TOT_NUM, "
    sStr = sStr & "         ZIP, ADDR1, ADDR2, TEL, CEL, "
    sStr = sStr & "         REGDATE, PAYGBN, IPHAKWONSER, "
    
    Select Case Trim(basModule.schcd)
        Case "S"
            sStr = sStr & " DECODE(PTS_SEL,'1','����','2','����','') AS PTS_SEL, "
        Case "P"
            sStr = sStr & " DECODE(PTS_SEL,'8','����','9','2010 ��','6','3���','') AS PTS_SEL, "
        Case Else
            sStr = sStr & " '' AS PTS_SEL, "
    End Select
    
    sStr = sStr & "         DECODE(MU_TYPE,'1','����','2','6�� �򰡿�','3','9�� �򰡿�','4','6�� �򰡿�','9','���ŵ��','5','9�� �򰡿�','����') AS MU_TYPE "
    
        sStr = sStr & " , "
        sStr = sStr & "        J01,"
        sStr = sStr & "        K01,"
        sStr = sStr & "        J02,"
        sStr = sStr & "        K02,"
        sStr = sStr & "        J03,"
        sStr = sStr & "        K03,"

        sStr = sStr & "        J04,"
        sStr = sStr & "        K04,"
        sStr = sStr & "        J05,"
        sStr = sStr & "        K05,"
        sStr = sStr & "        J06,"
        sStr = sStr & "        K06,"
        sStr = sStr & "        J07,"
        sStr = sStr & "        K07,"
        sStr = sStr & "        J08,"
        sStr = sStr & "        K08,"
        sStr = sStr & "        J09,"
        sStr = sStr & "        K09,"
        sStr = sStr & "        J10,"
        sStr = sStr & "        K10,"
        sStr = sStr & "        J11,"
        sStr = sStr & "        K11,"
        
        sStr = sStr & "        J12,"
        sStr = sStr & "        K12,"
        sStr = sStr & "        J13,"
        sStr = sStr & "        K13,"
        sStr = sStr & "        J14,"
        sStr = sStr & "        K14,"
        
        sStr = sStr & "        J15,"
        sStr = sStr & "        K15,"
        sStr = sStr & "        J16,"
        sStr = sStr & "        K16,"
        sStr = sStr & "        J17,"
        sStr = sStr & "        K17,"
        sStr = sStr & "        J18,"
        sStr = sStr & "        K18,"
        
        sStr = sStr & "        J19,"
        sStr = sStr & "        K19,"
        sStr = sStr & "        J20,"
        sStr = sStr & "        K20,"
        sStr = sStr & "        J21,"
        sStr = sStr & "        K21"
        
    Select Case Trim(Right(cboPassCN, 30))
        Case "ALL"      ' /* �հݻ��� ��� ��ȸ */
            sStr = sStr & " FROM (SELECT A.SCHNO, MAX(EXMID) AS EXMID, MAX(STDNM) AS STDNM, MAX(SEL1_SCH) AS SEL1_SCH, MAX(SEL2_SCH) AS SEL2_SCH, MAX(Birth_ymd) AS Birth_ymd,"
            sStr = sStr & "              MAX(KAEYOL) AS KAEYOL,"
            sStr = sStr & "              MAX(SEL1) AS SEL1, MAX(SEL2) AS SEL2, MAX(SEL3) AS SEL3, MAX(SEL4) AS SEL4, MAX(SEL5) SEL5, "
            sStr = sStr & "              MAX(CL_CLOSE) AS CL_CLOSE, "
            sStr = sStr & "              MAX(PAYOK) AS PAYOK, MAX(PAYNOT) AS PAYNOT, "
            sStr = sStr & "              MAX(K_NUM) AS K_NUM, MAX(M_NUM) AS M_NUM, MAX(E_NUM) AS E_NUM, MAX(TOT_NUM) AS TOT_NUM, "
            sStr = sStr & "              MAX(ZIP) AS ZIP, MAX(ADDR1) AS ADDR1, MAX(ADDR2) AS ADDR2, MAX(TEL) AS TEL, MAX(CEL) AS CEL, "
            sStr = sStr & "              MAX(REGDATE) AS REGDATE, MAX(PAYGBN) AS PAYGBN, MAX(IPHAKWONSER) AS IPHAKWONSER, MAX(PTS_SEL) AS PTS_SEL, MAX(MU_TYPE) AS MU_TYPE "
            
            sStr = sStr & " , "
                    sStr = sStr & "        SUM(J01) AS J01,"
                    sStr = sStr & "        SUM(K01) AS K01,"
                    sStr = sStr & "        SUM(J02) AS J02,"
                    sStr = sStr & "        SUM(K02) AS K02,"
                    sStr = sStr & "        SUM(J03) AS J03,"
                    sStr = sStr & "        SUM(K03) AS K03,"
                    
                    sStr = sStr & "        SUM(J04) AS J04,"
                    sStr = sStr & "        SUM(K04) AS K04,"
                    sStr = sStr & "        SUM(J05) AS J05,"
                    sStr = sStr & "        SUM(K05) AS K05,"
                    sStr = sStr & "        SUM(J06) AS J06,"
                    sStr = sStr & "        SUM(K06) AS K06,"
                    sStr = sStr & "        SUM(J07) AS J07,"
                    sStr = sStr & "        SUM(K07) AS K07,"
                    sStr = sStr & "        SUM(J08) AS J08,"
                    sStr = sStr & "        SUM(K08) AS K08,"
                    sStr = sStr & "        SUM(J09) AS J09,"
                    sStr = sStr & "        SUM(K09) AS K09,"
                    sStr = sStr & "        SUM(J10) AS J10,"
                    sStr = sStr & "        SUM(K10) AS K10,"
                    sStr = sStr & "        SUM(J11) AS J11,"
                    sStr = sStr & "        SUM(K11) AS K11,"
                    
                    sStr = sStr & "        SUM(J12) AS J12,"
                    sStr = sStr & "        SUM(K12) AS K12,"
                    sStr = sStr & "        SUM(J13) AS J13,"
                    sStr = sStr & "        SUM(K13) AS K13,"
                    sStr = sStr & "        SUM(J14) AS J14,"
                    sStr = sStr & "        SUM(K14) AS K14,"
                    
                    sStr = sStr & "        SUM(J15) AS J15,"
                    sStr = sStr & "        SUM(K15) AS K15,"
                    sStr = sStr & "        SUM(J16) AS J16,"
                    sStr = sStr & "        SUM(K16) AS K16,"
                    sStr = sStr & "        SUM(J17) AS J17,"
                    sStr = sStr & "        SUM(K17) AS K17,"
                    sStr = sStr & "        SUM(J18) AS J18,"
                    sStr = sStr & "        SUM(K18) AS K18,"
                    
                    sStr = sStr & "        SUM(J19) AS J19,"
                    sStr = sStr & "        SUM(K19) AS K19,"
                    sStr = sStr & "        SUM(J20) AS J20,"
                    sStr = sStr & "        SUM(K20) AS K20,"
                    sStr = sStr & "        SUM(J21) AS J21,"
                    sStr = sStr & "        SUM(K21) AS K21"
            
            sStr = sStr & "         FROM ("
            '==========================================================================================================
            
            sStr = sStr & "               SELECT SCHNO, EXMID, STDNM, SEL1_SCH, SEL2_SCH, Birth_ymd,"
            sStr = sStr & "                      KAEYOL,"
            sStr = sStr & "                      SEL1 , SEL2, SEL3, SEL4, SEL5, CL_CLOSE, "
            sStr = sStr & "                      PAYOK, PAYNOT, "
            sStr = sStr & "                      NVL(K_NUM,0) AS K_NUM, NVL(M_NUM,0) AS M_NUM, NVL(E_NUM,0) AS E_NUM,"
            sStr = sStr & "                      (NVL(K_NUM,0)+NVL(M_NUM,0)+NVL(E_NUM,0)) AS TOT_NUM ,"
            sStr = sStr & "                      SUBSTR(A.ZIP,1,3)||'-'||SUBSTR(A.ZIP,4,3) AS ZIP, A.ADDR1, A.ADDR2, A.TEL, A.CEL, "
            sStr = sStr & "                      TO_CHAR(A.REGDATE,'YYYY-MM-DD') AS REGDATE, GET_PAYGUBN(A.ORD_NO) AS PAYGBN, DECODE(A.PTS_SEL,'7','���ǰ��','8','����','') AS IPHAKWONSER, PTS_SEL, MU_TYPE "
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
            
            sStr = sStr & "                                WHERE ACID = '" & Trim(basModule.schcd) & "'"
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
                sStr = sStr & "                              AND Birth_ymd LIKE '" & Trim(fpBirth_ymd_F.UnFmtText) & "%'"
            End If
            If Trim(Right(cboSel1_SCH_F.Text, 30)) <> "X" Then
                sStr = sStr & "                              AND SEL1_SCH = '" & Trim(Right(cboSel1_SCH_F.Text, 30)) & "'"
            End If
            If Trim(Right(cboSel2_SCH_F.Text, 30)) <> "X" Then
                sStr = sStr & "                              AND SEL2_SCH = '" & Trim(Right(cboSel2_SCH_F.Text, 30)) & "'"
            End If
            
' 1������ ���� ó�� : ������ �ڵ��մϴ�. 2007.12.26 ############################################################################################
' ���������� ������ �־�� ��.
            Select Case Trim(basModule.schcd)
                Case "N", "S", "P", "J", "B"
                    sStr = sStr & "                          AND TO_CHAR(REGDATE,'YYYYMMDDHH24') >= '" & sChasuTimes & "' "
                Case "K", "W", "Q"
                    sStr = sStr & "                          AND TO_CHAR(REGDATE,'YYYYMMDDHH24') >= '" & sChasuTimes & "' "
            End Select
'###############################################################################################################################################

            
            sStr = sStr & "                                  AND CL_CLOSE IS NULL "
            
            sStr = sStr & "                                  AND BIGO2 IS NULL"                 '< 2008.12. ���ɺ� �л��� �⵵�� ���� �ƴϸ� NULL
            
            sStr = sStr & "                              )"
            sStr = sStr & "                         GROUP BY ACID"
            sStr = sStr & "                      ) B"
            sStr = sStr & "                WHERE A.ACID = B.ACID"
            sStr = sStr & "                  AND A.ACID = '" & Trim(basModule.schcd) & "'"
            
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

' 1������ ���� ó�� : ������ �ڵ��մϴ�. 2007.12.26 ############################################################################################
' ���������� ������ �־�� ��.
            Select Case Trim(basModule.schcd)
                Case "N", "S", "P", "J", "B"
                    sStr = sStr & "          AND TO_CHAR(REGDATE,'YYYYMMDDHH24') >= '" & sChasuTimes & "' "
                Case "K", "W", "Q"
                    sStr = sStr & "          AND TO_CHAR(REGDATE,'YYYYMMDDHH24') >= '" & sChasuTimes & "' "
            End Select
'###############################################################################################################################################
            
            sStr = sStr & "                  AND CL_CLOSE IS NULL "
            
            sStr = sStr & "                  AND BIGO2 IS NULL"                     '< 2008.12. ���ɺ� �л��� �⵵�� ���� �ƴϸ� NULL
            
            sStr = sStr & "               Union All"
            sStr = sStr & "               SELECT SCHNO, EXMID, STDNM, SEL1_SCH, SEL2_SCH, Birth_ymd,"
            sStr = sStr & "                      KAEYOL,"
            sStr = sStr & "                      SEL1 , SEL2, SEL3, SEL4, SEL5, CL_CLOSE, "
            sStr = sStr & "                      0 AS PAYOK, 0 AS PAYNOT, "
            sStr = sStr & "                      0 AS K_NUM, 0 AS M_NUM, 0 AS E_NUM, 0 AS TOT_NUM, "
            sStr = sStr & "                      SUBSTR(ZIP,1,3)||'-'||SUBSTR(ZIP,4,3) AS ZIP, ADDR1, ADDR2, TEL, CEL, "
            sStr = sStr & "                      TO_CHAR(REGDATE,'YYYY-MM-DD') AS REGDATE, GET_PAYGUBN(ORD_NO) AS PAYGBN, DECODE(PTS_SEL,'7','���ǰ��','8','����','') AS IPHAKWONSER, PTS_SEL, MU_TYPE "
            sStr = sStr & "                 From CLSTD01TB"
            sStr = sStr & "                WHERE (PASS1 = '" & Trim(basModule.schcd) & "'" & " OR"
            sStr = sStr & "                       PASS2 = '" & Trim(basModule.schcd) & "'" & " OR"
            sStr = sStr & "                       PASS3 = '" & Trim(basModule.schcd) & "'" & " OR"
            sStr = sStr & "                       PASS4 = '" & Trim(basModule.schcd) & "'" & " )"
            
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
            
' 1������ ���� ó�� : ������ �ڵ��մϴ�. 2007.12.26 ############################################################################################
' ���������� ������ �־�� ��.
            Select Case Trim(basModule.schcd)
                Case "N", "S", "P", "J", "B"
                    sStr = sStr & "          AND TO_CHAR(REGDATE,'YYYYMMDDHH24') >= '" & sChasuTimes & "' "
                Case "K", "W", "Q"
                    sStr = sStr & "          AND TO_CHAR(REGDATE,'YYYYMMDDHH24') >= '" & sChasuTimes & "' "
            End Select
'###############################################################################################################################################

            sStr = sStr & "                  AND CL_CLOSE IS NULL "
            
            sStr = sStr & "                  AND BIGO2 IS NULL"                     '< 2008.12. ���ɺ� �л��� �⵵�� ���� �ƴϸ� NULL
            
            '==========================================================================================================
            
            sStr = sStr & "               ) A,"
            
            sStr = sStr & "               ("
            
                    sStr = sStr & "         SELECT SCHNO,"
                    sStr = sStr & "                DECODE(TRIM(SUB_ID), '37', DECODE(SUB_NUM,'X',0, SUB_NUM), 0)    AS J01,    /* ���                  */"
                    sStr = sStr & "                DECODE(TRIM(SUB_ID), '37', DECODE(SUB_BAK,'X',0, SUB_BAK), 0)    AS K01,    /* �����  ���          */"
                    sStr = sStr & "                DECODE(TRIM(SUB_ID), '38', DECODE(SUB_NUM,'X',0, SUB_NUM), 0)    AS J02,    /* ��������              */"
                    sStr = sStr & "                DECODE(TRIM(SUB_ID), '38', DECODE(SUB_BAK,'X',0, SUB_BAK), 0)    AS K02,    /* �����  ��������      */"
                    sStr = sStr & "                DECODE(TRIM(SUB_ID), '39', DECODE(SUB_NUM,'X',0, SUB_NUM), 0)    AS J03,    /* �ܱ���                */"
                    sStr = sStr & "                DECODE(TRIM(SUB_ID), '39', DECODE(SUB_BAK,'X',0, SUB_BAK), 0)    AS K03,    /* �����  �ܱ���        */"
                    
                    sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(0) & "', DECODE(SUB_NUM,'X',0, SUB_NUM), '51', DECODE(SUB_NUM, 'X',0, SUB_NUM), 0) AS J04,    /* ��Ž-" & constSatams(0) & "        , ��Ž-����1             */"
                    sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(0) & "', DECODE(SUB_BAK,'X',0, SUB_BAK), '51', DECODE(SUB_BAK, 'X',0, SUB_BAK), 0) AS K04,    /* �����  ��Ž-" & constSatams(0) & "        , ��Ž-����1     */"
                    sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(1) & "', DECODE(SUB_NUM,'X',0, SUB_NUM), '52', DECODE(SUB_NUM, 'X',0, SUB_NUM), 0) AS J05,    /* ��Ž-" & constSatams(1) & "         , ��Ž-ȭ��1             */"
                    sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(1) & "', DECODE(SUB_BAK,'X',0, SUB_BAK), '52', DECODE(SUB_BAK, 'X',0, SUB_BAK), 0) AS K05,    /* �����  ��Ž-" & constSatams(1) & "         , ��Ž-ȭ��1     */"
                    sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(2) & "', DECODE(SUB_NUM,'X',0, SUB_NUM), '53', DECODE(SUB_NUM, 'X',0, SUB_NUM), 0) AS J06,    /* ��Ž-" & constSatams(2) & "         , ��Ž-�������1             */"
                    sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(2) & "', DECODE(SUB_BAK,'X',0, SUB_BAK), '53', DECODE(SUB_BAK, 'X',0, SUB_BAK), 0) AS K06,    /* �����  ��Ž-" & constSatams(2) & "         , ��Ž-�������1     */"
                    sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(3) & "', DECODE(SUB_NUM,'X',0, SUB_NUM), '54', DECODE(SUB_NUM, 'X',0, SUB_NUM), 0) AS J07,    /* ��Ž-" & constSatams(3) & "   , ��Ž-��������1         */"
                    sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(3) & "', DECODE(SUB_BAK,'X',0, SUB_BAK), '54', DECODE(SUB_BAK, 'X',0, SUB_BAK), 0) AS K07,    /* �����  ��Ž-" & constSatams(3) & "   , ��Ž-��������1 */"
                    sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(4) & "', DECODE(SUB_NUM,'X',0, SUB_NUM), '55', DECODE(SUB_NUM, 'X',0, SUB_NUM), 0) AS J08,    /* ��Ž-" & constSatams(4) & "       , ��Ž-����2             */"
                    sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(4) & "', DECODE(SUB_BAK,'X',0, SUB_BAK), '55', DECODE(SUB_BAK, 'X',0, SUB_BAK), 0) AS K08,    /* �����  ��Ž-" & constSatams(4) & "       , ��Ž-����2     */"
                    sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(5) & "', DECODE(SUB_NUM,'X',0, SUB_NUM), '56', DECODE(SUB_NUM, 'X',0, SUB_NUM), 0) AS J09,    /* ��Ž-" & constSatams(5) & "     , ��Ž-ȭ��2             */"
                    sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(5) & "', DECODE(SUB_BAK,'X',0, SUB_BAK), '56', DECODE(SUB_BAK, 'X',0, SUB_BAK), 0) AS K09,    /* �����  ��Ž-" & constSatams(5) & "     , ��Ž-ȭ��2     */"
                    sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(6) & "', DECODE(SUB_NUM,'X',0, SUB_NUM), '57', DECODE(SUB_NUM, 'X',0, SUB_NUM), 0) AS J10,      /* ��Ž-" & constSatams(6) & "     , ��Ž-�������2           */"
                    sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(6) & "', DECODE(SUB_BAK,'X',0, SUB_BAK), '57', DECODE(SUB_BAK, 'X',0, SUB_BAK), 0) AS K10,      /* ����� ��Ž-" & constSatams(6) & "     , ��Ž-�������2    */"
                    sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(7) & "', DECODE(SUB_NUM,'X',0, SUB_NUM), '58', DECODE(SUB_NUM, 'X',0, SUB_NUM), 0) AS J11,    /* ��Ž-" & constSatams(7) & "         , ��Ž-��������2         */"
                    sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(7) & "', DECODE(SUB_BAK,'X',0, SUB_BAK), '58', DECODE(SUB_BAK, 'X',0, SUB_BAK), 0) AS K11,    /* �����  ��Ž-" & constSatams(7) & "         , ��Ž-��������2 */"
                    
                    sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(8) & "', DECODE(SUB_NUM,'X',0, SUB_NUM), 0)    AS J12,    /* ��Ž-" & constSatams(8) & "          */"
                    sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(8) & "', DECODE(SUB_BAK,'X',0, SUB_BAK), 0)    AS K12,    /* �����  ��Ž-" & constSatams(8) & "  */"
                    sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(9) & "', DECODE(SUB_NUM,'X',0, SUB_NUM), 0)    AS J13,    /* ��Ž-" & constSatams(9) & "          */"
                    sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(9) & "', DECODE(SUB_BAK,'X',0, SUB_BAK), 0)    AS K13,    /* �����  ��Ž-" & constSatams(9) & "  */"
                    sStr = sStr & " '' AS J14,"
                    sStr = sStr & " '' AS K14,"
                    
                    sStr = sStr & "                DECODE(TRIM(SUB_ID), '31', DECODE(SUB_NUM,'X',0, SUB_NUM), '81', DECODE(SUB_NUM, 'X',0, SUB_NUM), 0) AS J15,    /* ����             , ������                 */"
                    sStr = sStr & "                DECODE(TRIM(SUB_ID), '31', DECODE(SUB_BAK,'X',0, SUB_BAK), '81', DECODE(SUB_BAK, 'X',0, SUB_BAK), 0) AS K15,    /* �����  ����             , ������         */"
                    sStr = sStr & "                DECODE(TRIM(SUB_ID), '32', DECODE(SUB_NUM,'X',0, SUB_NUM), '82', DECODE(SUB_NUM, 'X',0, SUB_NUM), 0) AS J16,    /* �Ͼ�             , �̻����               */"
                    sStr = sStr & "                DECODE(TRIM(SUB_ID), '32', DECODE(SUB_BAK,'X',0, SUB_BAK), '82', DECODE(SUB_BAK, 'X',0, SUB_BAK), 0) AS K16,    /* �����  �Ͼ�             , �̻����       */"
                    sStr = sStr & "                DECODE(TRIM(SUB_ID), '33', DECODE(SUB_NUM,'X',0, SUB_NUM), '83', DECODE(SUB_NUM, 'X',0, SUB_NUM), 0) AS J17,    /* �����ĳ�         , Ȯ�����               */"
                    sStr = sStr & "                DECODE(TRIM(SUB_ID), '33', DECODE(SUB_BAK,'X',0, SUB_BAK), '83', DECODE(SUB_BAK, 'X',0, SUB_BAK), 0) AS K17,    /* �����  �����ĳ�         , Ȯ�����       */"
                    sStr = sStr & "                DECODE(TRIM(SUB_ID), '34', DECODE(SUB_NUM,'X',0, SUB_NUM), '43', DECODE(SUB_NUM, 'X',0, SUB_NUM), 0) AS J18,    /* �Ҿ�             , ��������               */"
                    sStr = sStr & "                DECODE(TRIM(SUB_ID), '34', DECODE(SUB_BAK,'X',0, SUB_BAK), '43', DECODE(SUB_BAK, 'X',0, SUB_BAK), 0) AS K18,    /* �����  �Ҿ�             , ��������       */"
                    
                    sStr = sStr & "                DECODE(TRIM(SUB_ID), '35', DECODE(SUB_NUM,'X',0, SUB_NUM), 0)    AS J19,    /* �߱���                */"
                    sStr = sStr & "                DECODE(TRIM(SUB_ID), '35', DECODE(SUB_BAK,'X',0, SUB_BAK), 0)    AS K19,    /* �����  �߱���        */"
                    sStr = sStr & "                DECODE(TRIM(SUB_ID), '36', DECODE(SUB_NUM,'X',0, SUB_NUM), 0)    AS J20,    /* �ѹ�                  */"
                    sStr = sStr & "                DECODE(TRIM(SUB_ID), '36', DECODE(SUB_BAK,'X',0, SUB_BAK), 0)    AS K20,    /* �����  �ѹ�          */"
                    sStr = sStr & "                DECODE(TRIM(SUB_ID), '42', DECODE(SUB_NUM,'X',0, SUB_NUM), 0)    AS J21,    /* �ƶ���                */"
                    sStr = sStr & "                DECODE(TRIM(SUB_ID), '42', DECODE(SUB_BAK,'X',0, SUB_BAK), 0)    AS K21     /* �����  �ƶ���        */"
                    sStr = sStr & "           FROM CLSTD03TB"
            
            sStr = sStr & "                ) B"

            sStr = sStr & "        WHERE A.SCHNO = B.SCHNO(+)"
            sStr = sStr & "        GROUP BY A.SCHNO"
            sStr = sStr & "       )"
            
            sStr = sStr & "   WHERE SCHNO > ' '"
            
        Case Else       ' /* Ư�� �հ������� �հ��ڸ� ��ȸ�� */
            sStr = sStr & " FROM ("
            
            
            sStr = sStr & "        SELECT SCHNO, EXMID, STDNM, SEL1_SCH, SEL2_SCH, Birth_ymd,"
            sStr = sStr & "              KAEYOL,"
            sStr = sStr & "              SEL1 , SEL2, SEL3, SEL4, SEL5, CL_CLOSE, "
            sStr = sStr & "              0 AS PAYOK , 0 AS PAYNOT, "
            sStr = sStr & "              GET_INTERNET_TOT_STD_INWON('" & Trim(basModule.schcd) & "') AS PAYTOT, "           '< ��ü���� �ϴ� �Լ�
            sStr = sStr & "              TO_CHAR(REGDATE,'YYYY-MM-DD') AS REGDATE, GET_PAYGUBN(ORD_NO) AS PAYGBN, "
            sStr = sStr & "              ZIP, ADDR1, ADDR2, CEL, TEL, "
            sStr = sStr & "              NVL(K_NUM,0) AS K_NUM, NVL(M_NUM,0) AS M_NUM, NVL(E_NUM,0) AS E_NUM,"
            sStr = sStr & "              (NVL(K_NUM,0)+NVL(M_NUM,0)+NVL(E_NUM,0)) AS TOT_NUM , MU_TYPE, DECODE(PTS_SEL,'7','���ǰ��','8','����','') AS IPHAKWONSER"
            sStr = sStr & "         From CLSTD01TB"
            sStr = sStr & "        WHERE PASS" & Trim(Right(cboPassCN, 30)) & " = '" & Trim(basModule.schcd) & "'"
            
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
            
' 1������ ���� ó�� : ������ �ڵ��մϴ�. 2007.12.26 ############################################################################################
' ���������� ������ �־�� ��.
            Select Case Trim(basModule.schcd)
                Case "N", "S", "P", "J", "B"
                    sStr = sStr & "  AND TO_CHAR(REGDATE,'YYYYMMDDHH24') >= '" & sChasuTimes & "' "
                Case "K", "W", "Q"
                    sStr = sStr & "  AND TO_CHAR(REGDATE,'YYYYMMDDHH24') >= '" & sChasuTimes & "' "
            End Select
'###############################################################################################################################################

            sStr = sStr & "          AND CL_CLOSE IS NULL "
            
            sStr = sStr & "          AND BIGO2 IS NULL"                     '< 2008.12. ���ɺ� �л��� �⵵�� ���� �ƴϸ� NULL
            
            sStr = sStr & "      ) A,"
            
                sStr = sStr & "      ("
                            sStr = sStr & " SELECT SCHNO,"
                            sStr = sStr & "        SUM(J01) AS J01,"
                            sStr = sStr & "        SUM(K01) AS K01,"
                            sStr = sStr & "        SUM(J02) AS J02,"
                            sStr = sStr & "        SUM(K02) AS K02,"
                            sStr = sStr & "        SUM(J03) AS J03,"
                            sStr = sStr & "        SUM(K03) AS K03,"
                            sStr = sStr & " "
                            sStr = sStr & "        SUM(J04) AS J04,"
                            sStr = sStr & "        SUM(K04) AS K04,"
                            sStr = sStr & "        SUM(J05) AS J05,"
                            sStr = sStr & "        SUM(K05) AS K05,"
                            sStr = sStr & "        SUM(J06) AS J06,"
                            sStr = sStr & "        SUM(K06) AS K06,"
                            sStr = sStr & "        SUM(J07) AS J07,"
                            sStr = sStr & "        SUM(K07) AS K07,"
                            sStr = sStr & "        SUM(J08) AS J08,"
                            sStr = sStr & "        SUM(K08) AS K08,"
                            sStr = sStr & "        SUM(J09) AS J09,"
                            sStr = sStr & "        SUM(K09) AS K09,"
                            sStr = sStr & "        SUM(J10) AS J10,"
                            sStr = sStr & "        SUM(K10) AS K10,"
                            sStr = sStr & "        SUM(J11) AS J11,"
                            sStr = sStr & "        SUM(K11) AS K11,"
                            sStr = sStr & " "
                            sStr = sStr & "        SUM(J12) AS J12,"
                            sStr = sStr & "        SUM(K12) AS K12,"
                            sStr = sStr & "        SUM(J13) AS J13,"
                            sStr = sStr & "        SUM(K13) AS K13,"
                            sStr = sStr & "        SUM(J14) AS J14,"
                            sStr = sStr & "        SUM(K14) AS K14,"
                            sStr = sStr & " "
                            sStr = sStr & "        SUM(J15) AS J15,"
                            sStr = sStr & "        SUM(K15) AS K15,"
                            sStr = sStr & "        SUM(J16) AS J16,"
                            sStr = sStr & "        SUM(K16) AS K16,"
                            sStr = sStr & "        SUM(J17) AS J17,"
                            sStr = sStr & "        SUM(K17) AS K17,"
                            sStr = sStr & "        SUM(J18) AS J18,"
                            sStr = sStr & "        SUM(K18) AS K18,"
                            sStr = sStr & " "
                            sStr = sStr & "        SUM(J19) AS J19,"
                            sStr = sStr & "        SUM(K19) AS K19,"
                            sStr = sStr & "        SUM(J20) AS J20,"
                            sStr = sStr & "        SUM(K20) AS K20,"
                            sStr = sStr & "        SUM(J21) AS J21,"
                            sStr = sStr & "        SUM(K21) AS K21"
                            sStr = sStr & "   FROM ("
                            sStr = sStr & "         SELECT SCHNO,"
                            sStr = sStr & "                DECODE(TRIM(SUB_ID), '37', DECODE(SUB_NUM,'X',0, SUB_NUM), 0)    AS J01,    /* ���                  */"
                            sStr = sStr & "                DECODE(TRIM(SUB_ID), '37', DECODE(SUB_BAK,'X',0, SUB_BAK), 0)    AS K01,    /* �����  ���          */"
                            sStr = sStr & "                DECODE(TRIM(SUB_ID), '38', DECODE(SUB_NUM,'X',0, SUB_NUM), 0)    AS J02,    /* ��������              */"
                            sStr = sStr & "                DECODE(TRIM(SUB_ID), '38', DECODE(SUB_BAK,'X',0, SUB_BAK), 0)    AS K02,    /* �����  ��������      */"
                            sStr = sStr & "                DECODE(TRIM(SUB_ID), '39', DECODE(SUB_NUM,'X',0, SUB_NUM), 0)    AS J03,    /* �ܱ���                */"
                            sStr = sStr & "                DECODE(TRIM(SUB_ID), '39', DECODE(SUB_BAK,'X',0, SUB_BAK), 0)    AS K03,    /* �����  �ܱ���        */"
                            sStr = sStr & " "
                            sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(0) & "', DECODE(SUB_NUM,'X',0, SUB_NUM), '51', DECODE(SUB_NUM, 'X',0, SUB_NUM), 0) AS J04,    /* ��Ž-" & constSatams(0) & "        , ��Ž-����1             */"
                            sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(0) & "', DECODE(SUB_BAK,'X',0, SUB_BAK), '51', DECODE(SUB_BAK, 'X',0, SUB_BAK), 0) AS K04,    /* �����  ��Ž-" & constSatams(0) & "        , ��Ž-����1     */"
                            sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(1) & "', DECODE(SUB_NUM,'X',0, SUB_NUM), '52', DECODE(SUB_NUM, 'X',0, SUB_NUM), 0) AS J05,    /* ��Ž-" & constSatams(1) & "         , ��Ž-ȭ��1             */"
                            sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(1) & "', DECODE(SUB_BAK,'X',0, SUB_BAK), '52', DECODE(SUB_BAK, 'X',0, SUB_BAK), 0) AS K05,    /* �����  ��Ž-" & constSatams(1) & "         , ��Ž-ȭ��1     */"
                            sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(2) & "', DECODE(SUB_NUM,'X',0, SUB_NUM), '53', DECODE(SUB_NUM, 'X',0, SUB_NUM), 0) AS J06,    /* ��Ž-" & constSatams(2) & "         , ��Ž-�������1             */"
                            sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(2) & "', DECODE(SUB_BAK,'X',0, SUB_BAK), '53', DECODE(SUB_BAK, 'X',0, SUB_BAK), 0) AS K06,    /* �����  ��Ž-" & constSatams(2) & "         , ��Ž-�������1     */"
                            sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(3) & "', DECODE(SUB_NUM,'X',0, SUB_NUM), '54', DECODE(SUB_NUM, 'X',0, SUB_NUM), 0) AS J07,    /* ��Ž-" & constSatams(3) & "   , ��Ž-��������1         */"
                            sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(3) & "', DECODE(SUB_BAK,'X',0, SUB_BAK), '54', DECODE(SUB_BAK, 'X',0, SUB_BAK), 0) AS K07,    /* �����  ��Ž-" & constSatams(3) & "   , ��Ž-��������1 */"
                            sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(4) & "', DECODE(SUB_NUM,'X',0, SUB_NUM), '55', DECODE(SUB_NUM, 'X',0, SUB_NUM), 0) AS J08,    /* ��Ž-" & constSatams(4) & "       , ��Ž-����2             */"
                            sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(4) & "', DECODE(SUB_BAK,'X',0, SUB_BAK), '55', DECODE(SUB_BAK, 'X',0, SUB_BAK), 0) AS K08,    /* �����  ��Ž-" & constSatams(4) & "       , ��Ž-����2     */"
                            sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(5) & "', DECODE(SUB_NUM,'X',0, SUB_NUM), '56', DECODE(SUB_NUM, 'X',0, SUB_NUM), 0) AS J09,    /* ��Ž-" & constSatams(5) & "     , ��Ž-ȭ��2             */"
                            sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(5) & "', DECODE(SUB_BAK,'X',0, SUB_BAK), '56', DECODE(SUB_BAK, 'X',0, SUB_BAK), 0) AS K09,    /* �����  ��Ž-" & constSatams(5) & "     , ��Ž-ȭ��2     */"
                            sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(6) & "', DECODE(SUB_NUM,'X',0, SUB_NUM), '57', DECODE(SUB_NUM, 'X',0, SUB_NUM), 0) AS J10,      /* ��Ž-" & constSatams(6) & "     , ��Ž-�������2           */"
                            sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(6) & "', DECODE(SUB_BAK,'X',0, SUB_BAK), '57', DECODE(SUB_BAK, 'X',0, SUB_BAK), 0) AS K10,      /* ����� ��Ž-" & constSatams(6) & "     , ��Ž-�������2    */"
                            sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(7) & "', DECODE(SUB_NUM,'X',0, SUB_NUM), '58', DECODE(SUB_NUM, 'X',0, SUB_NUM), 0) AS J11,    /* ��Ž-" & constSatams(7) & "         , ��Ž-��������2         */"
                            sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(7) & "', DECODE(SUB_BAK,'X',0, SUB_BAK), '58', DECODE(SUB_BAK, 'X',0, SUB_BAK), 0) AS K11,    /* �����  ��Ž-" & constSatams(7) & "         , ��Ž-��������2 */"
                            
                            sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(8) & "', DECODE(SUB_NUM,'X',0, SUB_NUM), 0)    AS J12,    /* ��Ž-" & constSatams(8) & "          */"
                            sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(8) & "', DECODE(SUB_BAK,'X',0, SUB_BAK), 0)    AS K12,    /* �����  ��Ž-" & constSatams(8) & "  */"
                            sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(9) & "', DECODE(SUB_NUM,'X',0, SUB_NUM), 0)    AS J13,    /* ��Ž-" & constSatams(9) & "          */"
                            sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(9) & "', DECODE(SUB_BAK,'X',0, SUB_BAK), 0)    AS K13,    /* �����  ��Ž-" & constSatams(9) & "  */"
                            sStr = sStr & " '' AS J14,"
                            sStr = sStr & " '' AS K14,"
                            sStr = sStr & " "
                            sStr = sStr & "                DECODE(TRIM(SUB_ID), '31', DECODE(SUB_NUM,'X',0, SUB_NUM), '81', DECODE(SUB_NUM, 'X',0, SUB_NUM), 0) AS J15,    /* ����             , ������                 */"
                            sStr = sStr & "                DECODE(TRIM(SUB_ID), '31', DECODE(SUB_BAK,'X',0, SUB_BAK), '81', DECODE(SUB_BAK, 'X',0, SUB_BAK), 0) AS K15,    /* �����  ����             , ������         */"
                            sStr = sStr & "                DECODE(TRIM(SUB_ID), '32', DECODE(SUB_NUM,'X',0, SUB_NUM), '82', DECODE(SUB_NUM, 'X',0, SUB_NUM), 0) AS J16,    /* �Ͼ�             , �̻����               */"
                            sStr = sStr & "                DECODE(TRIM(SUB_ID), '32', DECODE(SUB_BAK,'X',0, SUB_BAK), '82', DECODE(SUB_BAK, 'X',0, SUB_BAK), 0) AS K16,    /* �����  �Ͼ�             , �̻����       */"
                            sStr = sStr & "                DECODE(TRIM(SUB_ID), '33', DECODE(SUB_NUM,'X',0, SUB_NUM), '83', DECODE(SUB_NUM, 'X',0, SUB_NUM), 0) AS J17,    /* �����ĳ�         , Ȯ�����               */"
                            sStr = sStr & "                DECODE(TRIM(SUB_ID), '33', DECODE(SUB_BAK,'X',0, SUB_BAK), '83', DECODE(SUB_BAK, 'X',0, SUB_BAK), 0) AS K17,    /* �����  �����ĳ�         , Ȯ�����       */"
                            sStr = sStr & "                DECODE(TRIM(SUB_ID), '34', DECODE(SUB_NUM,'X',0, SUB_NUM), '43', DECODE(SUB_NUM, 'X',0, SUB_NUM), 0) AS J18,    /* �Ҿ�             , ��������               */"
                            sStr = sStr & "                DECODE(TRIM(SUB_ID), '34', DECODE(SUB_BAK,'X',0, SUB_BAK), '43', DECODE(SUB_BAK, 'X',0, SUB_BAK), 0) AS K18,    /* �����  �Ҿ�             , ��������       */"
                            sStr = sStr & " "
                            sStr = sStr & "                DECODE(TRIM(SUB_ID), '35', DECODE(SUB_NUM,'X',0, SUB_NUM), 0)    AS J19,    /* �߱���                */"
                            sStr = sStr & "                DECODE(TRIM(SUB_ID), '35', DECODE(SUB_BAK,'X',0, SUB_BAK), 0)    AS K19,    /* �����  �߱���        */"
                            sStr = sStr & "                DECODE(TRIM(SUB_ID), '36', DECODE(SUB_NUM,'X',0, SUB_NUM), 0)    AS J20,    /* �ѹ�                  */"
                            sStr = sStr & "                DECODE(TRIM(SUB_ID), '36', DECODE(SUB_BAK,'X',0, SUB_BAK), 0)    AS K20,    /* �����  �ѹ�          */"
                            sStr = sStr & "                DECODE(TRIM(SUB_ID), '42', DECODE(SUB_NUM,'X',0, SUB_NUM), 0)    AS J21,    /* �ƶ���                */"
                            sStr = sStr & "                DECODE(TRIM(SUB_ID), '42', DECODE(SUB_BAK,'X',0, SUB_BAK), 0)    AS K21     /* �����  �ƶ���        */"
                            sStr = sStr & "           FROM CLSTD03TB"
                                      
                            sStr = sStr & "         )"
                            sStr = sStr & "  GROUP BY SCHNO"
        sStr = sStr & "             ) B"
            sStr = sStr & "  WHERE A.SCHNO = B.SCHNO(+)"
            
            
    End Select
    
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
        sStr = sStr & " AND Birth_ymd LIKE '" & Trim(fpBirth_ymd_F.UnFmtText) & "%'"
    End If
    If Trim(Right(cboSel1_SCH_F.Text, 30)) <> "X" Then
        sStr = sStr & " AND SEL1_SCH = '" & Trim(Right(cboSel1_SCH_F.Text, 30)) & "'"
    End If
    If Trim(Right(cboSel2_SCH_F.Text, 30)) <> "X" Then
        sStr = sStr & " AND SEL2_SCH = '" & Trim(Right(cboSel2_SCH_F.Text, 30)) & "'"
    End If
    
    sStr = sStr & "     AND CL_CLOSE IS NULL "
    
    sStr = sStr & "   ORDER BY EXMID "
    
    
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
'    '>> �������
'        If Trim(fpBirth_ymd_F.UnFmtText) > "" Then
'            sTmp = "%" & Trim(fpBirth_ymd_F.UnFmtText) & "%"
'            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'                Set DBParam = DBCmd.CreateParameter("Birth_ymd", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
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
                        
                    nTmp = 0:       If IsNumeric(.Fields("SUNHANG")) = True Then nTmp = .Fields("SUNHANG")
                        fpSunHang.value = nTmp
                    nTmp = 0:       If IsNumeric(.Fields("SUNUNG")) = True Then nTmp = .Fields("SUNUNG")
                        fpSuNung.value = nTmp
                        
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
                    sTmp = " ":
                    If IsNull(.Fields("SEL1_SCH")) = False Then
                        Select Case Trim(.Fields("SEL1_SCH"))
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
                                
                        End Select
                    End If
                    Call basFunction.Set_SprType_Text(sprSTD_F, "CENTER", "LEFT", LenB(sTmp), sTmp)
                
                
                sprSTD_F.Col = 5
                    sTmp = " "
                    If IsNull(.Fields("SEL2_SCH")) = False Then
                        Select Case Trim(.Fields("SEL2_SCH"))
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
                                
                            Case Else
                                sTmp = ""
                        End Select
                    End If
                    Call basFunction.Set_SprType_Text(sprSTD_F, "CENTER", "LEFT", LenB(sTmp), sTmp)
                
                
                sprSTD_F.Col = 6
                    sTmp = " ":  If IsNull(.Fields("Birth_ymd")) = False Then sTmp = Trim(.Fields("Birth_ymd"))
                    Call basFunction.Set_SprType_Text(sprSTD_F, "CENTER", "LEFT", LenB(sTmp), sTmp)
                
                sprSTD_F.SetCellBorder sprSTD_F.Col, sprSTD_F.Row, sprSTD_F.Col, sprSTD_F.Row, 2, basModule.SectionColor2, CellBorderStyleSolid
                
                
                sprSTD_F.Col = 7
                    nTmp = 0:   If IsNumeric(.Fields("K_NUM")) = True Then nTmp = CDbl(Trim(.Fields("K_NUM")))
                    Call basFunction.Set_SprType_Numeric(sprSTD_F, 0, 0, 999999, "", nTmp)
                sprSTD_F.Col = 8
                    nTmp = 0:   If IsNumeric(.Fields("M_NUM")) = True Then nTmp = CDbl(Trim(.Fields("M_NUM")))
                    Call basFunction.Set_SprType_Numeric(sprSTD_F, 0, 0, 999999, "", nTmp)
                sprSTD_F.Col = 9
                    nTmp = 0:   If IsNumeric(.Fields("E_NUM")) = True Then nTmp = CDbl(Trim(.Fields("E_NUM")))
                    Call basFunction.Set_SprType_Numeric(sprSTD_F, 0, 0, 999999, "", nTmp)
                sprSTD_F.Col = 10
                    nTmp = 0:   If IsNumeric(.Fields("TOT_NUM")) = True Then nTmp = CDbl(Trim(.Fields("TOT_NUM")))
                    Call basFunction.Set_SprType_Numeric(sprSTD_F, 0, 0, 999999, "", nTmp)
                
                sprSTD_F.Col = 11
                    sTmp = " ":  If IsNull(.Fields("GAEYUL")) = False Then sTmp = Trim(.Fields("GAEYUL"))
                    Call basFunction.Set_SprType_Text(sprSTD_F, "CENTER", "LEFT", LenB(sTmp), sTmp)
                
                sprSTD_F.SetCellBorder sprSTD_F.Col, sprSTD_F.Row, sprSTD_F.Col, sprSTD_F.Row, 2, basModule.SectionColor2, CellBorderStyleSolid
                
                '>> ���ð��� (��Ž/ ��Ž)
                For ni = 1 To SATAM_COUNT Step 1
                
                    If ni Mod 4 = 1 Then
                        sprSTD_F.SetCellBorder sprSTD_F.Col, sprSTD_F.Row, sprSTD_F.Col, sprSTD_F.Row, 2, basModule.SectionColor2, CellBorderStyleSolid
                    End If
                
                    sprSTD_F.Col = sprSTD_F.Col + 1
                    
                    Select Case ni
                        Case 1 To 8
                            sGbn = "SEL" & Trim(CStr(ni))
                        Case 9 To 11
                            If sKaeyol = "02" Or sKaeyol = "04" Or sKaeyol = "06" Then
                                sGbn = "X"
                            Else
                                sGbn = "SEL" & Trim(CStr(ni))
                            End If
                    End Select
                    
                    If sGbn = "X" Then
                        Call basFunction.Set_SprType_Text(sprSTD_F, "CENTER", "LEFT", 10, "")
                    Else
                        sTmp = IIf(Trim(.Fields(sGbn)) = "00", "", Trim(.Fields(sGbn)))
                        
                        If IsNull(.Fields(sGbn)) = False Then
                            If sTmp <> "" Then
                                Select Case sTmp
                                    Case constSatamCodes(0):  sTmp = constSatams(0)
                                    Case constSatamCodes(1):  sTmp = constSatams(1)
                                    Case constSatamCodes(2):  sTmp = constSatams(2)
                                    Case constSatamCodes(3):  sTmp = constSatams(3)
                                    Case constSatamCodes(4):  sTmp = constSatams(4)
                                    Case constSatamCodes(5):  sTmp = constSatams(5)
                                    Case constSatamCodes(6):  sTmp = constSatams(6)
                                    Case constSatamCodes(7):  sTmp = constSatams(7)
                                    Case constSatamCodes(8):  sTmp = constSatams(8)
                                    Case constSatamCodes(9):  sTmp = constSatams(9)
                                    
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
                            Call basFunction.Set_SprType_Text(sprSTD_F, "CENTER", "LEFT", LenB(sTmp), sTmp)
                        End If
                    End If
                Next ni
                
                '��Ž�����ϳ� �ٸ鼭 ��ĭ���� ó��
                sprSTD_F.Col = sprSTD_F.Col + 1
                Call basFunction.Set_SprType_Text(sprSTD_F, "CENTER", "LEFT", LenB(""), "")
                
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
                            Case "42":  sTmp = "�ƶ���"
                            Case "44":  sTmp = "��Ʈ����"
                            
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
                                    Case "93":  sTmp = "�ܱ���"     '< ����
                                    Case "94":  sTmp = ""           '< ����
                                    
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
                    
                sprSTD_F.Col = sprSTD_F.Col + 1     '< ���� �߰�
                    sTmp = " ":  If IsNull(.Fields("IPHAKWONSER")) = False Then sTmp = Trim(.Fields("IPHAKWONSER"))
                    Call basFunction.Set_SprType_Text(sprSTD_F, "CENTER", "LEFT", LenB(sTmp), sTmp)
                    
                sprSTD_F.Col = sprSTD_F.Col + 1     '< ���� �߰�
                    sTmp = " ":  If IsNull(.Fields("PTS_SEL")) = False Then sTmp = Trim(.Fields("PTS_SEL"))
                    Call basFunction.Set_SprType_Text(sprSTD_F, "CENTER", "LEFT", LenB(sTmp), sTmp)
                    
                sprSTD_F.Col = sprSTD_F.Col + 1     '< ���� �߰�
                    sTmp = " ":  If IsNull(.Fields("MU_TYPE")) = False Then sTmp = Trim(.Fields("MU_TYPE"))
                    Call basFunction.Set_SprType_Text(sprSTD_F, "CENTER", "LEFT", LenB(sTmp), sTmp)
                
                
                For nj = 1 To (nJumsu / 2) Step 1
                    
                    sJTmp = "J" & Trim(Format(CStr(nj), "00"))
                        sprSTD_F.Col = sprSTD_F.Col + 1
                            sTmp = "0":  If IsNull(.Fields(sJTmp)) = False Then sTmp = Trim(.Fields(sJTmp))
                            If sTmp <> "0" Then Call basFunction.Set_SprType_Numeric(sprSTD_F, 0, 0, 99999, "", CInt(sTmp))
                    
                    sJTmp = "K" & Trim(Format(CStr(nj), "00"))
                        sprSTD_F.Col = sprSTD_F.Col + 1
                            sTmp = "0":  If IsNull(.Fields(sJTmp)) = False Then sTmp = Trim(.Fields(sJTmp))
                            If sTmp <> "0" Then Call basFunction.Set_SprType_Numeric(sprSTD_F, 0, 0, 99999, "", CInt(sTmp))
                            
                Next nj
                
                
                .MoveNext
            Next nRec
            
            sprSTD_F.Row = 1:       sprSTD_F.Row2 = sprSTD_F.MaxRows
            sprSTD_F.Col = 1:       sprSTD_F.Col2 = sprSTD_F.MaxCols
            sprSTD_F.BlockMode = True
                sprSTD_F.BackColor = basModule.WhiteColor
                sprSTD_F.BackColorStyle = BackColorStyleUnderGrid
            sprSTD_F.BlockMode = False
            
            sprSTD_F.ColsFrozen = 3
            
        End If
    End With
    
    MsgBox "�л� ��ȸ�Ͽ����ϴ�.", vbInformation + vbOKOnly, "�л���ȸ"
    
    sprSTD_F.SetFocus
    
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
                    .BackColor = basModule.WhiteColor
                    .BackColorStyle = BackColorStyleUnderGrid
                .BlockMode = False
                
'                .Row = .ActiveRow
'                .Col = 1
'                    Call Show_Select_STD(Trim(.Text))
                
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
                '.SetActiveCell .ActiveCol, .ActiveRow
                
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
                .BackColor = basModule.WhiteColor
                .BackColorStyle = BackColorStyleUnderGrid
            .BlockMode = False
            
'            .Row = Row
'            .Col = 1
'                Call Show_Select_STD(Trim(.Text))
            
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
        'sprSTD_F.SetActiveCell Col, Row
        
    End With
    
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
    
End Sub

Private Sub Get_Excel_Data()

    Dim sPath       As String
    
    ' Excel Data ó��
    Dim xlsDBConn   As ADODB.Connection
    Dim DBExCmd     As ADODB.Command
    Dim DBExRec     As ADODB.Recordset
    
    Dim sConn       As String
    Dim sSql        As String
    
    Dim nRow        As Long
    Dim sTmp        As String
    Dim nTmp        As Long
    
    Dim nJumsu      As Long
    Dim ni          As Long
    Dim nC          As Long
    
    On Error GoTo ErrStmt1
    
    With dlgFile
        .CancelError = True
        .fileName = ""
        .InitDir = App.Path
        .Filter = "EXCEL FILE(*.XLS)|*.XLS"
        .DefaultExt = "*.XLS"
        .ShowOpen
        
        If (.fileName) = "" Then
            MsgBox "������ ������ �����ϴ�.", vbExclamation + vbOKOnly, Me.Caption
            Exit Sub
        End If
        
        sPath = .fileName
        
    End With
    
    On Error GoTo 0
    
    On Error GoTo ErrStmt2                          '>> error ó��
    
    Set xlsDBConn = New ADODB.Connection
    sConn = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
            "Data Source=" & sPath & ";" & _
            "Extended Properties=""Excel 8.0;HDR=no;"";"
    
    With xlsDBConn
        .ConnectionString = sConn                   ' �����ͺ��̽��� ������ �õ��մϴ�.
        .ConnectionTimeout = 30                     ' ���� �ð����� ������ ���� ������ �ڵ����� �����ϴ�.
        .Properties("Prompt") = adPromptNever       ' �̰��� ADO���� �⺻ ������Ʈ ����Դϴ�.
        .CursorLocation = adUseClient               ' Ŀ����ġ�� Client �ʿ� �ֽ��ϴ�.
        
        .Open                                       ' �����ͺ��̽��� ���ϴ�.
        
        Do While .State And adStateConnecting
            DoEvents
        Loop
    End With
       
'>> ���� DB Open
    sSql = ""
    sSql = sSql & " SELECT * "
    sSql = sSql & "   FROM [Sheet1$] "
    
    Set DBExCmd = New ADODB.Command
    Set DBExRec = New ADODB.Recordset
    
    DBExCmd.ActiveConnection = xlsDBConn
    DBExCmd.CommandText = sSql
    DBExCmd.CommandType = adCmdText
    DBExCmd.CommandTimeout = 30
    
    DBExRec.Open DBExCmd, , adOpenStatic, adLockReadOnly, -1
    Do While xlsDBConn.State And adStateExecuting
        DoEvents
    Loop
    
    If DBExRec.RecordCount = 0 Then
        Set DBExCmd = Nothing
        Set DBExRec = Nothing
        Set xlsDBConn = Nothing
        
        MsgBox "Excel Data�� �����ϴ�.", vbExclamation + vbOKOnly, "IT2007"
        Exit Sub
    End If
        
    
    sprExcel_STD_Data.MaxRows = 0       ' �ʱ�ȭ
    
    
    DBExRec.MoveFirst
        
    '## header 1 line skip
    DBExRec.MoveNext
    
    
    For nRow = 2 To DBExRec.RecordCount Step 1
    '�п��ڵ�
        sTmp = "":  If IsNull(DBExRec.Fields(0)) = False Then sTmp = UCase(Trim(DBExRec.Fields(0)))
        uExcel_StdData.ACID = sTmp
    '�����ȣ
        sTmp = "":  If IsNull(DBExRec.Fields(1)) = False Then sTmp = Trim(DBExRec.Fields(1))
        uExcel_StdData.EXMID = sTmp
    '�л���
        sTmp = "":  If IsNull(DBExRec.Fields(2)) = False Then sTmp = Trim(DBExRec.Fields(2))
        uExcel_StdData.STDNM = sTmp
    '�������
        sTmp = "":  If IsNull(DBExRec.Fields(3)) = False Then sTmp = Trim(DBExRec.Fields(3))
        sTmp = Replace(sTmp, "-", "", 1, -1, vbTextCompare)
        If basFunction.LenKor(sTmp) > 6 Then
            sTmp = Left(sTmp, 4) & "-" & Mid(sTmp, 5, 2) & "-" & Mid(sTmp, 7, 2)
        End If
        uExcel_StdData.Birth_ymd = sTmp
    '��.������
        sTmp = "1"
        If IsNull(DBExRec.Fields(4)) = False Then
            sTmp = UCase(Trim(DBExRec.Fields(4)))
            Select Case sTmp
                Case "0", "1"
                    'no action
                Case Else
                    sTmp = "1"
                    
            End Select
        End If
        uExcel_StdData.EXMTYPE = sTmp
    '�迭
        sTmp = "01"
        If Trim(basModule.schcd) = "N" Then             '< �迭 : 2008.01.09 - �뷮��
            If IsNull(DBExRec.Fields(5)) = False Then
                sTmp = UCase(Trim(DBExRec.Fields(5)))
                Select Case sTmp
                    Case "1" To "9"
                        sTmp = Format(sTmp, "00")
                    Case "�ι�", "��"
                        sTmp = "01"
                    Case "�ڿ�", "��"
                        sTmp = "02"
                    Case "��ü", "��"
                        sTmp = "03"
                    
                    Case "����(��)", "������"
                        sTmp = "04"
                    Case "�ι�����", "�����ι�"
                        sTmp = "05"
                    Case "�ڿ�����", "�����ڿ�"
                        sTmp = "06"
                        
                    Case "�ż��ι�"
                        sTmp = "07"
                    Case "�ż��ڿ�"
                        sTmp = "08"
'                    Case "�ż������ι�"
'                        sTmp = "09"
'                    Case "�ż������ڿ�"
'                        sTmp = "10"
                    
                    Case "�����ι�", "����"
                            sTmp = "11"
                    Case "�����ڿ�", "����"
                        sTmp = "12"
                    Case "��ü", "��"
                        sTmp = "13"
                    
                    Case "�����(��)", "�������"
                        sTmp = "14"
                    Case "���ι�����", "������ι�"
                        sTmp = "15"
                    Case "���ڿ�����", "������ڿ�"
                        sTmp = "16"
                    
                    Case Else
                        sTmp = "01"
                End Select
            End If
        ElseIf Trim(basModule.schcd) = "K" Or Trim(basModule.schcd) = "W" Or Trim(basModule.schcd) = "Q" Then        '< �迭 : 2008.01.10 - ����, 2008.03.24
            If IsNull(DBExRec.Fields(5)) = False Then
                sTmp = UCase(Trim(DBExRec.Fields(5)))
                Select Case sTmp
                    Case "1" To "9"
                        sTmp = Format(sTmp, "00")
                    Case "�ι�", "��"
                        sTmp = "01"
                    Case "�ڿ�", "��"
                        sTmp = "02"
                    
                    Case "�ְ�����", "�ֹ�"
                        sTmp = "04"
                    Case "�ְ��Ǵ�", "����"
                        sTmp = "05"
                    
                    Case "�߰�����", "�߹�"
                        sTmp = "06"
                    Case "�߰��Ǵ�", "����"
                        sTmp = "07"
                    
                    Case "�������ι�"
                        sTmp = "11"
                    Case "�������ڿ�"
                        sTmp = "12"
                        
                    Case "�������ι�16"
                        sTmp = "16"
                    Case "�������ڿ�17"
                        sTmp = "17"
                        
                    Case Else
                        sTmp = "01"
                End Select
            End If
        ElseIf Trim(basModule.schcd) = "S" Then             '< �迭 : 2008.02.15 - ����
            If IsNull(DBExRec.Fields(5)) = False Then
                sTmp = UCase(Trim(DBExRec.Fields(5)))
                Select Case sTmp
                    Case "1" To "9"
                        sTmp = Format(sTmp, "00")
                    Case "�ι�", "��"
                        sTmp = "01"
                    Case "�ڿ�", "��"
                        sTmp = "02"
                    
                    Case "Ư��", "Ư���ι�"
                        sTmp = "03"
                    Case "Ư��", "Ư���ڿ�"
                        sTmp = "04"
                        
                    Case "�����ι�"
                        sTmp = "05"
                    Case "�����ڿ�"
                        sTmp = "06"
                    Case "��������"
                        sTmp = "08"
                        
                    Case "�ż��ι�"
                        sTmp = "11"
                    Case "�ż��ڿ�"
                        sTmp = "12"
                        
                    Case "�ι������̾�"
                        sTmp = "18"
                    Case "�ڿ������̾�"
                        sTmp = "19"
                        
                    Case Else
                        sTmp = "01"
                End Select
            End If
        ElseIf Trim(basModule.schcd) = "P" Then             '< �迭 : 2008.02.15 - ����
            If IsNull(DBExRec.Fields(5)) = False Then
                sTmp = UCase(Trim(DBExRec.Fields(5)))
                Select Case sTmp
                    Case "1" To "9"
                        sTmp = Format(sTmp, "00")
                    Case "�ι�", "��"
                        sTmp = "01"
                    Case "�ڿ�", "��"
                        sTmp = "02"
                    
                    Case "Ư��", "Ư���ι�"
                        sTmp = "03"
                    Case "Ư��", "Ư���ڿ�"
                        sTmp = "04"
                        
                    Case Else
                        sTmp = "01"
                End Select
            End If
        
        ElseIf Trim(basModule.schcd) = "J" Then             '< ����
            If IsNull(DBExRec.Fields(5)) = False Then
                sTmp = UCase(Trim(DBExRec.Fields(5)))
                Select Case sTmp
                    Case "1" To "9"
                        sTmp = Format(sTmp, "00")
                    Case "�ι�", "��"
                        sTmp = "01"
                    Case "�ڿ�", "��"
                        sTmp = "02"
                    
                    Case "�ż��ι�"
                        sTmp = "11"
                    Case "�ż��ڿ�"
                        sTmp = "12"
                    
                    Case "�ι������̾�"
                        sTmp = "18"
                    Case "�ڿ������̾�"
                        sTmp = "19"
                        
                    Case Else
                        sTmp = "01"
                End Select
            End If
            
        Else
            If IsNull(DBExRec.Fields(5)) = False Then
                sTmp = UCase(Trim(DBExRec.Fields(5)))
                Select Case sTmp
                    Case "1" To "9"
                        sTmp = Format(sTmp, "00")
                    Case "�ι�", "��"
                        sTmp = "01"
                    Case "�ڿ�", "��"
                        sTmp = "02"
                    Case "��ü", "��"
                        sTmp = "03"
                    Case Else
                        sTmp = "01"
                End Select
            End If
        End If
        uExcel_StdData.kaeyol = sTmp
        
    '1 �����п�
        sTmp = Trim(basModule.schcd)
        If IsNull(DBExRec.Fields(6)) = False Then
            sTmp = UCase(Trim(DBExRec.Fields(6)))
            Select Case sTmp
                Case "N", "K", "S", "P", "M", "W", "Q", "J", "B"
                    ' NEXT
                Case "�뷮��"
                    sTmp = "N"
                Case "����"
                    sTmp = "K"
                Case "����"
                    sTmp = "S"
                Case "����M", "���ĸ��̸�", "���� MIMAC", "����MIMAC", "����"
                    sTmp = "P"
                Case "����M", "�������̸�", "���� MIMAC", "����MIMAC", "����"
                    sTmp = "M"
                    
                Case "�ָ����Ǵ�", "�ָ���", "�ֹ�"
                    sTmp = "W"
                Case "�߰����Ǵ�", "�߰���", "�߹�"
                    sTmp = "Q"
                    
                Case "����"
                    sTmp = "J"
                Case "�λ�"
                    sTmp = "B"
                
                Case Else
                    sTmp = Trim(basModule.schcd)
            End Select
        End If
        uExcel_StdData.WANT_ACID1 = sTmp
        
    '2 �����п�
        sTmp = Trim(basModule.schcd)
        If IsNull(DBExRec.Fields(7)) = False Then
            sTmp = UCase(Trim(DBExRec.Fields(7)))
            Select Case sTmp
                Case "N", "K", "S", "P", "M", "W", "Q", "J", "B"
                    ' NEXT
                Case "�뷮��"
                    sTmp = "N"
                Case "����"
                    sTmp = "K"
                Case "����"
                    sTmp = "S"
                Case "����M", "���ĸ��̸�", "���� MIMAC", "����MIMAC", "����"
                    sTmp = "P"
                Case "����M", "�������̸�", "���� MIMAC", "����MIMAC", "����"
                    sTmp = "M"
                    
                Case "�ָ����Ǵ�", "�ָ���", "�ֹ�"
                    sTmp = "W"
                Case "�߰����Ǵ�", "�߰���", "�߹�"
                    sTmp = "Q"
                    
                Case "����"
                    sTmp = "J"
                Case "�λ�"
                    sTmp = "B"
                    
                Case Else
                    sTmp = Trim(basModule.schcd)
            End Select
        End If
        uExcel_StdData.WANT_ACID2 = sTmp
        
    '����
        nTmp = 0:  If IsNumeric(DBExRec.Fields(8)) = True Then nTmp = CLng(Trim(DBExRec.Fields(8)))
        uExcel_StdData.KOR = nTmp
    '����
        nTmp = 0:  If IsNumeric(DBExRec.Fields(9)) = True Then nTmp = CLng(Trim(DBExRec.Fields(9)))
        uExcel_StdData.ENG = nTmp
    '����
        nTmp = 0:  If IsNumeric(DBExRec.Fields(10)) = True Then nTmp = CLng(Trim(DBExRec.Fields(10)))
        uExcel_StdData.MAT = nTmp
        
    '��Ž
        uExcel_StdData.SATAM1 = ""
        uExcel_StdData.SATAM2 = ""
        uExcel_StdData.SATAM3 = ""
        uExcel_StdData.SATAM4 = ""
        uExcel_StdData.SATAM5 = ""
        uExcel_StdData.SATAM6 = ""
        uExcel_StdData.SATAM7 = ""
        uExcel_StdData.SATAM8 = ""
        uExcel_StdData.SATAM9 = ""
        uExcel_StdData.SATAM10 = ""
        uExcel_StdData.SATAM11 = ""
        
        For ni = 1 To 11 Step 1
            sTmp = ""
            nC = 10 + ni
            If IsNull(DBExRec.Fields(nC)) = False Then sTmp = Trim(DBExRec.Fields(nC))
            
            Select Case sTmp
                Case ""
                    'no action
                Case constSatams(0)
                    uExcel_StdData.SATAM1 = constSatamCodes(0) & "|"
                Case constSatams(1)
                    uExcel_StdData.SATAM2 = constSatamCodes(1) & "|"
                Case constSatams(2)
                    uExcel_StdData.SATAM3 = constSatamCodes(2) & "|"
                Case constSatams(3)
                    uExcel_StdData.SATAM4 = constSatamCodes(3) & "|"
                Case constSatams(4)
                    uExcel_StdData.SATAM5 = constSatamCodes(4) & "|"
                Case constSatams(5)
                    uExcel_StdData.SATAM6 = constSatamCodes(5) & "|"
                Case constSatams(6)
                    uExcel_StdData.SATAM7 = constSatamCodes(6) & "|"
                Case constSatams(7)
                    uExcel_StdData.SATAM8 = constSatamCodes(7) & "|"
                Case constSatams(8)
                    uExcel_StdData.SATAM9 = constSatamCodes(8) & "|"
                Case constSatams(9)
                    uExcel_StdData.SATAM10 = constSatamCodes(9) & "|"
'                Case "����"
'                    uExcel_StdData.SATAM11 = constSatamCodes(10) & "|"
            End Select
        Next ni
    '��2�ܱ���
        uExcel_StdData.ENG2 = ""
        
        sTmp = ""
            nC = 10 + 11 + 1
            If IsNull(DBExRec.Fields(nC)) = False Then sTmp = Trim(DBExRec.Fields(nC))
            
            Select Case sTmp
                Case ""
                    'no action
                Case "����"
                    uExcel_StdData.ENG2 = "31|"
                Case "�Ͼ�"
                    uExcel_StdData.ENG2 = "32|"
                Case "����", "�����ĳ�"
                    uExcel_StdData.ENG2 = "33|"
                Case "�Ҿ�"
                    uExcel_StdData.ENG2 = "34|"
                Case "�߱�", "�߾�"
                    uExcel_StdData.ENG2 = "35|"
                Case "�ѹ�"
                    uExcel_StdData.ENG2 = "36|"
                
                '<< ���� >> : 2008.01.09
                Case "���"
                    uExcel_StdData.ENG2 = "37|"
                Case "����"
                    uExcel_StdData.ENG2 = "38|"
                Case "����"
                    uExcel_StdData.ENG2 = "39|"
                Case "�����"
                    uExcel_StdData.ENG2 = "40|"
                Case "��������"
                    uExcel_StdData.ENG2 = "41|"
                Case "�ƶ���"
                    uExcel_StdData.ENG2 = "42|"
                Case "��Ʈ����"
                    uExcel_StdData.ENG2 = "44|"
                    
            End Select
    '��Ž
        uExcel_StdData.GWATAM1 = ""
        uExcel_StdData.GWATAM2 = ""
        uExcel_StdData.GWATAM3 = ""
        uExcel_StdData.GWATAM4 = ""
        uExcel_StdData.GWATAM5 = ""
        uExcel_StdData.GWATAM6 = ""
        uExcel_StdData.GWATAM7 = ""
        uExcel_StdData.GWATAM8 = ""
        
        For ni = 1 To 8 Step 1
            sTmp = ""
            nC = 10 + ni
            If IsNull(DBExRec.Fields(nC)) = False Then sTmp = Trim(DBExRec.Fields(nC))
            
            Select Case sTmp
                Case ""
                    'no action
                Case "��1"
                    uExcel_StdData.GWATAM1 = "51|"
                Case "ȭ1"
                    uExcel_StdData.GWATAM2 = "52|"
                Case "��1"
                    uExcel_StdData.GWATAM3 = "53|"
                Case "��1"
                    uExcel_StdData.GWATAM4 = "54|"
                Case "��2"
                    uExcel_StdData.GWATAM5 = "55|"
                Case "ȭ2"
                    uExcel_StdData.GWATAM6 = "56|"
                Case "��2"
                    uExcel_StdData.GWATAM7 = "57|"
                Case "��2"
                    uExcel_StdData.GWATAM8 = "58|"
            End Select
        Next ni
    '����
        uExcel_StdData.SURI = ""
        
        sTmp = ""
            nC = 10 + 11 + 1
            If IsNull(DBExRec.Fields(nC)) = False Then sTmp = Trim(DBExRec.Fields(nC))
            
            Select Case sTmp
                Case ""
                    'no action
                Case "����"
                    uExcel_StdData.SURI = "81|"
                Case "�̻�"
                    uExcel_StdData.SURI = "82|"
                Case "Ȯ��"
                    uExcel_StdData.SURI = "83|"
                Case "����"
                    uExcel_StdData.SURI = "84|"
            End Select
    '���
        uExcel_StdData.NONSUL1 = ""
        uExcel_StdData.NONSUL2 = ""
        uExcel_StdData.NONSUL3 = ""
        uExcel_StdData.NONSUL4 = ""
        
        For ni = 1 To 4 Step 1
            sTmp = ""
            nC = 10 + 11 + 1 + ni
            If IsNull(DBExRec.Fields(nC)) = False Then sTmp = Trim(DBExRec.Fields(nC))
            
            Select Case sTmp
                Case ""
                    'no action
                Case "���"
                    uExcel_StdData.NONSUL1 = "91|"
                Case "����"
                    uExcel_StdData.NONSUL2 = "92|"
                Case "�ܱ���"                           '< ����
                    uExcel_StdData.NONSUL3 = "93|"
                Case ""                                 '< ����
                    uExcel_StdData.NONSUL4 = "94|"
            End Select
        Next ni
        
        
    '## �������忡 ������ �ֱ� --------------------------------------------------------------------
        With sprExcel_STD_Data
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows:            .RowHeight(.Row) = 13
            
            '>> �п�
                .Col = 1
                    sTmp = uExcel_StdData.ACID
                    Call basFunction.Set_SprType_Text(sprExcel_STD_Data, "center", "left", basFunction.LenKor(sTmp), sTmp)
                    
            '>> �����ȣ
                .Col = .Col + 1
                    sTmp = uExcel_StdData.EXMID
                    Call basFunction.Set_SprType_Text(sprExcel_STD_Data, "center", "left", basFunction.LenKor(sTmp), sTmp)
            '>> �л���
                .Col = .Col + 1
                    sTmp = uExcel_StdData.STDNM
                    Call basFunction.Set_SprType_Text(sprExcel_STD_Data, "center", "left", basFunction.LenKor(sTmp), sTmp)
            '>> �������
                .Col = .Col + 1
                    sTmp = Replace(uExcel_StdData.Birth_ymd, "-", "", 1, -1, vbTextCompare)
                    Call basFunction.Set_SprType_Text(sprExcel_STD_Data, "center", "left", basFunction.LenKor(sTmp), sTmp)
            '>> ��.������
                .Col = .Col + 1
                    sTmp = uExcel_StdData.EXMTYPE
                    Call basFunction.Set_SprType_Text(sprExcel_STD_Data, "center", "left", basFunction.LenKor(sTmp), sTmp)
            '>> �迭
                .Col = .Col + 1
                    sTmp = uExcel_StdData.kaeyol
                    Call basFunction.Set_SprType_Text(sprExcel_STD_Data, "center", "left", basFunction.LenKor(sTmp), sTmp)
                    
            '>> 1 �����п�
                .Col = .Col + 1
                    sTmp = uExcel_StdData.WANT_ACID1
                    Call basFunction.Set_SprType_Text(sprExcel_STD_Data, "center", "left", basFunction.LenKor(sTmp), sTmp)
            '>> 2 �����п�
                .Col = .Col + 1
                    sTmp = uExcel_StdData.WANT_ACID2
                    Call basFunction.Set_SprType_Text(sprExcel_STD_Data, "center", "left", basFunction.LenKor(sTmp), sTmp)
                    
            '>> ����
                .Col = .Col + 1
                    nTmp = uExcel_StdData.KOR
                    Call basFunction.Set_SprType_Numeric(sprExcel_STD_Data, 0, 0, 9999, "", nTmp)
            '>> ����
                .Col = .Col + 1
                    nTmp = uExcel_StdData.ENG
                    Call basFunction.Set_SprType_Numeric(sprExcel_STD_Data, 0, 0, 9999, "", nTmp)
            '>> ����
                .Col = .Col + 1
                    nTmp = uExcel_StdData.MAT
                    Call basFunction.Set_SprType_Numeric(sprExcel_STD_Data, 0, 0, 9999, "", nTmp)
                    
            '>> ��Ž
                .Col = .Col + 1
                    sTmp = ""
                    sTmp = sTmp & Trim(uExcel_StdData.SATAM1)
                    sTmp = sTmp & Trim(uExcel_StdData.SATAM2)
                    sTmp = sTmp & Trim(uExcel_StdData.SATAM3)
                    sTmp = sTmp & Trim(uExcel_StdData.SATAM4)
                    sTmp = sTmp & Trim(uExcel_StdData.SATAM5)
                    sTmp = sTmp & Trim(uExcel_StdData.SATAM6)
                    sTmp = sTmp & Trim(uExcel_StdData.SATAM7)
                    sTmp = sTmp & Trim(uExcel_StdData.SATAM8)
                    sTmp = sTmp & Trim(uExcel_StdData.SATAM9)
                    sTmp = sTmp & Trim(uExcel_StdData.SATAM10)
                    sTmp = sTmp & Trim(uExcel_StdData.SATAM11)
                    
                    sTmp = Replace(sTmp, " ", "", 1, -1, vbTextCompare)
                    Call basFunction.Set_SprType_Text(sprExcel_STD_Data, "center", "left", basFunction.LenKor(sTmp), sTmp)
                    
            '>> ��2�ܱ���
                .Col = .Col + 1
                    sTmp = Trim(uExcel_StdData.ENG2)
                    Call basFunction.Set_SprType_Text(sprExcel_STD_Data, "center", "left", basFunction.LenKor(sTmp), sTmp)
                    
            '>> ��Ž
                .Col = .Col + 1
                    sTmp = ""
                    sTmp = sTmp & Trim(uExcel_StdData.GWATAM1)
                    sTmp = sTmp & Trim(uExcel_StdData.GWATAM2)
                    sTmp = sTmp & Trim(uExcel_StdData.GWATAM3)
                    sTmp = sTmp & Trim(uExcel_StdData.GWATAM4)
                    sTmp = sTmp & Trim(uExcel_StdData.GWATAM5)
                    sTmp = sTmp & Trim(uExcel_StdData.GWATAM6)
                    sTmp = sTmp & Trim(uExcel_StdData.GWATAM7)
                    sTmp = sTmp & Trim(uExcel_StdData.GWATAM8)
                    
                    sTmp = Replace(sTmp, " ", "", 1, -1, vbTextCompare)
                    Call basFunction.Set_SprType_Text(sprExcel_STD_Data, "center", "left", basFunction.LenKor(sTmp), sTmp)
                    
            '>> ����
                .Col = .Col + 1
                    sTmp = Trim(uExcel_StdData.SURI)
                    Call basFunction.Set_SprType_Text(sprExcel_STD_Data, "center", "left", basFunction.LenKor(sTmp), sTmp)
                    
            '>> ���
                .Col = .Col + 1
                    sTmp = ""
                    sTmp = sTmp & Trim(uExcel_StdData.NONSUL1)
                    sTmp = sTmp & Trim(uExcel_StdData.NONSUL2)
                    sTmp = sTmp & Trim(uExcel_StdData.NONSUL3)
                    sTmp = sTmp & Trim(uExcel_StdData.NONSUL4)
                    
                    sTmp = Replace(sTmp, " ", "", 1, -1, vbTextCompare)
                    Call basFunction.Set_SprType_Text(sprExcel_STD_Data, "center", "left", basFunction.LenKor(sTmp), sTmp)
                    
        End With
        
        DBExRec.MoveNext
        
    Next nRow
    
    
    
    With sprExcel_STD_Data
        If .MaxRows > 0 Then
            .Row = 1:   .Row2 = .MaxRows
            .Col = 1:   .Col2 = .MaxCols
            .BlockMode = True
                .BackColor = basModule.WhiteColor
                .BackColorStyle = BackColorStyleUnderGrid
            .BlockMode = False
            
            '.ColsFrozen = 3
            '.SetCellBorder 3, 1, 3, .MaxRows, 2, basModule.SectionColor1, CellBorderStyleSolid
            
        End If
    End With

    
    Set DBExRec = Nothing
    Set DBExCmd = Nothing
    Set xlsDBConn = Nothing
    
    MsgBox "�л� �����ڷḦ ������ �Խ��ϴ�.", vbInformation + vbOKOnly, Me.Caption
    
    On Error GoTo 0
    Exit Sub
ErrStmt1:
    MsgBox "���� ���ϼ����� �Ͻʽÿ�.", vbExclamation + vbOKOnly, Me.Caption
    Exit Sub
ErrStmt2:
    Set DBExRec = Nothing
    Set DBExCmd = Nothing
    xlsDBConn.Close
    Set xlsDBConn = Nothing
    
    MsgBox "EXCEL �ڷ� Open�� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, Me.Caption
    On Error GoTo 0
    Exit Sub
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
            
        sprExcel_STD_Data.Enabled = True
        
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

'>> �л���� ����
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
    
        '>> ��Ͽ���
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
                Set DBParam = DBCmd.CreateParameter("V_Birth_ymd", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
        
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
'            If Trim(Right(cboPass1.Text, 30)) <> "X" Then
'                sTmp = Trim(Right(cboPass1.Text, 30))
'            End If
            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
                Set DBParam = DBCmd.CreateParameter("V_PASS1", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
        '>> 2���� �հ��п�
            sTmp = ""
'            If Trim(Right(cboPass2.Text, 30)) <> "X" Then
'                sTmp = Trim(Right(cboPass2.Text, 30))
'            End If
            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
                Set DBParam = DBCmd.CreateParameter("V_PASS2", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
        '>> 3���� �հ��п�
            sTmp = ""
'            If Trim(Right(cboPass3.Text, 30)) <> "X" Then
'                sTmp = Trim(Right(cboPass3.Text, 30))
'            End If
            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
                Set DBParam = DBCmd.CreateParameter("V_PASS3", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
        '>> 4���� �հ��п�
            sTmp = ""
'            If Trim(Right(cboPass4.Text, 30)) <> "X" Then
'                sTmp = Trim(Right(cboPass4.Text, 30))
'            End If
            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
                Set DBParam = DBCmd.CreateParameter("V_PASS4", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
            
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
    sStr = sStr & "  SELECT SCHNO AS �ý����ڵ�   , "
    sStr = sStr & "         ACID  AS �п�   , "
    sStr = sStr & "         EXMID AS �����ȣ, STDNM AS �л�, SUBSTR(Birth_ymd,1,4)||'-'||SUBSTR(Birth_ymd,5,2) ||'-'||SUBSTR(Birth_ymd,7,2) AS �������,"
    sStr = sStr & "         DECODE(EXMTYPE,'0','������','1','������') AS ��������, "
    sStr = sStr & "         DECODE(KAEYOL,'01','�ι�',"
    sStr = sStr & "                       '02','�ڿ�',"
'<< �迭 >> : 2008.01.09
    If Trim(basModule.schcd) = "N" Then
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
        sStr = sStr & "                   '16','��)�ڿ�����',"
    End If
'<< �迭 >> : 2008.01.09
    If Trim(basModule.schcd) = "K" Or Trim(basModule.schcd) = "W" Or Trim(basModule.schcd) = "Q" Then
        sStr = sStr & "                   '04','�ָ�����',"
        sStr = sStr & "                   '05','�ָ��Ǵ�',"
        sStr = sStr & "                   '06','�߰�����',"
        sStr = sStr & "                   '07','�߰��Ǵ�',"
        
        sStr = sStr & "                   '11','�������ι�',"
        sStr = sStr & "                   '12','�������ڿ�',"
        
        sStr = sStr & "                   '16','�������ι�16',"
        sStr = sStr & "                   '17','�������ڿ�17',"
        
    End If
'<< �迭 >> : 2008.02.15
    If Trim(basModule.schcd) = "S" Then
        sStr = sStr & "                   '03','��ü��',"
        
        sStr = sStr & "                   '05','�����ι�',"
        sStr = sStr & "                   '06','�����ڿ�',"
        
        sStr = sStr & "                   '11','�ż��ι�',"
        sStr = sStr & "                   '12','�ż��ڿ�',"
        
        sStr = sStr & "                   '18','�ι������̾�',"
        sStr = sStr & "                   '19','�ڿ������̾�',"
        sStr = sStr & "                   '21','�����Ư���ι�',"
        sStr = sStr & "                   '22','�����Ư���ڿ�',"
    
    End If
'<< �迭 >> : 2008.02.15
    If Trim(basModule.schcd) = "P" Then
        sStr = sStr & "                   '03','Ư���ι�',"
        sStr = sStr & "                   '04','Ư���ڿ�',"
    End If
    
    If Trim(basModule.schcd) = "J" Then
        sStr = sStr & "                   '11','�ż��ι�',"
        sStr = sStr & "                   '12','�ż��ڿ�',"
        
        sStr = sStr & "                   '18','�ι������̾�',"
        sStr = sStr & "                   '19','�ڿ������̾�',"
        sStr = sStr & "                   '21','�����Ư���ι�',"
        sStr = sStr & "                   '22','�����Ư���ڿ�',"
    End If

    sStr = sStr & "                       '','��Ÿ') AS �迭,"
    
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
    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'53|') > 0 THEN     /* ��Ž-�������1 */"
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
    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'57|') > 0 THEN     /* ��Ž-�������2 */"
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
    sStr = sStr & "  "
    sStr = sStr & "      /* ��2�ܱ��� & ���� */"
    sStr = sStr & "              CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'31|') > 0 THEN '����'"
    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'32|') > 0 THEN '�Ͼ�'"
    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'33|') > 0 THEN '����'"
    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'34|') > 0 THEN '�Ҿ�'"
    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'35|') > 0 THEN '�߾�'"
    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'36|') > 0 THEN '�ѹ�'"
    
    '<< ���� >> : 2008.01.09
    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'37|') > 0 THEN '���'"
    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'38|') > 0 THEN '����'"
    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'39|') > 0 THEN '����'"
    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'40|') > 0 THEN '�����'"
    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'41|') > 0 THEN '��������'"
    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'42|') > 0 THEN '�ƶ���'"
    
    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL4,'81|') > 0 THEN '����'"
    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL4,'82|') > 0 THEN '�̻�'"
    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL4,'83|') > 0 THEN 'Ȯ��'"
    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL4,'84|') > 0 THEN '����'"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             ' '"
    sStr = sStr & "         END END END END END END END END END END END END END END END END ��2����,"
    sStr = sStr & "  "
    sStr = sStr & "      /* ��� */"
    sStr = sStr & "         CASE WHEN INSTR(SEL5,'91|') > 0 THEN         /* ��� */"
    sStr = sStr & "             '���'"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             ' '"
    sStr = sStr & "         END �����,"
    sStr = sStr & "         CASE WHEN INSTR(SEL5,'92|') > 0 THEN         /* ���� */"
    sStr = sStr & "             '����'"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             ' '"
    sStr = sStr & "         END �������,"
    sStr = sStr & "         CASE WHEN INSTR(SEL5,'93|') > 0 THEN         /* �ܱ��� */"      '< ����
    sStr = sStr & "             '�ܱ���'"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             ' '"
    sStr = sStr & "         END ��Ž���,"
    sStr = sStr & "         CASE WHEN INSTR(SEL5,'94|') > 0 THEN         /*  */"            '< ����
    sStr = sStr & "             ''"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             ' '"
    sStr = sStr & "         END ��Ž���,"
    sStr = sStr & "  "
    sStr = sStr & "         CY_ACNT AS �������, TOT_AMT AS ��ü�ݾ�    ,"
    sStr = sStr & "         NVL(BASE_AMT1    ,0) AS �⺻�ݾ�1  ,"
    sStr = sStr & "         NVL(BASE_AMT2    ,0) AS �⺻�ݾ�2  ,"
    sStr = sStr & "         NVL(BASE_AMT3    ,0) AS �⺻�ݾ�3  ,"
    sStr = sStr & "         NVL(BASE_AMT4    ,0) AS �⺻�ݾ�4  ,"
    sStr = sStr & "         NVL(BASE_AMT5    ,0) AS �⺻�ݾ�5  ,"
    sStr = sStr & "         NVL(BASE_AMT6    ,0) AS �⺻�ݾ�6  ,"
    sStr = sStr & "         NVL(BASE_AMT7    ,0) AS �⺻�ݾ�7  ,"
    sStr = sStr & "         NVL(BASE_AMT8    ,0) AS �⺻�ݾ�8  ,"
    sStr = sStr & "         NVL(TAMGU_AMT1   ,0) AS Ž�������ݾ�1 ,"
    sStr = sStr & "         NVL(TAMGU_AMT2   ,0) AS Ž�������ݾ�2 ,"
    sStr = sStr & "         NVL(TAMGU_AMT3   ,0) AS Ž�������ݾ�3 ,"
    sStr = sStr & "         NVL(TAMGU_AMT4   ,0) AS Ž�������ݾ�4 ,"
    sStr = sStr & "         NVL(TAMGU_AMT5   ,0) AS Ž�������ݾ�5 ,"
    sStr = sStr & "         NVL(TAMGU_AMT6   ,0) AS Ž�������ݾ�6 ,"
    sStr = sStr & "         NVL(TAMGU_AMT7   ,0) AS Ž�������ݾ�7 ,"
    sStr = sStr & "         NVL(TAMGU_AMT8   ,0) AS Ž�������ݾ�8 ,"
    sStr = sStr & "         NVL(TAMGU_AMT9   ,0) AS Ž�������ݾ�9 ,"
    sStr = sStr & "         NVL(TAMGU_AMT10  ,0) AS Ž�������ݾ�10,"
    sStr = sStr & "         NVL(TAMGU_AMT11  ,0) AS Ž�������ݾ�11,"
    
    sStr = sStr & "         K_NUM AS �������, M_NUM AS ��������, E_NUM AS ��������, "
    sStr = sStr & "         (NVL(K_NUM,0)+NVL(M_NUM,0)+NVL(E_NUM,0)) AS ��ü����,"
    sStr = sStr & "         N_NUM AS ���ŵ��, "
    
    
    sStr = sStr & "         DECODE(SEL1_SCH,'N','�뷮��','K','����','S','����','P','���ĸ��̸�','M','�������̸�', 'W', '�ָ����Ǵ�','Q','�߰����Ǵ�','Y','����', 'B','�λ�') AS ��1����,"
    sStr = sStr & "         DECODE(SEL2_SCH,'N','�뷮��','K','����','S','����','P','���ĸ��̸�','M','�������̸�', 'W', '�ָ����Ǵ�','Q','�߰����Ǵ�','H','����', 'B','�λ�') AS ��2����,"
    
    sStr = sStr & "         DECODE(PASS1,'N','�뷮��','K','����','S','����','P','���ĸ��̸�','M','�������̸�', 'W', '�ָ����Ǵ�','Q','�߰����Ǵ�','Y','����', 'B','�λ�') AS �հ�1   ,"
    sStr = sStr & "         DECODE(PASS2,'N','�뷮��','K','����','S','����','P','���ĸ��̸�','M','�������̸�', 'W', '�ָ����Ǵ�','Q','�߰����Ǵ�','Y','����', 'B','�λ�') AS �հ�2   ,"
    sStr = sStr & "         DECODE(PASS3,'N','�뷮��','K','����','S','����','P','���ĸ��̸�','M','�������̸�', 'W', '�ָ����Ǵ�','Q','�߰����Ǵ�','Y','����', 'B','�λ�') AS �հ�3   ,"
    sStr = sStr & "         DECODE(PASS4,'N','�뷮��','K','����','S','����','P','���ĸ��̸�','M','�������̸�', 'W', '�ָ����Ǵ�','Q','�߰����Ǵ�','Y','����', 'B','�λ�') AS �հ�4   ,"
    
    
    sStr = sStr & "         DECODE(SEX,'M','��','F','��') AS ����        , "
    sStr = sStr & "         SUBSTR(ZIP,1,3)||'-'||SUBSTR(ZIP,4,3) AS �����ȣ, ADDR1 AS �����ּ�      , ADDR2 AS ���ּ�     ,"
    sStr = sStr & "         TEL AS ��ȭ��ȣ, CEL AS �ڵ���        , EMAIL AS �̸���     ,"
    sStr = sStr & "         HIGH_SCH AS ����б� , GRADE_YEAR AS �����⵵ ,"
    sStr = sStr & "         PRNT_NM AS �кθ�� , DECODE(PRNT_RLTN,'1','��','2','��','3','��Ÿ') AS �кθ����, "
    sStr = sStr & "         SUBSTR(PRNT_ZIP,1,3)||'-'||SUBSTR(PRNT_ZIP,4,3) AS �кθ�_�����ȣ, PRNT_ADDR1 AS �кθ�_�����ּ� , PRNT_ADDR2 AS �кθ�_���ּ�,"
    sStr = sStr & "         PRNT_TEL AS �кθ�_��ȭ��ȣ  , PRNT_CEL AS �кθ�_�ڵ���   , PRNT_JOB AS �кθ�_����   , PRNT_W_TEL AS �кθ�_������ȭ ,"
    sStr = sStr & "         PHOTO_PATH AS �����������, "
    sStr = sStr & "         DECODE(R_WAY,'1','�п����','2','���ͳݵ��','3','�п����') AS ��Ϲ�ȣ, "
    sStr = sStr & "         ORD_NO AS �ֹ���ȣ, "
    sStr = sStr & "         ACID||EXMID AS �̹������ϸ�, "
    'sStr = sStr & "         DECODE(ACID,'" & Trim(basModule.SchCD) & "','',ACID) AS WANT_ACID "
    'sStr = sStr & "         DECODE(ACID,'" & Trim(basModule.SchCD) & "','" & Trim(basModule.SchCD) & "',ACID) AS WANT_ACID, "       '< TEST
    Select Case Trim(basModule.schcd)
        Case "S"
            'sStr = sStr & " DECODE(PTS_SEL,'1','����','2','6�� �򰡿�','3','9�� �򰡿�','4','6�� �򰡿�','5','9�� �򰡿�','����') AS ����, "
            sStr = sStr & " DECODE(PTS_SEL,'1','����','2','����') AS ����, "
'        Case "P"
'            sStr = sStr & " DECODE(PTS_SEL,'8','���ɵ��','9','2007 ��','6','3���','','') AS ����, "
        Case Else
            sStr = sStr & " '' AS ����, "
    End Select
    sStr = sStr & "         REGDATE AS �������, GET_PAYGUBN(ORD_NO) AS ������, "
    sStr = sStr & "         DECODE(MU_TYPE,'1','���ɵ��','2','6�� �򰡿�','3','9�� �򰡿�','4','6�� �򰡿�','5','9�� �򰡿�','9','���ŵ��','����') AS ���, "
    sStr = sStr & "         CL_CLOSE AS �Ϸ���  "
    
        sStr = sStr & " , "
        sStr = sStr & "        J01 AS ���          ,"
        sStr = sStr & "        K01 AS ���_��       ,"
        sStr = sStr & "        J02 AS ������        ,"
        sStr = sStr & "        K02 AS ��������_��   ,"
        sStr = sStr & "        J03 AS �ܱ���        ,"
        sStr = sStr & "        K03 AS �ܱ���_��     ,"
                                   
        sStr = sStr & "        J04 AS " & constSatams(0) & "_��1      ,"
        sStr = sStr & "        K04 AS " & constSatams(0) & "_��1_��   ,"
        sStr = sStr & "        J05 AS " & constSatams(1) & "_ȭ1      ,"
        sStr = sStr & "        K05 AS " & constSatams(1) & "_ȭ1_��   ,"
        sStr = sStr & "        J06 AS " & constSatams(2) & "_��1      ,"
        sStr = sStr & "        K06 AS " & constSatams(2) & "_��1_��   ,"
        sStr = sStr & "        J07 AS " & constSatams(3) & "_����1    ,"
        sStr = sStr & "        K07 AS " & constSatams(3) & "_����1_�� ,"
        sStr = sStr & "        J08 AS " & constSatams(4) & "_��2      ,"
        sStr = sStr & "        K08 AS " & constSatams(4) & "_��2_��   ,"
        sStr = sStr & "        J09 AS " & constSatams(5) & "_ȭ2      ,"
        sStr = sStr & "        K09 AS " & constSatams(5) & "_ȭ2_��   ,"
        sStr = sStr & "        J10 AS " & constSatams(6) & "_��2      ,"
        sStr = sStr & "        K10 AS " & constSatams(6) & "_��2_��   ,"
        sStr = sStr & "        J11 AS " & constSatams(7) & "_����2    ,"
        sStr = sStr & "        K11 AS " & constSatams(7) & "_����2_�� ,"
                                   
        sStr = sStr & "        J12 AS " & constSatams(8) & "          ,"
        sStr = sStr & "        K12 AS " & constSatams(8) & "_��       ,"
        sStr = sStr & "        J13 AS " & constSatams(9) & "          ,"
        sStr = sStr & "        K13 AS " & constSatams(9) & "_��       ,"
        sStr = sStr & " '' AS J14,"
        sStr = sStr & " '' AS K14,"
                                   
        sStr = sStr & "        J15 AS ����_����     ,"
        sStr = sStr & "        K15 AS ����_����_��  ,"
        sStr = sStr & "        J16 AS �Ͼ�_�̻�     ,"
        sStr = sStr & "        K16 AS �Ͼ�_�̻�_��  ,"
        sStr = sStr & "        J17 AS ����_Ȯ��     ,"
        sStr = sStr & "        K17 AS ����_Ȯ��_��  ,"
        sStr = sStr & "        J18 AS �Ҿ�_������   ,"
        sStr = sStr & "        K18 AS �Ҿ�_������_��,"
                                   
        sStr = sStr & "        J19 AS �߾�          ,"
        sStr = sStr & "        K19 AS �߾�_��       ,"
        sStr = sStr & "        J20 AS �ѹ�          ,"
        sStr = sStr & "        K20 AS �ѹ�_��       ,"
        sStr = sStr & "        J21 AS �ƶ���        ,"
        sStr = sStr & "        K21 AS �ƶ���_��     "
        
    sStr = sStr & "    FROM ( "
    
            sStr = sStr & "  SELECT A. SCHNO           ,"
            sStr = sStr & "         MAX(ACID      ) AS ACID       ,"
            sStr = sStr & "         MAX(EXMID     ) AS EXMID      ,"
            sStr = sStr & "         MAX(STDNM     ) AS STDNM      ,"
            sStr = sStr & "         MAX(Birth_ymd     ) AS Birth_ymd      ,"
            sStr = sStr & "         MAX(EXMTYPE   ) AS EXMTYPE    , MAX(KAEYOL    ) AS KAEYOL     ,"
            sStr = sStr & "         MAX(SEL1      ) AS SEL1       , MAX(SEL2      ) AS SEL2       , MAX(SEL3      ) AS SEL3      , MAX(SEL4      ) AS SEL4      , MAX(SEL5      ) AS  SEL5      , "
            sStr = sStr & "         MAX(K_NUM     ) AS K_NUM      , MAX(M_NUM     ) AS M_NUM      , MAX(E_NUM     ) AS E_NUM     , MAX(TOT_NUM   ) AS TOT_NUM   , MAX(N_NUM   ) AS  N_NUM   ,"
            sStr = sStr & "         MAX(SEL1_SCH  ) AS SEL1_SCH   , MAX(SEL2_SCH  ) AS SEL2_SCH   ,"
            sStr = sStr & "         MAX(PASS1     ) AS PASS1      , MAX(PASS2     ) AS PASS2      , MAX(PASS3     ) AS PASS3     , MAX(PASS4     ) AS PASS4     , MAX(CL_CLOSE  ) AS  CL_CLOSE  ,"
            sStr = sStr & "         MAX(CY_ACNT   ) AS CY_ACNT    , MAX(TOT_AMT   ) AS TOT_AMT    ,"
            sStr = sStr & "         MAX(BASE_AMT1 ) AS BASE_AMT1  , MAX(BASE_AMT2 ) AS BASE_AMT2  , MAX(BASE_AMT3 ) AS BASE_AMT3 , MAX(BASE_AMT4 ) AS BASE_AMT4 ,"
            sStr = sStr & "         MAX(BASE_AMT5 ) AS BASE_AMT5  , MAX(BASE_AMT6 ) AS BASE_AMT6  , MAX(BASE_AMT7 ) AS BASE_AMT7 , MAX(BASE_AMT8 ) AS BASE_AMT8 ,"
            sStr = sStr & "         MAX(TAMGU_AMT1) AS TAMGU_AMT1 , MAX(TAMGU_AMT2) AS TAMGU_AMT2 , MAX(TAMGU_AMT3) AS TAMGU_AMT3, MAX(TAMGU_AMT4) AS TAMGU_AMT4, MAX(TAMGU_AMT5) AS  TAMGU_AMT5,"
            sStr = sStr & "         MAX(TAMGU_AMT6) AS TAMGU_AMT6 , MAX(TAMGU_AMT7) AS TAMGU_AMT7 , MAX(TAMGU_AMT8) AS TAMGU_AMT8, MAX(TAMGU_AMT9) AS TAMGU_AMT9, MAX(TAMGU_AMT10) AS TAMGU_AMT10, MAX(TAMGU_AMT11) AS TAMGU_AMT11,"
            sStr = sStr & "         MAX(SEX       ) AS SEX        ,"
            sStr = sStr & "         MAX(ZIP       ) AS ZIP        , MAX(ADDR1     ) AS ADDR1      , MAX(ADDR2     ) AS ADDR2     ,"
            sStr = sStr & "         MAX(TEL       ) AS TEL        , MAX(CEL       ) AS CEL        , MAX(EMAIL     ) AS EMAIL     ,"
            sStr = sStr & "         MAX(HIGH_SCH  ) AS HIGH_SCH   , MAX(GRADE_YEAR) AS GRADE_YEAR ,"
            sStr = sStr & "         MAX(PRNT_NM   ) AS PRNT_NM    , MAX(PRNT_RLTN ) AS PRNT_RLTN  ,"
            sStr = sStr & "         MAX(PRNT_ZIP  ) AS PRNT_ZIP   , MAX(PRNT_ADDR1) AS PRNT_ADDR1 , MAX(PRNT_ADDR2) AS PRNT_ADDR2,"
            sStr = sStr & "         MAX(PRNT_TEL  ) AS PRNT_TEL   , MAX(PRNT_CEL  ) AS PRNT_CEL   , MAX(PRNT_JOB  ) AS PRNT_JOB  , MAX(PRNT_W_TEL) AS PRNT_W_TEL,"
            sStr = sStr & "         MAX(PHOTO_PATH) AS PHOTO_PATH , MAX(R_WAY     ) AS R_WAY      , MAX(ORD_NO    ) AS ORD_NO    , "
            sStr = sStr & "         MAX(TO_CHAR(REGDATE,'YYYY-MM-DD HH24:MI:SS')) AS REGDATE      , MAX(PTS_SEL   ) AS PTS_SEL   , MAX(MU_TYPE) AS MU_TYPE "
            
            sStr = sStr & "        , "
            sStr = sStr & "        SUM(J01) AS J01,"
            sStr = sStr & "        SUM(K01) AS K01,"
            sStr = sStr & "        SUM(J02) AS J02,"
            sStr = sStr & "        SUM(K02) AS K02,"
            sStr = sStr & "        SUM(J03) AS J03,"
            sStr = sStr & "        SUM(K03) AS K03,"
            sStr = sStr & " "
            sStr = sStr & "        SUM(J04) AS J04,"
            sStr = sStr & "        SUM(K04) AS K04,"
            sStr = sStr & "        SUM(J05) AS J05,"
            sStr = sStr & "        SUM(K05) AS K05,"
            sStr = sStr & "        SUM(J06) AS J06,"
            sStr = sStr & "        SUM(K06) AS K06,"
            sStr = sStr & "        SUM(J07) AS J07,"
            sStr = sStr & "        SUM(K07) AS K07,"
            sStr = sStr & "        SUM(J08) AS J08,"
            sStr = sStr & "        SUM(K08) AS K08,"
            sStr = sStr & "        SUM(J09) AS J09,"
            sStr = sStr & "        SUM(K09) AS K09,"
            sStr = sStr & "        SUM(J10) AS J10,"
            sStr = sStr & "        SUM(K10) AS K10,"
            sStr = sStr & "        SUM(J11) AS J11,"
            sStr = sStr & "        SUM(K11) AS K11,"
            sStr = sStr & " "
            sStr = sStr & "        SUM(J12) AS J12,"
            sStr = sStr & "        SUM(K12) AS K12,"
            sStr = sStr & "        SUM(J13) AS J13,"
            sStr = sStr & "        SUM(K13) AS K13,"
            sStr = sStr & "        SUM(J14) AS J14,"
            sStr = sStr & "        SUM(K14) AS K14,"
            sStr = sStr & " "
            sStr = sStr & "        SUM(J15) AS J15,"
            sStr = sStr & "        SUM(K15) AS K15,"
            sStr = sStr & "        SUM(J16) AS J16,"
            sStr = sStr & "        SUM(K16) AS K16,"
            sStr = sStr & "        SUM(J17) AS J17,"
            sStr = sStr & "        SUM(K17) AS K17,"
            sStr = sStr & "        SUM(J18) AS J18,"
            sStr = sStr & "        SUM(K18) AS K18,"
            sStr = sStr & " "
            sStr = sStr & "        SUM(J19) AS J19,"
            sStr = sStr & "        SUM(K19) AS K19,"
            sStr = sStr & "        SUM(J20) AS J20,"
            sStr = sStr & "        SUM(K20) AS K20,"
            sStr = sStr & "        SUM(J21) AS J21,"
            sStr = sStr & "        SUM(K21) AS K21"
                            
            sStr = sStr & "    FROM ("
            '---------------------------------------------------------------------------- ��ü�л� ��ȸ START
            sStr = sStr & "          SELECT *"
            sStr = sStr & "            FROM CLSTD01TB"
            sStr = sStr & "           WHERE ACID = '" & Trim(basModule.schcd) & "'"
            sStr = sStr & "             AND EXMID > ' ' "
            sStr = sStr & "             AND BIGO2 IS NULL "
' 1������ ���� ó�� : ������ �ڵ��մϴ�. 2007.12.26 ############################################################################################
' ���������� ������ �־�� ��.
            sStr = sStr & "             AND TO_CHAR(REGDATE,'YYYYMMDDHH24') > '2007120113' "
'###############################################################################################################################################

            sStr = sStr & "          UNION ALL"
            '---------------------------------------------------------------------------- ��ü�л� ��ȸ END
            '---------------------------------------------------------------------------- �հ��� ��ȸ START
            sStr = sStr & "          SELECT *"
            sStr = sStr & "            From CLSTD01TB"
            sStr = sStr & "           WHERE (PASS1 = '" & Trim(basModule.schcd) & "'" & " OR"
            sStr = sStr & "                  PASS2 = '" & Trim(basModule.schcd) & "'" & " OR"
            sStr = sStr & "                  PASS3 = '" & Trim(basModule.schcd) & "'" & " OR"
            sStr = sStr & "                  PASS4 = '" & Trim(basModule.schcd) & "'" & " )"
            sStr = sStr & "             AND EXMID > ' ' "
            sStr = sStr & "             AND BIGO2 IS NULL "
' 1������ ���� ó�� : ������ �ڵ��մϴ�. 2007.12.26 ############################################################################################
' ���������� ������ �־�� ��.
            sStr = sStr & "             AND TO_CHAR(REGDATE,'YYYYMMDDHH24') > '2007120113' "
'###############################################################################################################################################

            sStr = sStr & "          ) A,"
            
            sStr = sStr & "          ("
            
            sStr = sStr & "         SELECT SCHNO,"
            sStr = sStr & "                DECODE(TRIM(SUB_ID), '37', DECODE(SUB_NUM,'X',0, SUB_NUM), 0)    AS J01,    /* ���                  */"
            sStr = sStr & "                DECODE(TRIM(SUB_ID), '37', DECODE(SUB_BAK,'X',0, SUB_BAK), 0)    AS K01,    /* �����  ���          */"
            sStr = sStr & "                DECODE(TRIM(SUB_ID), '38', DECODE(SUB_NUM,'X',0, SUB_NUM), 0)    AS J02,    /* ��������              */"
            sStr = sStr & "                DECODE(TRIM(SUB_ID), '38', DECODE(SUB_BAK,'X',0, SUB_BAK), 0)    AS K02,    /* �����  ��������      */"
            sStr = sStr & "                DECODE(TRIM(SUB_ID), '39', DECODE(SUB_NUM,'X',0, SUB_NUM), 0)    AS J03,    /* �ܱ���                */"
            sStr = sStr & "                DECODE(TRIM(SUB_ID), '39', DECODE(SUB_BAK,'X',0, SUB_BAK), 0)    AS K03,    /* �����  �ܱ���        */"
            
            sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(0) & "', DECODE(SUB_NUM,'X',0, SUB_NUM), '51', DECODE(SUB_NUM, 'X',0, SUB_NUM), 0) AS J04,    /* ��Ž-" & constSatams(0) & "        , ��Ž-����1             */"
            sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(0) & "', DECODE(SUB_BAK,'X',0, SUB_BAK), '51', DECODE(SUB_BAK, 'X',0, SUB_BAK), 0) AS K04,    /* �����  ��Ž-" & constSatams(0) & "        , ��Ž-����1     */"
            sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(1) & "', DECODE(SUB_NUM,'X',0, SUB_NUM), '52', DECODE(SUB_NUM, 'X',0, SUB_NUM), 0) AS J05,    /* ��Ž-" & constSatams(1) & "         , ��Ž-ȭ��1             */"
            sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(1) & "', DECODE(SUB_BAK,'X',0, SUB_BAK), '52', DECODE(SUB_BAK, 'X',0, SUB_BAK), 0) AS K05,    /* �����  ��Ž-" & constSatams(1) & "         , ��Ž-ȭ��1     */"
            sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(2) & "', DECODE(SUB_NUM,'X',0, SUB_NUM), '53', DECODE(SUB_NUM, 'X',0, SUB_NUM), 0) AS J06,    /* ��Ž-" & constSatams(2) & "         , ��Ž-�������1             */"
            sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(2) & "', DECODE(SUB_BAK,'X',0, SUB_BAK), '53', DECODE(SUB_BAK, 'X',0, SUB_BAK), 0) AS K06,    /* �����  ��Ž-" & constSatams(2) & "         , ��Ž-�������1     */"
            sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(3) & "', DECODE(SUB_NUM,'X',0, SUB_NUM), '54', DECODE(SUB_NUM, 'X',0, SUB_NUM), 0) AS J07,    /* ��Ž-" & constSatams(3) & "   , ��Ž-��������1         */"
            sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(3) & "', DECODE(SUB_BAK,'X',0, SUB_BAK), '54', DECODE(SUB_BAK, 'X',0, SUB_BAK), 0) AS K07,    /* �����  ��Ž-" & constSatams(3) & "   , ��Ž-��������1 */"
            sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(4) & "', DECODE(SUB_NUM,'X',0, SUB_NUM), '55', DECODE(SUB_NUM, 'X',0, SUB_NUM), 0) AS J08,    /* ��Ž-" & constSatams(4) & "       , ��Ž-����2             */"
            sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(4) & "', DECODE(SUB_BAK,'X',0, SUB_BAK), '55', DECODE(SUB_BAK, 'X',0, SUB_BAK), 0) AS K08,    /* �����  ��Ž-" & constSatams(4) & "       , ��Ž-����2     */"
            sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(5) & "', DECODE(SUB_NUM,'X',0, SUB_NUM), '56', DECODE(SUB_NUM, 'X',0, SUB_NUM), 0) AS J09,    /* ��Ž-" & constSatams(5) & "     , ��Ž-ȭ��2             */"
            sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(5) & "', DECODE(SUB_BAK,'X',0, SUB_BAK), '56', DECODE(SUB_BAK, 'X',0, SUB_BAK), 0) AS K09,    /* �����  ��Ž-" & constSatams(5) & "     , ��Ž-ȭ��2     */"
            sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(6) & "', DECODE(SUB_NUM,'X',0, SUB_NUM), '57', DECODE(SUB_NUM, 'X',0, SUB_NUM), 0) AS J10,      /* ��Ž-" & constSatams(6) & "     , ��Ž-�������2           */"
            sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(6) & "', DECODE(SUB_BAK,'X',0, SUB_BAK), '57', DECODE(SUB_BAK, 'X',0, SUB_BAK), 0) AS K10,      /* ����� ��Ž-" & constSatams(6) & "     , ��Ž-�������2    */"
            sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(7) & "', DECODE(SUB_NUM,'X',0, SUB_NUM), '58', DECODE(SUB_NUM, 'X',0, SUB_NUM), 0) AS J11,    /* ��Ž-" & constSatams(7) & "         , ��Ž-��������2         */"
            sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(7) & "', DECODE(SUB_BAK,'X',0, SUB_BAK), '58', DECODE(SUB_BAK, 'X',0, SUB_BAK), 0) AS K11,    /* �����  ��Ž-" & constSatams(7) & "         , ��Ž-��������2 */"
            
            sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(8) & "', DECODE(SUB_NUM,'X',0, SUB_NUM), 0)    AS J12,    /* ��Ž-" & constSatams(8) & "          */"
            sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(8) & "', DECODE(SUB_BAK,'X',0, SUB_BAK), 0)    AS K12,    /* �����  ��Ž-" & constSatams(8) & "  */"
            sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(9) & "', DECODE(SUB_NUM,'X',0, SUB_NUM), 0)    AS J13,    /* ��Ž-" & constSatams(9) & "          */"
            sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(9) & "', DECODE(SUB_BAK,'X',0, SUB_BAK), 0)    AS K13,    /* �����  ��Ž-" & constSatams(9) & "  */"
            sStr = sStr & " '' AS J14,"
            sStr = sStr & " '' AS K14,"
            
            sStr = sStr & "                DECODE(TRIM(SUB_ID), '31', DECODE(SUB_NUM,'X',0, SUB_NUM), '81', DECODE(SUB_NUM, 'X',0, SUB_NUM), 0) AS J15,    /* ����             , ������                 */"
            sStr = sStr & "                DECODE(TRIM(SUB_ID), '31', DECODE(SUB_BAK,'X',0, SUB_BAK), '81', DECODE(SUB_BAK, 'X',0, SUB_BAK), 0) AS K15,    /* �����  ����             , ������         */"
            sStr = sStr & "                DECODE(TRIM(SUB_ID), '32', DECODE(SUB_NUM,'X',0, SUB_NUM), '82', DECODE(SUB_NUM, 'X',0, SUB_NUM), 0) AS J16,    /* �Ͼ�             , �̻����               */"
            sStr = sStr & "                DECODE(TRIM(SUB_ID), '32', DECODE(SUB_BAK,'X',0, SUB_BAK), '82', DECODE(SUB_BAK, 'X',0, SUB_BAK), 0) AS K16,    /* �����  �Ͼ�             , �̻����       */"
            sStr = sStr & "                DECODE(TRIM(SUB_ID), '33', DECODE(SUB_NUM,'X',0, SUB_NUM), '83', DECODE(SUB_NUM, 'X',0, SUB_NUM), 0) AS J17,    /* �����ĳ�         , Ȯ�����               */"
            sStr = sStr & "                DECODE(TRIM(SUB_ID), '33', DECODE(SUB_BAK,'X',0, SUB_BAK), '83', DECODE(SUB_BAK, 'X',0, SUB_BAK), 0) AS K17,    /* �����  �����ĳ�         , Ȯ�����       */"
            sStr = sStr & "                DECODE(TRIM(SUB_ID), '34', DECODE(SUB_NUM,'X',0, SUB_NUM), '43', DECODE(SUB_NUM, 'X',0, SUB_NUM), 0) AS J18,    /* �Ҿ�             , ��������               */"
            sStr = sStr & "                DECODE(TRIM(SUB_ID), '34', DECODE(SUB_BAK,'X',0, SUB_BAK), '43', DECODE(SUB_BAK, 'X',0, SUB_BAK), 0) AS K18,    /* �����  �Ҿ�             , ��������       */"
            
            sStr = sStr & "                DECODE(TRIM(SUB_ID), '35', DECODE(SUB_NUM,'X',0, SUB_NUM), 0)    AS J19,    /* �߱���                */"
            sStr = sStr & "                DECODE(TRIM(SUB_ID), '35', DECODE(SUB_BAK,'X',0, SUB_BAK), 0)    AS K19,    /* �����  �߱���        */"
            sStr = sStr & "                DECODE(TRIM(SUB_ID), '36', DECODE(SUB_NUM,'X',0, SUB_NUM), 0)    AS J20,    /* �ѹ�                  */"
            sStr = sStr & "                DECODE(TRIM(SUB_ID), '36', DECODE(SUB_BAK,'X',0, SUB_BAK), 0)    AS K20,    /* �����  �ѹ�          */"
            sStr = sStr & "                DECODE(TRIM(SUB_ID), '42', DECODE(SUB_NUM,'X',0, SUB_NUM), 0)    AS J21,    /* �ƶ���                */"
            sStr = sStr & "                DECODE(TRIM(SUB_ID), '42', DECODE(SUB_BAK,'X',0, SUB_BAK), 0)    AS K21     /* �����  �ƶ���        */"
            sStr = sStr & "           FROM CLSTD03TB"
            
            sStr = sStr & "        ) B"

            sStr = sStr & "   WHERE A.SCHNO = B.SCHNO(+)"
            sStr = sStr & "   GROUP BY A.SCHNO"
            '---------------------------------------------------------------------------- �հ��� ��ȸ END
    
    sStr = sStr & "    ) "
    sStr = sStr & " ORDER BY EXMID "
    
    
    
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


