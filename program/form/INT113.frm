VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form INT113 
   Caption         =   "���л��� >> ���п��� ��� >> ���� ���п��� ��� (����)"
   ClientHeight    =   11265
   ClientLeft      =   555
   ClientTop       =   2970
   ClientWidth     =   15810
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   11265
   ScaleWidth      =   15810
   Begin VB.Frame Frame2 
      BackColor       =   &H0082C8E8&
      BorderStyle     =   0  '����
      Caption         =   "Frame2"
      Height          =   495
      Left            =   30
      TabIndex        =   136
      Top             =   0
      Width           =   14445
      Begin VB.Frame Frame1 
         BackColor       =   &H00D2EAF5&
         BorderStyle     =   0  '����
         Caption         =   "Frame1"
         Height          =   435
         Left            =   30
         TabIndex        =   137
         Top             =   30
         Width           =   14385
         Begin VB.ComboBox cboSel 
            Height          =   300
            Left            =   720
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   139
            Top             =   -30
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.ComboBox cboinGbn 
            Height          =   300
            Left            =   9240
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   6
            Top             =   90
            Width           =   885
         End
         Begin VB.ComboBox cboExmType 
            Height          =   300
            Left            =   4710
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   3
            Top             =   90
            Width           =   855
         End
         Begin VB.TextBox txtPage 
            Enabled         =   0   'False
            Height          =   375
            Left            =   13410
            TabIndex        =   138
            Text            =   "txtPage"
            Top             =   30
            Width           =   615
         End
         Begin VB.CommandButton cmdShiftLeft 
            Caption         =   "��"
            Height          =   375
            Left            =   13020
            TabIndex        =   9
            Top             =   30
            Width           =   345
         End
         Begin VB.CommandButton cmdShiftRight 
            Caption         =   "��"
            Height          =   375
            Left            =   14040
            TabIndex        =   10
            Top             =   30
            Width           =   345
         End
         Begin VB.CommandButton cmdPrintAll 
            Caption         =   "��üpage���"
            Height          =   375
            Left            =   11580
            TabIndex        =   8
            Top             =   30
            Width           =   1365
         End
         Begin VB.CommandButton cmdPrint 
            Caption         =   "����page���"
            Height          =   375
            Left            =   10140
            TabIndex        =   7
            Top             =   30
            Width           =   1365
         End
         Begin VB.ComboBox cboKaeyol 
            Height          =   300
            Left            =   6000
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   4
            Top             =   90
            Width           =   915
         End
         Begin VB.CommandButton cmdFind 
            Caption         =   "�л���ȸ(&F)"
            Height          =   375
            Left            =   30
            TabIndex        =   0
            Top             =   30
            Width           =   1215
         End
         Begin VB.TextBox txtStdNM 
            Height          =   285
            Left            =   7380
            TabIndex        =   5
            Text            =   "txtStdNM"
            Top             =   98
            Width           =   855
         End
         Begin EditLib.fpMask fpExmID_S 
            Height          =   285
            Left            =   2040
            TabIndex        =   1
            Top             =   60
            Width           =   735
            _Version        =   196608
            _ExtentX        =   1296
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
            Height          =   285
            Left            =   3150
            TabIndex        =   2
            Top             =   75
            Width           =   735
            _Version        =   196608
            _ExtentX        =   1296
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
         Begin VB.Label NonPrintLbl 
            Alignment       =   1  '������ ����
            BackStyle       =   0  '����
            Caption         =   "���ͳ�/�п�"
            Height          =   210
            Index           =   5
            Left            =   8130
            TabIndex        =   145
            Top             =   135
            Width           =   1095
         End
         Begin VB.Label NonPrintLbl 
            Alignment       =   1  '������ ����
            BackStyle       =   0  '����
            Caption         =   "��/������"
            Height          =   210
            Index           =   4
            Left            =   3720
            TabIndex        =   144
            Top             =   135
            Width           =   975
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
            Left            =   5640
            TabIndex        =   143
            Top             =   150
            Width           =   945
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
            TabIndex        =   142
            Top             =   30
            Visible         =   0   'False
            Width           =   945
         End
         Begin VB.Label NonPrintLbl 
            BackStyle       =   0  '����
            Caption         =   "�����ȣ        ����"
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
            Left            =   1320
            TabIndex        =   141
            Top             =   120
            Width           =   2355
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
            Index           =   0
            Left            =   6990
            TabIndex        =   140
            Top             =   150
            Width           =   945
         End
      End
   End
   Begin VB.PictureBox pReportControl 
      Height          =   9855
      Left            =   0
      ScaleHeight     =   9795
      ScaleWidth      =   14415
      TabIndex        =   11
      Top             =   510
      Width           =   14475
      Begin VB.PictureBox pReportViewer 
         Appearance      =   0  '���
         BackColor       =   &H80000005&
         BorderStyle     =   0  '����
         ForeColor       =   &H80000008&
         Height          =   9825
         Left            =   -15
         ScaleHeight     =   9825
         ScaleWidth      =   14175
         TabIndex        =   13
         Top             =   -45
         Width           =   14175
         Begin VB.TextBox ����_�������� 
            Appearance      =   0  '���
            BorderStyle     =   0  '����
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   8220
            TabIndex        =   148
            Text            =   "������,Ȯ�����,�̻����"
            Top             =   7230
            Width           =   2745
         End
         Begin VB.TextBox ����_��ȸ���� 
            Appearance      =   0  '���
            BorderStyle     =   0  '����
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   7500
            TabIndex        =   147
            Text            =   "��ȸ����"
            Top             =   6480
            Width           =   2925
         End
         Begin VB.TextBox ����_�ڿ����� 
            Appearance      =   0  '���
            BorderStyle     =   0  '����
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   7500
            TabIndex        =   146
            Text            =   "�ڿ�����"
            Top             =   7230
            Width           =   2925
         End
         Begin VB.TextBox ������� 
            Appearance      =   0  '���
            BorderStyle     =   0  '����
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   10425
            TabIndex        =   53
            Text            =   "100"
            Top             =   9225
            Width           =   375
         End
         Begin VB.TextBox �����迭2 
            Alignment       =   2  '��� ����
            Appearance      =   0  '���
            BorderStyle     =   0  '����
            BeginProperty Font 
               Name            =   "����"
               Size            =   18
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   645
            TabIndex        =   52
            Text            =   "���ɴ��"
            Top             =   2370
            Width           =   1515
         End
         Begin VB.TextBox ��� 
            Appearance      =   0  '���
            BorderStyle     =   0  '����
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   7800
            TabIndex        =   51
            Text            =   "100"
            Top             =   9225
            Width           =   375
         End
         Begin VB.TextBox �����迭 
            Appearance      =   0  '���
            BorderStyle     =   0  '����
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   14.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   11400
            TabIndex        =   50
            Text            =   "��.ü�ɰ�"
            Top             =   540
            Width           =   1980
         End
         Begin VB.TextBox ���� 
            Appearance      =   0  '���
            BorderStyle     =   0  '����
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   9360
            TabIndex        =   49
            Text            =   "100"
            Top             =   9225
            Width           =   375
         End
         Begin VB.TextBox ���� 
            Appearance      =   0  '���
            BorderStyle     =   0  '����
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   8490
            TabIndex        =   48
            Text            =   "100"
            Top             =   9225
            Width           =   375
         End
         Begin VB.TextBox �����ȣ 
            Appearance      =   0  '���
            BorderStyle     =   0  '����
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   14.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   11730
            TabIndex        =   47
            Text            =   "N12501"
            Top             =   930
            Width           =   1035
         End
         Begin VB.TextBox ������� 
            BorderStyle     =   0  '����
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   8730
            TabIndex        =   46
            Text            =   "9999-99-99"
            Top             =   3135
            Width           =   2955
         End
         Begin VB.TextBox ���� 
            Appearance      =   0  '���
            BorderStyle     =   0  '����
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   6300
            TabIndex        =   45
            Text            =   "����"
            Top             =   3135
            Width           =   645
         End
         Begin VB.TextBox �л����� 
            Appearance      =   0  '���
            BorderStyle     =   0  '����
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2220
            TabIndex        =   44
            Text            =   "ȫ�浿"
            Top             =   3135
            Width           =   1545
         End
         Begin VB.TextBox �л��ּ�2 
            Appearance      =   0  '���
            BorderStyle     =   0  '����
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2190
            TabIndex        =   43
            Text            =   "53-21 �ֿ���� ���� 201ȣ "
            Top             =   4095
            Width           =   5055
         End
         Begin VB.TextBox �л�����ó_�޴��� 
            Appearance      =   0  '���
            BorderStyle     =   0  '����
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   8730
            TabIndex        =   42
            Text            =   "011-9490-8607"
            Top             =   4095
            Width           =   2955
         End
         Begin VB.TextBox �л�����ó_�� 
            Appearance      =   0  '���
            BorderStyle     =   0  '����
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   8730
            TabIndex        =   41
            Text            =   "02-2104-8600"
            Top             =   3615
            Width           =   2955
         End
         Begin VB.TextBox �л��̸��� 
            Appearance      =   0  '���
            BorderStyle     =   0  '����
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   8730
            TabIndex        =   40
            Text            =   "iiiboss_12345@mail.naver.com"
            Top             =   4545
            Width           =   2955
         End
         Begin VB.TextBox �����⵵ 
            Appearance      =   0  '���
            BorderStyle     =   0  '����
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   5400
            TabIndex        =   39
            Text            =   "2005"
            Top             =   4545
            Width           =   495
         End
         Begin VB.TextBox �л���Ű� 
            Appearance      =   0  '���
            BorderStyle     =   0  '����
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2190
            TabIndex        =   38
            Text            =   "�л���Ű�"
            Top             =   4545
            Width           =   1995
         End
         Begin VB.TextBox �л��ּ�1 
            Appearance      =   0  '���
            BorderStyle     =   0  '����
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2190
            TabIndex        =   37
            Text            =   "���� ���ı� ������"
            Top             =   3705
            Width           =   5055
         End
         Begin VB.TextBox ��ȣ���ּ�2 
            Appearance      =   0  '���
            BorderStyle     =   0  '����
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2190
            TabIndex        =   36
            Text            =   "���� �߱� �Ŵ絿 ��������..................."
            Top             =   6015
            Width           =   5055
         End
         Begin VB.TextBox ��ȣ�ڿ���ó_���� 
            Appearance      =   0  '���
            BorderStyle     =   0  '����
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   8730
            TabIndex        =   35
            Text            =   "02-2104-8600"
            Top             =   6000
            Width           =   1395
         End
         Begin VB.TextBox ��ȣ�ڿ���ó_�޴��� 
            Appearance      =   0  '���
            BorderStyle     =   0  '����
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   10200
            TabIndex        =   34
            Text            =   "011-9490-8607"
            Top             =   6000
            Width           =   1425
         End
         Begin VB.TextBox ��ȣ������ 
            Appearance      =   0  '���
            BorderStyle     =   0  '����
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   8730
            TabIndex        =   33
            Text            =   "��ȣ�����ֽ�ȸ��"
            Top             =   5535
            Width           =   2955
         End
         Begin VB.TextBox ��ȣ���ּ�1 
            Appearance      =   0  '���
            BorderStyle     =   0  '����
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2190
            TabIndex        =   32
            Text            =   "���� �߱� �Ŵ絿 ��������..................."
            Top             =   5625
            Width           =   5055
         End
         Begin VB.TextBox ��ȣ�ڰ��� 
            Appearance      =   0  '���
            BorderStyle     =   0  '����
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   6660
            TabIndex        =   31
            Text            =   "�θ�"
            Top             =   5040
            Width           =   555
         End
         Begin VB.TextBox ��ȣ�ڼ��� 
            Appearance      =   0  '���
            BorderStyle     =   0  '����
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2190
            TabIndex        =   30
            Text            =   "ȫ�浿"
            Top             =   5055
            Width           =   1545
         End
         Begin VB.TextBox ����_����Ž�� 
            Appearance      =   0  '���
            BorderStyle     =   0  '����
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3390
            TabIndex        =   29
            Text            =   "����II,����II,����II"
            Top             =   7230
            Width           =   4095
         End
         Begin VB.TextBox ����_�ܱ��� 
            Appearance      =   0  '���
            BorderStyle     =   0  '����
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3390
            TabIndex        =   28
            Text            =   "����,�Ҿ�,�Ͼ�"
            Top             =   6840
            Width           =   4395
         End
         Begin VB.TextBox ����_��ȸŽ�� 
            Appearance      =   0  '���
            BorderStyle     =   0  '����
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3390
            TabIndex        =   27
            Text            =   "�����,�����,����"
            Top             =   6480
            Width           =   4095
         End
         Begin VB.TextBox ��ȣ�ڿ�����ȣ 
            Appearance      =   0  '���
            BorderStyle     =   0  '����
            BeginProperty Font 
               Name            =   "����"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2190
            TabIndex        =   26
            Text            =   "(100-100)"
            Top             =   5430
            Width           =   1005
         End
         Begin VB.TextBox �л�������ȣ 
            Appearance      =   0  '���
            BorderStyle     =   0  '����
            BeginProperty Font 
               Name            =   "����"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2190
            TabIndex        =   25
            Text            =   "(100-100)"
            Top             =   3510
            Width           =   1005
         End
         Begin VB.TextBox ������_���� 
            Appearance      =   0  '���
            BorderStyle     =   0  '����
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   12810
            TabIndex        =   24
            Text            =   "100"
            Top             =   8910
            Width           =   375
         End
         Begin VB.TextBox ������_���� 
            Appearance      =   0  '���
            BorderStyle     =   0  '����
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   12810
            TabIndex        =   23
            Text            =   "100"
            Top             =   9210
            Width           =   375
         End
         Begin VB.TextBox ������_���� 
            Appearance      =   0  '���
            BorderStyle     =   0  '����
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   12810
            TabIndex        =   22
            Text            =   "100"
            Top             =   9480
            Width           =   375
         End
         Begin VB.TextBox �п����� 
            Appearance      =   0  '���
            BorderStyle     =   0  '����
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   14.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   12780
            TabIndex        =   21
            Text            =   "-int"
            Top             =   930
            Width           =   675
         End
         Begin VB.TextBox �����п� 
            Appearance      =   0  '���
            BorderStyle     =   0  '����
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   14.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   11430
            TabIndex        =   20
            Text            =   "K"
            Top             =   930
            Width           =   315
         End
         Begin VB.TextBox ��ȣ�ڿ���ó 
            Appearance      =   0  '���
            BorderStyle     =   0  '����
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   8730
            TabIndex        =   19
            Text            =   "011-9490-8607"
            Top             =   5040
            Width           =   1485
         End
         Begin VB.TextBox ��� 
            Appearance      =   0  '���
            BorderStyle     =   0  '����
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   11310
            TabIndex        =   18
            Text            =   "���"
            Top             =   240
            Width           =   2175
         End
         Begin VB.TextBox �������2 
            Appearance      =   0  '���
            BorderStyle     =   0  '����
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   10515
            TabIndex        =   17
            Text            =   "100"
            Top             =   9450
            Width           =   375
         End
         Begin VB.TextBox ���2 
            Appearance      =   0  '���
            BorderStyle     =   0  '����
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   7860
            TabIndex        =   16
            Text            =   "100"
            Top             =   9450
            Width           =   375
         End
         Begin VB.TextBox ����2 
            Appearance      =   0  '���
            BorderStyle     =   0  '����
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   9420
            TabIndex        =   15
            Text            =   "100"
            Top             =   9450
            Width           =   375
         End
         Begin VB.TextBox ����2 
            Appearance      =   0  '���
            BorderStyle     =   0  '����
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   8550
            TabIndex        =   14
            Text            =   "100"
            Top             =   9450
            Width           =   375
         End
         Begin VB.Label Label1 
            BackStyle       =   0  '����
            Caption         =   "��.��.������"
            BeginProperty Font 
               Name            =   "����"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   10080
            TabIndex        =   135
            Top             =   8940
            Width           =   945
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   24
            X1              =   8250
            X2              =   8250
            Y1              =   8880
            Y2              =   9690
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            BorderWidth     =   2
            Index           =   1
            X1              =   11730
            X2              =   11730
            Y1              =   2970
            Y2              =   7920
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   53
            X1              =   6585
            X2              =   6585
            Y1              =   8880
            Y2              =   9690
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '����
            Caption         =   "11"
            BeginProperty Font 
               Name            =   "����"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   41
            Left            =   6240
            TabIndex        =   134
            Top             =   9015
            Width           =   195
         End
         Begin VB.Label OPTIONS 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  '����
            Caption         =   "���ڿ��� �л� �� ����(��)���� �����ϴ� �л��� ���� ǥ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   2
            Left            =   510
            TabIndex        =   133
            Top             =   8400
            Width           =   7395
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '����
            Caption         =   "���"
            BeginProperty Font 
               Name            =   "����"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   42
            Left            =   7740
            TabIndex        =   132
            Top             =   8940
            Width           =   375
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '����
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   86
            Left            =   11670
            TabIndex        =   131
            Top             =   8985
            Width           =   225
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '����
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   83
            Left            =   11670
            TabIndex        =   130
            Top             =   9195
            Width           =   225
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '����
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   82
            Left            =   11670
            TabIndex        =   129
            Top             =   9405
            Width           =   225
         End
         Begin VB.Shape Boxs 
            BorderColor     =   &H00FF0000&
            Height          =   795
            Index           =   1
            Left            =   9960
            Top             =   480
            Width           =   3555
         End
         Begin VB.Shape Boxs 
            BorderColor     =   &H00FF0000&
            BorderWidth     =   2
            Height          =   4965
            Index           =   0
            Left            =   525
            Top             =   2970
            Width           =   13005
         End
         Begin VB.Shape Boxs 
            BorderColor     =   &H00FF0000&
            Height          =   585
            Index           =   2
            Left            =   510
            Top             =   2250
            Width           =   1755
         End
         Begin VB.Shape Boxs 
            BorderColor     =   &H00FF0000&
            Height          =   825
            Index           =   5
            Left            =   510
            Top             =   8865
            Width           =   13005
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            BorderWidth     =   2
            Index           =   28
            X1              =   510
            X2              =   11730
            Y1              =   4890
            Y2              =   4890
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            BorderWidth     =   2
            Index           =   0
            X1              =   510
            X2              =   11700
            Y1              =   6330
            Y2              =   6330
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            BorderStyle     =   3  '��
            Index           =   2
            X1              =   510
            X2              =   14130
            Y1              =   1380
            Y2              =   1380
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            BorderStyle     =   3  '��
            Index           =   3
            X1              =   525
            X2              =   5700
            Y1              =   8685
            Y2              =   8685
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   4
            X1              =   1080
            X2              =   11730
            Y1              =   3450
            Y2              =   3450
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   5
            X1              =   2070
            X2              =   11730
            Y1              =   3930
            Y2              =   3930
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   6
            X1              =   1080
            X2              =   11730
            Y1              =   4410
            Y2              =   4410
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   7
            X1              =   1080
            X2              =   11730
            Y1              =   5370
            Y2              =   5370
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   8
            X1              =   2070
            X2              =   11730
            Y1              =   5850
            Y2              =   5850
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   9
            X1              =   1560
            X2              =   11730
            Y1              =   6750
            Y2              =   6750
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   10
            X1              =   1080
            X2              =   11730
            Y1              =   7140
            Y2              =   7140
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   11
            X1              =   3330
            X2              =   11730
            Y1              =   7500
            Y2              =   7500
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   12
            X1              =   11280
            X2              =   13500
            Y1              =   870
            Y2              =   870
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   13
            X1              =   11280
            X2              =   11280
            Y1              =   480
            Y2              =   1260
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   14
            X1              =   1080
            X2              =   1080
            Y1              =   2970
            Y2              =   7920
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   15
            X1              =   7650
            X2              =   7650
            Y1              =   2970
            Y2              =   6330
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   16
            X1              =   8640
            X2              =   8640
            Y1              =   2970
            Y2              =   6330
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   17
            X1              =   2070
            X2              =   2070
            Y1              =   2970
            Y2              =   6330
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   18
            X1              =   5730
            X2              =   5730
            Y1              =   2970
            Y2              =   3450
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   19
            X1              =   4740
            X2              =   4740
            Y1              =   2970
            Y2              =   3450
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   20
            X1              =   5700
            X2              =   5700
            Y1              =   4890
            Y2              =   5370
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   21
            X1              =   4710
            X2              =   4710
            Y1              =   4890
            Y2              =   5370
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   22
            X1              =   1560
            X2              =   1560
            Y1              =   6330
            Y2              =   7920
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   23
            X1              =   3330
            X2              =   3330
            Y1              =   6330
            Y2              =   7920
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            BorderWidth     =   2
            Index           =   26
            X1              =   11730
            X2              =   13500
            Y1              =   5190
            Y2              =   5190
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            BorderWidth     =   2
            Index           =   27
            X1              =   11730
            X2              =   13500
            Y1              =   5550
            Y2              =   5550
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   29
            X1              =   510
            X2              =   6600
            Y1              =   9270
            Y2              =   9270
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   30
            X1              =   1515
            X2              =   1515
            Y1              =   8865
            Y2              =   9675
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   31
            X1              =   1050
            X2              =   1050
            Y1              =   8865
            Y2              =   9675
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   32
            X1              =   2535
            X2              =   2535
            Y1              =   8865
            Y2              =   9675
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   33
            X1              =   2025
            X2              =   2025
            Y1              =   8865
            Y2              =   9675
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   34
            X1              =   3555
            X2              =   3555
            Y1              =   8865
            Y2              =   9675
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   35
            X1              =   3045
            X2              =   3045
            Y1              =   8865
            Y2              =   9675
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   36
            X1              =   4575
            X2              =   4575
            Y1              =   8865
            Y2              =   9675
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   37
            X1              =   4065
            X2              =   4065
            Y1              =   8865
            Y2              =   9675
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   38
            X1              =   5595
            X2              =   5595
            Y1              =   8865
            Y2              =   9675
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   39
            X1              =   5085
            X2              =   5085
            Y1              =   8865
            Y2              =   9675
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   41
            X1              =   6105
            X2              =   6105
            Y1              =   8865
            Y2              =   9675
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   42
            X1              =   10020
            X2              =   10020
            Y1              =   8865
            Y2              =   9675
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   43
            X1              =   8970
            X2              =   8970
            Y1              =   8865
            Y2              =   9675
         End
         Begin VB.Line Lines_opt 
            BorderColor     =   &H00FF0000&
            Index           =   2
            X1              =   11055
            X2              =   11055
            Y1              =   8865
            Y2              =   9675
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   46
            X1              =   11970
            X2              =   11970
            Y1              =   8865
            Y2              =   9675
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   47
            X1              =   11550
            X2              =   11550
            Y1              =   8865
            Y2              =   9675
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   49
            X1              =   12570
            X2              =   12570
            Y1              =   8865
            Y2              =   9675
         End
         Begin VB.Line Lines_opt 
            BorderColor     =   &H00FF0000&
            Index           =   1
            X1              =   6600
            X2              =   11550
            Y1              =   9150
            Y2              =   9150
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   50
            X1              =   11970
            X2              =   13500
            Y1              =   9135
            Y2              =   9135
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   51
            X1              =   11970
            X2              =   13500
            Y1              =   9405
            Y2              =   9405
         End
         Begin VB.Label Labels 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  '����
            Caption         =   "�й� :"
            BeginProperty Font 
               Name            =   "����"
               Size            =   14.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   7
            Left            =   2430
            TabIndex        =   128
            Top             =   2460
            Width           =   750
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   52
            X1              =   2370
            X2              =   5040
            Y1              =   2820
            Y2              =   2820
         End
         Begin VB.Label Labels 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  '����
            Caption         =   "�������� ������ �� �����п� �ܿ� �ٸ� �п����� ������ ���� ��� 2������"
            BeginProperty Font 
               Name            =   "����"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   2
            Left            =   9930
            TabIndex        =   127
            Top             =   1560
            Width           =   3735
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '����
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   13
            Left            =   750
            TabIndex        =   126
            Top             =   3330
            Width           =   195
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '����
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   14
            Left            =   750
            TabIndex        =   125
            Top             =   4290
            Width           =   195
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '����
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   3
            Left            =   750
            TabIndex        =   124
            Top             =   5040
            Width           =   195
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '����
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   4
            Left            =   750
            TabIndex        =   123
            Top             =   5910
            Width           =   195
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '����
            Caption         =   "ȣ"
            BeginProperty Font 
               Name            =   "����"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   5
            Left            =   750
            TabIndex        =   122
            Top             =   5475
            Width           =   195
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '����
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   6
            Left            =   750
            TabIndex        =   121
            Top             =   6510
            Width           =   195
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '����
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   8
            Left            =   750
            TabIndex        =   120
            Top             =   7185
            Width           =   195
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '����
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   9
            Left            =   750
            TabIndex        =   119
            Top             =   6855
            Width           =   195
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '����
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   10
            Left            =   750
            TabIndex        =   118
            Top             =   7530
            Width           =   195
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '����
            Caption         =   "��     ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   11
            Left            =   1260
            TabIndex        =   117
            Top             =   3150
            Width           =   675
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '����
            Caption         =   "��     ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   12
            Left            =   1260
            TabIndex        =   116
            Top             =   3840
            Width           =   675
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '����
            Caption         =   "��     ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   15
            Left            =   1260
            TabIndex        =   115
            Top             =   5070
            Width           =   675
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '����
            Caption         =   "��     ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   16
            Left            =   1260
            TabIndex        =   114
            Top             =   5760
            Width           =   675
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '����
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   17
            Left            =   1230
            TabIndex        =   113
            Top             =   6450
            Width           =   195
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '����
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   18
            Left            =   1230
            TabIndex        =   112
            Top             =   6690
            Width           =   195
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '����
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   19
            Left            =   1230
            TabIndex        =   111
            Top             =   6930
            Width           =   195
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '����
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   20
            Left            =   1230
            TabIndex        =   110
            Top             =   7200
            Width           =   195
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '����
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   21
            Left            =   1230
            TabIndex        =   109
            Top             =   7440
            Width           =   195
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '����
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   22
            Left            =   1230
            TabIndex        =   108
            Top             =   7680
            Width           =   195
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '����
            Caption         =   "��ȸŽ��[��2 �Ǵ� 3]"
            BeginProperty Font 
               Name            =   "����"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   23
            Left            =   1650
            TabIndex        =   107
            Top             =   6480
            Width           =   1635
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '����
            Caption         =   "��2�ܱ���[��1]"
            BeginProperty Font 
               Name            =   "����"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   24
            Left            =   1650
            TabIndex        =   106
            Top             =   6870
            Width           =   1155
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '����
            Caption         =   "����Ž��[��3]"
            BeginProperty Font 
               Name            =   "����"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   26
            Left            =   1650
            TabIndex        =   105
            Top             =   7470
            Width           =   1125
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '����
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   27
            Left            =   720
            TabIndex        =   104
            Top             =   9015
            Width           =   315
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '����
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   28
            Left            =   720
            TabIndex        =   103
            Top             =   9405
            Width           =   315
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '����
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "����"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   30
            Left            =   1215
            TabIndex        =   102
            Top             =   9015
            Width           =   195
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '����
            Caption         =   "2"
            BeginProperty Font 
               Name            =   "����"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   31
            Left            =   1710
            TabIndex        =   101
            Top             =   9015
            Width           =   195
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '����
            Caption         =   "3"
            BeginProperty Font 
               Name            =   "����"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   32
            Left            =   2220
            TabIndex        =   100
            Top             =   9015
            Width           =   195
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '����
            Caption         =   "4"
            BeginProperty Font 
               Name            =   "����"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   33
            Left            =   2715
            TabIndex        =   99
            Top             =   9015
            Width           =   195
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '����
            Caption         =   "5"
            BeginProperty Font 
               Name            =   "����"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   34
            Left            =   3225
            TabIndex        =   98
            Top             =   9015
            Width           =   195
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '����
            Caption         =   "6"
            BeginProperty Font 
               Name            =   "����"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   35
            Left            =   3750
            TabIndex        =   97
            Top             =   9015
            Width           =   195
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '����
            Caption         =   "7"
            BeginProperty Font 
               Name            =   "����"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   36
            Left            =   4260
            TabIndex        =   96
            Top             =   9015
            Width           =   195
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '����
            Caption         =   "8"
            BeginProperty Font 
               Name            =   "����"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   37
            Left            =   4755
            TabIndex        =   95
            Top             =   9015
            Width           =   195
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '����
            Caption         =   "9"
            BeginProperty Font 
               Name            =   "����"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   38
            Left            =   5235
            TabIndex        =   94
            Top             =   9015
            Width           =   195
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '����
            Caption         =   "10"
            BeginProperty Font 
               Name            =   "����"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   39
            Left            =   5730
            TabIndex        =   93
            Top             =   9015
            Width           =   195
         End
         Begin VB.Label �������� 
            Alignment       =   2  '��� ����
            BackStyle       =   0  '����
            Caption         =   "����"
            BeginProperty Font 
               Name            =   "����"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   8310
            TabIndex        =   92
            Top             =   8940
            Width           =   675
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '����
            Caption         =   "�ܱ���(����)"
            BeginProperty Font 
               Name            =   "����"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   44
            Left            =   9045
            TabIndex        =   91
            Top             =   8940
            Width           =   975
         End
         Begin VB.Label OPTIONS 
            BackStyle       =   0  '����
            Caption         =   "Ȯ ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   13
            Left            =   11115
            TabIndex        =   90
            Top             =   8955
            Width           =   465
         End
         Begin VB.Label OPTIONS 
            BackStyle       =   0  '����
            Caption         =   "(��)"
            BeginProperty Font 
               Name            =   "����"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   14
            Left            =   11160
            TabIndex        =   89
            Top             =   9330
            Width           =   285
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '����
            Caption         =   "�� ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   47
            Left            =   12090
            TabIndex        =   88
            Top             =   9225
            Width           =   405
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '����
            Caption         =   "�� ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   48
            Left            =   12090
            TabIndex        =   87
            Top             =   8925
            Width           =   405
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '����
            Caption         =   "�� ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   49
            Left            =   12090
            TabIndex        =   86
            Top             =   9495
            Width           =   405
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '����
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   50
            Left            =   13230
            TabIndex        =   85
            Top             =   8925
            Width           =   165
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '����
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   51
            Left            =   13230
            TabIndex        =   84
            Top             =   9225
            Width           =   165
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '����
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   52
            Left            =   13230
            TabIndex        =   83
            Top             =   9495
            Width           =   165
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '����
            Caption         =   "�������"
            BeginProperty Font 
               Name            =   "����"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   53
            Left            =   7770
            TabIndex        =   82
            Top             =   3150
            Width           =   675
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '����
            Caption         =   "��     ȭ"
            BeginProperty Font 
               Name            =   "����"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   54
            Left            =   7770
            TabIndex        =   81
            Top             =   3630
            Width           =   675
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '����
            Caption         =   "�� �� ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   55
            Left            =   7770
            TabIndex        =   80
            Top             =   4110
            Width           =   675
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '����
            Caption         =   "�� �� ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   56
            Left            =   7770
            TabIndex        =   79
            Top             =   4560
            Width           =   675
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '����
            Caption         =   "��ȭ(�޴���)"
            BeginProperty Font 
               Name            =   "����"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   57
            Left            =   7680
            TabIndex        =   78
            Top             =   5070
            Width           =   975
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '����
            Caption         =   "����(�ٹ�ó)"
            BeginProperty Font 
               Name            =   "����"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   58
            Left            =   7680
            TabIndex        =   77
            Top             =   5550
            Width           =   975
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '����
            Caption         =   "�� �� �� ȭ"
            BeginProperty Font 
               Name            =   "����"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   59
            Left            =   7710
            TabIndex        =   76
            Top             =   6030
            Width           =   855
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '����
            Caption         =   "�뼺�п� ���п���"
            BeginProperty Font 
               Name            =   "����"
               Size            =   20.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   60
            Left            =   1590
            TabIndex        =   75
            Top             =   750
            Width           =   3585
         End
         Begin VB.Label Labels 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  '����
            Caption         =   "2013��"
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   14.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   62
            Left            =   480
            TabIndex        =   74
            Top             =   840
            Width           =   825
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '����
            Caption         =   "�ؼ����ȣ"
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   11.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   61
            Left            =   9990
            TabIndex        =   73
            Top             =   780
            Width           =   1275
         End
         Begin VB.Image Photo 
            Height          =   2145
            Left            =   11730
            Stretch         =   -1  'True
            Top             =   3000
            Width           =   1785
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '����
            Caption         =   "�� �� ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   63
            Left            =   1260
            TabIndex        =   72
            Top             =   4560
            Width           =   675
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '����
            Caption         =   "2013�� ����ī��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   24
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Index           =   64
            Left            =   510
            TabIndex        =   71
            Top             =   1620
            Width           =   4065
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '����
            Caption         =   "��     ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   65
            Left            =   4890
            TabIndex        =   70
            Top             =   3150
            Width           =   675
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '����
            Caption         =   "��     ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   66
            Left            =   4890
            TabIndex        =   69
            Top             =   5070
            Width           =   675
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '����
            Caption         =   "�� ��  ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   72
            Left            =   12270
            TabIndex        =   68
            Top             =   5310
            Width           =   675
         End
         Begin VB.Label OPTIONS 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  '����
            Caption         =   "���ι��� �л����� ��ȭŽ�� 11���� �� 4������� ������ �� ������, ��2�ܱ���� 6���� �� 1������ ������ �� �ֽ��ϴ�."
            BeginProperty Font 
               Name            =   "����"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   0
            Left            =   510
            TabIndex        =   67
            Top             =   8010
            Width           =   15000
         End
         Begin VB.Label OPTIONS 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  '����
            Caption         =   "���ڿ��� �л����� ������������ 1����, ����Ž�������� 3������� ������ �� �ֽ��ϴ�."
            BeginProperty Font 
               Name            =   "����"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   1
            Left            =   510
            TabIndex        =   66
            Top             =   8190
            Width           =   7395
         End
         Begin VB.Label Labels 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  '����
            Caption         =   "ǥ���Ͻÿ�."
            BeginProperty Font 
               Name            =   "����"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   78
            Left            =   10080
            TabIndex        =   65
            Top             =   1800
            Width           =   2535
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '����
            Caption         =   "�����б�"
            BeginProperty Font 
               Name            =   "����"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   79
            Left            =   4230
            TabIndex        =   64
            Top             =   4560
            Width           =   675
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '����
            Caption         =   "�� 2�� ����(����)"
            BeginProperty Font 
               Name            =   "����"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   80
            Left            =   6000
            TabIndex        =   63
            Top             =   4560
            Width           =   1365
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '����
            Caption         =   "�л���"
            BeginProperty Font 
               Name            =   "����"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   81
            Left            =   5850
            TabIndex        =   62
            Top             =   5070
            Width           =   585
         End
         Begin VB.Label OPTIONS 
            BackStyle       =   0  '����
            Caption         =   "�� ��2�ܱ�� �������� ���� �л��� ����, ����,"
            BeginProperty Font 
               Name            =   "����"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   3
            Left            =   7980
            TabIndex        =   61
            Top             =   6780
            Width           =   3765
         End
         Begin VB.Label OPTIONS 
            BackStyle       =   0  '����
            Caption         =   "��    ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   4
            Left            =   6810
            TabIndex        =   60
            Top             =   9495
            Width           =   735
         End
         Begin VB.Label OPTIONS 
            BackStyle       =   0  '����
            Caption         =   "ǥ������"
            BeginProperty Font 
               Name            =   "����"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   11
            Left            =   6780
            TabIndex        =   59
            Top             =   9225
            Width           =   705
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '����
            Caption         =   "����.�򰡿�"
            BeginProperty Font 
               Name            =   "����"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   46
            Left            =   6660
            TabIndex        =   58
            Top             =   8940
            Width           =   915
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   44
            X1              =   7590
            X2              =   7590
            Y1              =   8880
            Y2              =   9690
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '����
            Caption         =   "���� �Ʒ��κ��� �������� ���ÿ�"
            BeginProperty Font 
               Name            =   "����"
               Size            =   6.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   67
            Left            =   5730
            TabIndex        =   57
            Top             =   8625
            Width           =   2055
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            BorderStyle     =   3  '��
            Index           =   45
            X1              =   7800
            X2              =   14160
            Y1              =   8685
            Y2              =   8685
         End
         Begin VB.Line Lines_opt 
            BorderColor     =   &H00FF0000&
            Index           =   4
            X1              =   6600
            X2              =   11070
            Y1              =   9420
            Y2              =   9420
         End
         Begin VB.Label OPTIONS 
            BackStyle       =   0  '����
            Caption         =   "�� �ڿ��� ���м����� 2�б���� ���м��� 3����, ��.��.�� �� 1�������� ����� �� �ֽ��ϴ�."
            BeginProperty Font 
               Name            =   "����"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   21
            Left            =   3540
            TabIndex        =   56
            Top             =   7620
            Width           =   7365
         End
         Begin VB.Label OPTIONS 
            BackStyle       =   0  '����
            Caption         =   "����(6�� ���� ����)�� ������ �� �ֽ��ϴ�."
            BeginProperty Font 
               Name            =   "����"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   5
            Left            =   8190
            TabIndex        =   55
            Top             =   6960
            Width           =   3315
         End
         Begin VB.Label Labels 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  '����
            Caption         =   "ǥ���Ͻÿ�."
            BeginProperty Font 
               Name            =   "����"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   1
            Left            =   10140
            TabIndex        =   54
            Top             =   2010
            Width           =   2535
         End
         Begin VB.Shape FillBOXs 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  '�������� ����
            BorderStyle     =   0  '����
            Height          =   3375
            Index           =   0
            Left            =   1080
            Top             =   2970
            Width           =   990
         End
         Begin VB.Shape FillBOXs 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  '�������� ����
            BorderStyle     =   0  '����
            Height          =   1605
            Index           =   8
            Left            =   1080
            Top             =   6330
            Width           =   2250
         End
         Begin VB.Shape FillBOXs 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  '�������� ����
            BorderStyle     =   0  '����
            Height          =   480
            Index           =   6
            Left            =   4710
            Top             =   4890
            Width           =   990
         End
         Begin VB.Shape FillBOXs 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  '�������� ����
            BorderStyle     =   0  '����
            Height          =   480
            Index           =   5
            Left            =   4740
            Top             =   2970
            Width           =   990
         End
         Begin VB.Shape FillBOXs 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  '�������� ����
            BorderStyle     =   0  '����
            Height          =   3365
            Index           =   7
            Left            =   7650
            Top             =   2970
            Width           =   990
         End
         Begin VB.Shape FillBOXs 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  '�������� ����
            BorderStyle     =   0  '����
            Height          =   825
            Index           =   10
            Left            =   6600
            Top             =   8865
            Width           =   1020
         End
         Begin VB.Shape FillBOXs 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  '�������� ����
            BorderStyle     =   0  '����
            Height          =   825
            Index           =   9
            Left            =   510
            Top             =   8880
            Width           =   540
         End
         Begin VB.Shape FillBOXs 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  '�������� ����
            BorderStyle     =   0  '����
            Height          =   825
            Index           =   11
            Left            =   11550
            Top             =   8865
            Width           =   420
         End
         Begin VB.Shape FillBOXs 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  '�������� ����
            BorderStyle     =   0  '����
            Height          =   360
            Index           =   4
            Left            =   11730
            Top             =   5190
            Width           =   1785
         End
         Begin VB.Shape FillBOXs 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  '�������� ����
            BorderStyle     =   0  '����
            Height          =   795
            Index           =   1
            Left            =   9960
            Top             =   480
            Width           =   1320
         End
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   9765
         Left            =   14190
         TabIndex        =   12
         Top             =   0
         Width           =   225
      End
   End
   Begin MSComDlg.CommonDialog dlgPrint 
      Left            =   3420
      Top             =   10410
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   1860
      Top             =   10350
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   2490
      Top             =   10350
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
            Picture         =   "INT113.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "INT113"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################################################
'   �� �� ��  �� : �뼺�п� ���л���, �ݹ��� & �ð�ǥ ���α׷�
'   ����ý��۸� :
'   ��   ��   �� : INT110
'   �� ��  �� �� : ���п��� ���
'
'   ��   ��   �� : 2007/08/31
'   ��   ��   �� : ���ϱ�
' --------------------------------------------------------------------------------------------------------------
' --------------------------------------------------------------------------------------------------------------
'                 ��     ��     ��     ��
' --------------------------------------------------------------------------------------------------------------
'   1. ������ : 2007.12.11
'   2. ��  �� : ��¹� ����
'   1. ������ : 2009.11.05
'   2. ��  �� : ���� ���п��� 1���������� ����
'################################################################################################################

Option Explicit

Private Type tSTD
    SCHNO       As String
    ACID        As String
    EXMID       As String
    STDNM       As String
    Birth_ymd       As String
    
    EXMTYPE     As String
    KAEYOL      As String
    
    SEL1        As String
    SEL2        As String
    SEL3        As String
    SEL4        As String
    SEL5        As String
    
    K_NUM       As Long
    M_NUM       As Long
    E_NUM       As Long
    TOT_NUM     As Long
    
    JK_NUM      As Long
    JM_NUM      As Long
    JE_NUM      As Long
    JTOT_NUM    As Long
    
    KK_NUM      As Long
    KM_NUM      As Long
    KE_NUM      As Long
    KTOT_NUM    As Long
    
    
    SEL1_SCH    As String
    SEL2_SCH    As String
    
    PASS1       As String
    PASS2       As String
    PASS3       As String
    PASS4       As String
    CL_CLOSE    As String
    CY_ACNT     As String
    TOT_AMT     As Long
    
    BASE_AMT1   As Long
    BASE_AMT2   As Long
    BASE_AMT3   As Long
    BASE_AMT4   As Long
    
    BASE_AMT5   As Long
    BASE_AMT6   As Long
    BASE_AMT7   As Long
    BASE_AMT8   As Long
    BASE_AMT9   As Long
    BASE_AMT10  As Long
    
    TAMGU_AMT1  As Long
    TAMGU_AMT2  As Long
    TAMGU_AMT3  As Long
    TAMGU_AMT4  As Long
    TAMGU_AMT5  As Long
    TAMGU_AMT6  As Long
    TAMGU_AMT7  As Long
    TAMGU_AMT8  As Long
    TAMGU_AMT9  As Long
    TAMGU_AMT10 As Long
    TAMGU_AMT11 As Long
    TAMGU_AMT12 As Long
    
    SEX         As String
    
    ZIP         As String
    ADDR1       As String
    ADDR2       As String
    TEL         As String
    CEL         As String
    EMAIL       As String
    
    HIGH_SCH    As String
    GRADE_YEAR  As String
    
    PRNT_NM     As String
    PRNT_RLTN   As String
    PRNT_ZIP    As String
    PRNT_ADDR1  As String
    PRNT_ADDR2  As String
    PRNT_TEL    As String
    PRNT_CEL    As String
    PRNT_JOB    As String
    PRNT_W_TEL  As String
    
    PHOTO_PATH  As String
    PTS_SEL     As String
    R_WAY       As String
    
    ORD_NO      As String
    IMAGE_FILE  As String
    WANT_ACID   As String
    IMAGE_DIR   As String
    GR          As String
End Type
Private uSTD() As tSTD

Private sSavePath   As String       '<< image ���
Private nTotRec     As Long         '<< ��ü �л���

Private Const Kangnam = "/NDOC/dshw/kangnam/register/"
Private Const MKangnam = "/NDOC/dshw/mkangnam/register/"
Private Const MSongpa = "/NDOC/dshw/msongpa/register/"
Private Const Noryangjin = "/NDOC/dshw/noryangjin/register/"
Private Const Songpa = "/NDOC/dshw/songpa/register/"
Private Const MGwanghwa = "/NDOC/dshw/kwanghwamun/register/"
Private Const Busan = "/NDOC/dshw/busan/register/"

Private Sub Form_Load()
    Me.Top = 0
    Me.Left = 0
    Me.Width = 14550
    Me.Height = 10900
    
    Me.Tag = "LOAD"
        nTotRec = 0
        Call Clear_Form_Control
        
        sSavePath = App.Path & "\PHOTO"
        If Dir(sSavePath, vbDirectory) = "" Then
            Call MkDir(sSavePath)
        End If
        
        VScroll1.Min = 1
        VScroll1.Max = 100
        VScroll1.SmallChange = 1
        VScroll1.LargeChange = 1
        VScroll1.Enabled = False
        
        Me.Width = 14550
        Me.Height = 10755
        
        fpExmID_S.Text = ""
        fpExmID_E.Text = ""
        
        '>> ��/������
        With cboExmType
            .Clear
            .AddItem "��ü" & Space(30) & "XX"
            .AddItem "������" & Space(30) & "1"
            .AddItem "������" & Space(30) & "0"
            
            .ListIndex = 0
        End With
        
        OPTIONS(11).Caption = "ǥ������"
        
        '2011-01-10 ���ѿ� ���� Ȳ���� ���� ��û
        
        Select Case Trim(basModule.SchCD)
            Case "K", "W", "Q"
                OPTIONS(4).Caption = "�� �� ��"
            Case Else
                OPTIONS(4).Caption = "��   ��"
        End Select
        
'        Select Case Trim(basModule.SchCD)
'            Case "K", "W", "Q", "J"
'                OPTIONS(11).Caption = "����"
'                OPTIONS(4).Caption = "ǥ��"
'            Case "M"
'                OPTIONS(11).Caption = "ǥ������"
'                OPTIONS(4).Caption = "��   ��"
'            Case Else
'                'NO ACTION
'        End Select
        
        
        '>> �迭
        With cboKaeyol
            .Clear
            .AddItem "��ü" & Space(30) & "XX"
            
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
            If Trim(basModule.SchCD) = "K" Then             '< ����
                .AddItem "�ָ�����" & Space(30) & "04"
                .AddItem "�ָ��Ǵ�" & Space(30) & "05"
            
                .AddItem "�߰�����" & Space(30) & "06"
                .AddItem "�߰��Ǵ�" & Space(30) & "07"
            
                .AddItem "�������ι�" & Space(30) & "11"
                .AddItem "�������ڿ�" & Space(30) & "12"
                
                .AddItem "�������ι�16" & Space(30) & "16"
                .AddItem "�������ڿ�17" & Space(30) & "17"
                
            End If
        '<< �迭 >> : 2009.01.08
            Select Case Trim(basModule.SchCD)
                Case "S", "P"
'                    .AddItem "��ü��" & Space(30) & "03"
'
'                    .AddItem "�����ι�" & Space(30) & "05"
'                    .AddItem "�����ڿ�" & Space(30) & "06"
                    
                    .AddItem "�ι������̾�" & Space(30) & "18"
                    .AddItem "�ڿ������̾�" & Space(30) & "19"
                    
            End Select
            
            Select Case Trim(basModule.SchCD)
                Case "J"
                    .AddItem "��ü��" & Space(30) & "03"
                    
                    .AddItem "�ż��ι�" & Space(30) & "11"
                    .AddItem "�ż��ڿ�" & Space(30) & "12"
                    
                    .AddItem "�ι������̾�" & Space(30) & "18"
                    .AddItem "�ڿ������̾�" & Space(30) & "19"
                    
            End Select
            
        '<< �迭 >> : 2009.01.09
            If Trim(basModule.SchCD) = "B" Then             '< �λ�
                
                .AddItem "���м����ι�" & Space(30) & "05"
                .AddItem "���м����ڿ�" & Space(30) & "06"
                
                .AddItem "��.�����ι�" & Space(30) & "07"
                .AddItem "��.�����ڿ�" & Space(30) & "08"
                
                .AddItem "��ȭ�ι�" & Space(30) & "09"
                .AddItem "��ȭ�ڿ�" & Space(30) & "10"
                
            End If
            
            .ListIndex = 0
        End With
        
        txtStdNM.Text = ""
        
        '>> ���ͳ�/�п� ����
        With cboinGbn
            .Clear
            .AddItem "��ü" & Space(30) & "ALL"
            .AddItem "���ͳ�" & Space(30) & "INT"
            .AddItem "�п�" & Space(30) & "HAK"
            
            .ListIndex = 0
        End With
        
        '>> �����/ ���չ� ����
        With cboSel
            .Clear
            '.AddItem "����" & Space(30) & "01"
            .AddItem "����" & Space(30) & "02"
            
            .ListIndex = 0
        End With
        
        ReDim uSTD(0) As tSTD
        
    Me.Tag = ""
    
End Sub

Private Sub Clear_Form_Control()
    Dim UsrCtl      As Control
    For Each UsrCtl In Me
        With UsrCtl
             If UCase(TypeName(UsrCtl)) = "TEXTBOX" Then .Text = ""
             If UCase(TypeName(UsrCtl)) = "LINE" Then .BorderColor = &H0
             If UCase(TypeName(UsrCtl)) = "SHAPE" Then .BorderColor = &H0
        End With
    Next
    
    �л�����.Tag = ""
    �����ȣ.Tag = ""
    
    ������_����.Text = ""
    ������_����.Text = ""
    ������_����.Text = ""
    
    'Height = 3990
    'Width = 4890   ' ���̿� �ʺ� �����մϴ�.
    Set Photo.Picture = imgList.ListImages.Item(1).Picture
    
    OPTIONS(3).Visible = False
    OPTIONS(5).Visible = False
    
    
'>> �г⺰ ����
    Select Case Trim(basModule.SchCD)
    
        Case "N"
            OPTIONS(0) = "���ι��� �л����� ��ȸŽ�� 11���� �� 4������� ������ �� ������, ��2�ܱ���� 6���� �� 1������ ������ �� �ֽ��ϴ�."
            OPTIONS(1) = "���ڿ��� �л����� ������������ 1����, ����Ž�������� 4������� ������ �� �ֽ��ϴ�."
            OPTIONS(2) = "���ڿ��� �л� �� ����(��)���� �����ϴ� �л��� ���� ǥ��"
            
            Labels(2).Caption = "�� ������ �ȿ��� �����Ͻÿ�."
            Labels(78).Caption = "����� �������� ��û�ڰ� ���� ���"
            Labels(1).Caption = "���� �������� ���� ���� �ֽ��ϴ�."
            
           
            
            OPTIONS(3).Visible = False
'            FillBOXs_opt(0).Visible = False
'            Lines_opt(3).Visible = False
            'Lines_opt(0).Visible = False
            
            OPTIONS(21).Visible = False
'            FillBOXs_opt(21).Visible = False
            'Lines_opt(21).Visible = False
            'Lines_opt(22).Visible = False
        
        Case "K"
            OPTIONS(0) = "���ι��� ��ȸŽ�� ���� �� ��2�л��� ��2�ܱ�� �������� ���� �л��� ����, ����, ����(5�� ���� ����)�� ������ �� �ֽ��ϴ�."
            OPTIONS(1) = ""
            OPTIONS(2) = ""
            
            Labels(2).Caption = "�� ������ �ȿ��� �����Ͻÿ�."
            Labels(78).Caption = "�������� �������� �����п� �ܿ� �ٸ� �п�����"
            Labels(1).Caption = "  ������ ���� ��� 2������ ǥ���Ͻÿ�."
            
           
            
            Labels(46).Caption = "2012 ����"
            
'            OPTIONS(21).Visible = False
            
        Case "W", "Q"
            OPTIONS(0) = "���ι��� ��ȸŽ�� ���� �� ��2�л��� ��2�ܱ�� �������� ���� �л��� ����, ����, ����(5�� ���� ����)�� ������ �� �ֽ��ϴ�."
            OPTIONS(1) = ""
            OPTIONS(2) = ""
            OPTIONS(21) = ""
            
            Labels(2).Caption = "�� ������ �ȿ��� �����Ͻÿ�."
            Labels(78).Caption = "�������� �������� �����п� �ܿ� �ٸ� �п�����"
            Labels(1).Caption = "  ������ ���� ��� 2������ ǥ���Ͻÿ�."
            
            
        
        Case "J", "B"
            OPTIONS(0) = ""
            OPTIONS(1) = ""
            OPTIONS(2) = ""
            OPTIONS(21) = "�ؼ���� �����л��� I���� 2����, I���� ������ ���� �� II 1���� ����"
            Labels(2).Caption = "�� ������ �ȿ��� �����Ͻÿ�.(*�� �ʼ������̰� �� �ܿ��� ���������Դϴ�."
            Labels(78).Caption = ""
            Labels(1).Caption = ""
            
            
            
'            OPTIONS(3).Visible = False
'            FillBOXs_opt(0).Visible = False
            'Lines_opt(3).Visible = False
            'Lines_opt(0).Visible = False
            
            'OPTIONS(21).Visible = False
'            FillBOXs_opt(21).Visible = False
            'Lines_opt(21).Visible = False
            'Lines_opt(22).Visible = False
            
            'Boxs(4).Visible = False
            'FillBOXs(3).Visible = False
            'Lines(25).Visible = False
            
        Case "S"
            OPTIONS(0) = "���ι��� �л����� ��ȸŽ�������� 4������ �����Ͽ��� �մϴ�."
            OPTIONS(1) = "����2�ܱ��� �ð��� ��2�ܱ�� ���� �ʴ� �л����� ���� ����, �����, ���������� ���ð��뿡 �����ϴ� 1������ �����Ͻñ� �ٶ��ϴ�."
            OPTIONS(2) = "���ڿ��� �л����� ���񥰿����� 3����, ���� �������� 1������ �����ؾ� �մϴ�."
            
            Labels(2).Caption = "�� ������ �ȿ��� �����Ͻÿ�."
            Labels(78).Caption = ""
            Labels(1).Caption = ""
            
            
'            OPTIONS(3).Visible = False
'            FillBOXs_opt(0).Visible = False
            'Lines_opt(3).Visible = False
            'Lines_opt(0).Visible = False
            
            'OPTIONS(21).Visible = False
'            FillBOXs_opt(21).Visible = False
            'Lines_opt(21).Visible = False
            'Lines_opt(22).Visible = False
            
            'Boxs(4).Visible = False
            'FillBOXs(3).Visible = False
            'Lines(25).Visible = False
            
        Case "P"
            OPTIONS(0) = "���ι��� �л����� ��ȸŽ�������� 4������ �����Ͽ��� �մϴ�."
            OPTIONS(1) = "���ڿ��� �л����� ���񥰿����� 3����, ���� �������� 1������ �����ؾ� �մϴ�."
            OPTIONS(2) = ""
            
            Labels(2).Caption = ""
            Labels(78).Caption = ""
            Labels(1).Caption = ""
            
'            Boxs(4).Visible = False
'
'            Labels(2).Visible = False
'            Labels(78).Visible = False
'            Labels(1).Visible = False
'
'            FillBOXs(3).Visible = False
'            Lines(25).Visible = False
            
        Case "M"
            OPTIONS(21) = "��I���� 3���� ���� �Ǵ� I���� 2����, II���� 1���� ����"
            OPTIONS(0) = "���ι��� ��ȸŽ�� ���� �� ��2 �л���, ��2�ܱ�� �������� ���� �л��� ���ð��뿡 ���, ����, �ܱ���(����)�� ������ �� �ֽ��ϴ�."
            OPTIONS(1) = ""
            OPTIONS(2) = ""
            
            Labels(2).Caption = ""
            Labels(78).Caption = ""
            Labels(1).Caption = ""
            
    End Select
    
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
            VScroll1.value = nS - 1
            VScroll1.Enabled = False
                Call Std_Data_Show(VScroll1.value)
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
            VScroll1.value = nS + 1
            VScroll1.Enabled = False
                Call Std_Data_Show(VScroll1.value)
            VScroll1.Enabled = True
        End If
    End If
End Sub



'>> �л� ��ȸ
Private Sub cmdFind_Click()
    
    On Error GoTo ErrStmt
    Me.MousePointer = vbHourglass
    
    ReDim uSTD(0) As tSTD
    
    cmdFind.Enabled = False
        Call Get_STD_Data
        
    cmdFind.Enabled = True
    
    Me.MousePointer = vbDefault
    Exit Sub
ErrStmt:
    Me.MousePointer = vbDefault
    MsgBox "�л���ȸ�� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "�л���ȸ"
    On Error GoTo 0

End Sub

Private Sub Get_STD_Data()
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    
    Dim nLength     As Long
    Dim sStr        As String
    Dim ni          As Integer
    Dim nRec        As Long
    
    Dim sTmp        As String
    Dim nTmp        As Long
    
    Dim sFilePath   As String
    
    
    On Error GoTo ErrStmt
    
    sStr = ""
    sStr = sStr & "  SELECT ROWNUM AS ID, "
    sStr = sStr & "         SCHNO      , ACID       , EXMID      , STDNM      , SUBSTR(Birth_ymd,1,4)||'-'||SUBSTR(Birth_ymd,5,2) ||'-'||SUBSTR(Birth_ymd,7,2) AS Birth_ymd,"
    sStr = sStr & "         EXMTYPE    , KAEYOL     ,"
    sStr = sStr & "         SEL1       , SEL2       , SEL3       , SEL4       , SEL5       ,"
    sStr = sStr & "         K_NUM      , M_NUM      , E_NUM      , TOT_NUM    ,"
    sStr = sStr & "         SEL1_SCH   , SEL2_SCH   ,"
    sStr = sStr & "         PASS1      , PASS2      , PASS3      , PASS4      , CL_CLOSE   ,"
    sStr = sStr & "         CY_ACNT    , TOT_AMT    ,"
    sStr = sStr & "         BASE_AMT1  , BASE_AMT2  , BASE_AMT3  , BASE_AMT4  , "
    sStr = sStr & "         BASE_AMT5  , BASE_AMT6  , BASE_AMT7  , BASE_AMT8  , BASE_AMT9  , BASE_AMT10 ,"
    sStr = sStr & "         TAMGU_AMT1 , TAMGU_AMT2 , TAMGU_AMT3 , TAMGU_AMT4 , TAMGU_AMT5 ,"
    sStr = sStr & "         TAMGU_AMT6 , TAMGU_AMT7 , TAMGU_AMT8 , TAMGU_AMT9 , TAMGU_AMT10, TAMGU_AMT11, TAMGU_AMT12,"
    sStr = sStr & "         DECODE(SEX,'M','��','F','��') AS SEX        , "
    sStr = sStr & "         SUBSTR(ZIP,1,3)||'-'||SUBSTR(ZIP,4,3) AS ZIP, ADDR1      , ADDR2      ,"
    sStr = sStr & "         TEL        , CEL        , EMAIL      ,"
    sStr = sStr & "         HIGH_SCH   , GRADE_YEAR ,"
    sStr = sStr & "         PRNT_NM    , DECODE(PRNT_RLTN,'1','��','2','��','3',' ') AS PRNT_RLTN, "
    sStr = sStr & "         SUBSTR(PRNT_ZIP,1,3)||'-'||SUBSTR(PRNT_ZIP,4,3) AS PRNT_ZIP, PRNT_ADDR1 , PRNT_ADDR2 ,"
    sStr = sStr & "         PRNT_TEL   , PRNT_CEL   , PRNT_JOB   , PRNT_W_TEL ,"
    sStr = sStr & "         PHOTO_PATH , DECODE(R_WAY,'1','','2','-int','3','') AS R_WAY, PTS_SEL, ORD_NO, "
    sStr = sStr & "         ACID||EXMID AS IMAGE_FILE, "
    sStr = sStr & "         DECODE(ACID,'" & Trim(basModule.SchCD) & "','',ACID) AS WANT_ACID, "
    
    
    sStr = sStr & "         DECODE(GR,'1','���ɵ��','2','6�� �򰡿�','3','9�� �򰡿�','4','6�� �򰡿�','5','9�� �򰡿�','') AS GR, "            '<< 2009�� �ٲﳻ��
'    Select Case Trim(basModule.SchCD)
'        Case "S"
'            sStr = sStr & " DECODE(GR,'1','���ɵ��','2','2009 ��','','') AS GR, "
'        Case "P"
'            sStr = sStr & " DECODE(GR,'8','���ɵ��','9','2009 ��','6','3���','','') AS GR, "
'        Case Else
'            sStr = sStr & " '' AS GR, "
'    End Select
    
    'sStr = sStr & "         DECODE(ACID,'" & Trim(basModule.SchCD) & "','" & Trim(basModule.SchCD) & "',ACID) AS WANT_ACID "       '< TEST
    
    '****************************** < IMAGE ���� ���丮 > **********************************************
    Select Case basModule.SchCD
        Case "N"                '< �뷮��
            sStr = sStr & "'" & Noryangjin & "'||"
        Case "K", "W", "Q"      '< ����
            sStr = sStr & "'" & Kangnam & "'||"
        Case "S"                '< ����
            sStr = sStr & "'" & Songpa & "'||"
        Case "P"                '< ���ĸ��̸�
            sStr = sStr & "'" & MSongpa & "'||"
        Case "M"                '< �������̸�
            sStr = sStr & "'" & MKangnam & "'||"
        Case "J"                '< ����
            sStr = sStr & "'" & MGwanghwa & "'||"
        Case "B"                '< �λ� ���̸�
            sStr = sStr & "'" & Busan & "'||"
        
    End Select
                            sStr = sStr & "DECODE("
                                    sStr = sStr & "     KAEYOL||EXMTYPE,"
                                    sStr = sStr & "         '010','1A',"
                                    sStr = sStr & "         '011','1B',"
                                    sStr = sStr & "         '020','2A',"
                                    sStr = sStr & "         '021','2B',"
                                    sStr = sStr & "         '030','3A',"
                                    sStr = sStr & "         '031','3B',"
                                    sStr = sStr & "         '040','4A',"
                                    sStr = sStr & "         '041','4B',"
                                    sStr = sStr & "         '050','ETC',"
                                    sStr = sStr & "         '051','5B',"
                                    sStr = sStr & "         '060','6A',"
                                    sStr = sStr & "         '061','6B',"
                                    sStr = sStr & "         '070','7A',"
                                    sStr = sStr & "         '071','7B',"
                                    sStr = sStr & "         '080','8A',"
                                    sStr = sStr & "         '081','8B',"
                                    sStr = sStr & "         '090','9A',"
                                    sStr = sStr & "         '091','9B',"
                                    
                                    sStr = sStr & "         '110','1A',"
                                    sStr = sStr & "         '111','1B',"
                                    sStr = sStr & "         '120','2A',"
                                    sStr = sStr & "         '121','2B',"
                                    sStr = sStr & "         '130','3A',"
                                    sStr = sStr & "         '131','3B',"
                                    sStr = sStr & "         '140','4A',"
                                    sStr = sStr & "         '141','4B',"
                                    sStr = sStr & "         '150','ETC',"
                                    sStr = sStr & "         '151','5B',"
                                    sStr = sStr & "         '160','6A',"
                                    sStr = sStr & "         '161','6B',"
                                    
                                    sStr = sStr & "         '170','7A',"
                                    
                                    sStr = sStr & "         '180','1A',"
                                    sStr = sStr & "         '190','1B'"
                                    
                            sStr = sStr & "       )||'/'||ORD_NO||'.jpg' AS IMAGE_DIR "
                            
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
    '******************************************************************************************************
    
    sStr = sStr & "    FROM ( "
    
            sStr = sStr & "  SELECT A.SCHNO         ,"
            sStr = sStr & "         MAX(ACID      ) AS ACID       ,"
            sStr = sStr & "         MAX(EXMID     ) AS EXMID      ,"
            sStr = sStr & "         MAX(STDNM     ) AS STDNM      ,"
            sStr = sStr & "         MAX(Birth_ymd     ) AS Birth_ymd      ,"
            sStr = sStr & "         MAX(EXMTYPE   ) AS EXMTYPE    , MAX(KAEYOL    ) AS KAEYOL     ,"
            sStr = sStr & "         MAX(SEL1      ) AS SEL1       , MAX(SEL2      ) AS SEL2       , MAX(SEL3      ) AS SEL3      , MAX(SEL4      ) AS SEL4      , MAX(SEL5      ) AS  SEL5      ,"
            sStr = sStr & "         MAX(K_NUM     ) AS K_NUM      , MAX(M_NUM     ) AS M_NUM      , MAX(E_NUM     ) AS E_NUM     , MAX(TOT_NUM   ) AS TOT_NUM   ,"
            sStr = sStr & "         MAX(SEL1_SCH  ) AS SEL1_SCH   , MAX(SEL2_SCH  ) AS SEL2_SCH   ,"
            sStr = sStr & "         MAX(PASS1     ) AS PASS1      , MAX(PASS2     ) AS PASS2      , MAX(PASS3     ) AS PASS3     , MAX(PASS4     ) AS PASS4     , MAX(CL_CLOSE  ) AS  CL_CLOSE  ,"
            sStr = sStr & "         MAX(CY_ACNT   ) AS CY_ACNT    , MAX(TOT_AMT   ) AS TOT_AMT    ,"
            sStr = sStr & "         MAX(BASE_AMT1 ) AS BASE_AMT1  , MAX(BASE_AMT2 ) AS BASE_AMT2  , MAX(BASE_AMT3 ) AS BASE_AMT3 , MAX(BASE_AMT4 ) AS BASE_AMT4 ,"
            sStr = sStr & "         MAX(BASE_AMT5 ) AS BASE_AMT5  , MAX(BASE_AMT6 ) AS BASE_AMT6  , MAX(BASE_AMT7 ) AS BASE_AMT7 , MAX(BASE_AMT8 ) AS BASE_AMT8 , MAX(BASE_AMT9 ) AS BASE_AMT9  , MAX(BASE_AMT10) AS BASE_AMT10   ,"
            sStr = sStr & "         MAX(TAMGU_AMT1) AS TAMGU_AMT1 , MAX(TAMGU_AMT2) AS TAMGU_AMT2 , MAX(TAMGU_AMT3) AS TAMGU_AMT3, MAX(TAMGU_AMT4) AS TAMGU_AMT4, MAX(TAMGU_AMT5) AS  TAMGU_AMT5,"
            sStr = sStr & "         MAX(TAMGU_AMT6) AS TAMGU_AMT6 , MAX(TAMGU_AMT7) AS TAMGU_AMT7 , MAX(TAMGU_AMT8) AS TAMGU_AMT8, MAX(TAMGU_AMT9) AS TAMGU_AMT9, MAX(TAMGU_AMT10) AS TAMGU_AMT10, MAX(TAMGU_AMT11) AS TAMGU_AMT11, MAX(TAMGU_AMT12) AS TAMGU_AMT12,"
            sStr = sStr & "         MAX(SEX       ) AS SEX        ,"
            sStr = sStr & "         MAX(ZIP       ) AS ZIP        , MAX(ADDR1     ) AS ADDR1      , MAX(ADDR2     ) AS ADDR2     ,"
            sStr = sStr & "         MAX(TEL       ) AS TEL        , MAX(CEL       ) AS CEL        , MAX(EMAIL     ) AS EMAIL     ,"
            sStr = sStr & "         MAX(HIGH_SCH  ) AS HIGH_SCH   , MAX(GRADE_YEAR) AS GRADE_YEAR ,"
            sStr = sStr & "         MAX(PRNT_NM   ) AS PRNT_NM    , MAX(PRNT_RLTN ) AS PRNT_RLTN  ,"
            sStr = sStr & "         MAX(PRNT_ZIP  ) AS PRNT_ZIP   , MAX(PRNT_ADDR1) AS PRNT_ADDR1 , MAX(PRNT_ADDR2) AS PRNT_ADDR2,"
            sStr = sStr & "         MAX(PRNT_TEL  ) AS PRNT_TEL   , MAX(PRNT_CEL  ) AS PRNT_CEL   , MAX(PRNT_JOB  ) AS PRNT_JOB  , MAX(PRNT_W_TEL) AS PRNT_W_TEL,"
            sStr = sStr & "         MAX(PHOTO_PATH) AS PHOTO_PATH , MAX(R_WAY     ) AS R_WAY      , MAX(PTS_SEL   ) AS PTS_SEL   , MAX(ORD_NO    ) AS ORD_NO    , MAX(MU_TYPE) AS GR "
            
                    '2010.12.14 ���� doubleó�� �Ǽ� MAX�� ���� ���ѿ�
                    
                    sStr = sStr & " , "
                    sStr = sStr & "        MAX(J01) AS J01,"
                    sStr = sStr & "        MAX(K01) AS K01,"
                    sStr = sStr & "        MAX(J02) AS J02,"
                    sStr = sStr & "        MAX(K02) AS K02,"
                    sStr = sStr & "        MAX(J03) AS J03,"
                    sStr = sStr & "        MAX(K03) AS K03,"
                    
                    sStr = sStr & "        MAX(J04) AS J04,"
                    sStr = sStr & "        MAX(K04) AS K04,"
                    sStr = sStr & "        MAX(J05) AS J05,"
                    sStr = sStr & "        MAX(K05) AS K05,"
                    sStr = sStr & "        MAX(J06) AS J06,"
                    sStr = sStr & "        MAX(K06) AS K06,"
                    sStr = sStr & "        MAX(J07) AS J07,"
                    sStr = sStr & "        MAX(K07) AS K07,"
                    sStr = sStr & "        MAX(J08) AS J08,"
                    sStr = sStr & "        MAX(K08) AS K08,"
                    sStr = sStr & "        MAX(J09) AS J09,"
                    sStr = sStr & "        MAX(K09) AS K09,"
                    sStr = sStr & "        MAX(J10) AS J10,"
                    sStr = sStr & "        MAX(K10) AS K10,"
                    sStr = sStr & "        MAX(J11) AS J11,"
                    sStr = sStr & "        MAX(K11) AS K11,"
                    
                    sStr = sStr & "        MAX(J12) AS J12,"
                    sStr = sStr & "        MAX(K12) AS K12,"
                    sStr = sStr & "        MAX(J13) AS J13,"
                    sStr = sStr & "        MAX(K13) AS K13,"
                    sStr = sStr & "        MAX(J14) AS J14,"
                    sStr = sStr & "        MAX(K14) AS K14,"
                    
                    sStr = sStr & "        MAX(J15) AS J15,"
                    sStr = sStr & "        MAX(K15) AS K15,"
                    sStr = sStr & "        MAX(J16) AS J16,"
                    sStr = sStr & "        MAX(K16) AS K16,"
                    sStr = sStr & "        MAX(J17) AS J17,"
                    sStr = sStr & "        MAX(K17) AS K17,"
                    sStr = sStr & "        MAX(J18) AS J18,"
                    sStr = sStr & "        MAX(K18) AS K18,"
                    
                    sStr = sStr & "        MAX(J19) AS J19,"
                    sStr = sStr & "        MAX(K19) AS K19,"
                    sStr = sStr & "        MAX(J20) AS J20,"
                    sStr = sStr & "        MAX(K20) AS K20,"
                    sStr = sStr & "        MAX(J21) AS J21,"
                    sStr = sStr & "        MAX(K21) AS K21"
                    
            sStr = sStr & "    FROM ("
            '---------------------------------------------------------------------------- ��ü�л� ��ȸ START
            sStr = sStr & "          SELECT *"
            sStr = sStr & "            FROM CLSTD01TB"
            sStr = sStr & "           WHERE ACID = '" & Trim(basModule.SchCD) & "'"
            '>> �����ȣ
            If Trim(fpExmID_S.UnFmtText) > " " And Trim(fpExmID_E.UnFmtText) > " " Then
                sStr = sStr & "         AND EXMID BETWEEN '" & Trim(fpExmID_S.UnFmtText) & "' AND '" & Trim(fpExmID_E.UnFmtText) & "'"
            ElseIf Trim(fpExmID_S.UnFmtText) > " " And Trim(fpExmID_E.UnFmtText) = " " Then
                sStr = sStr & "         AND EXMID BETWEEN '" & Trim(fpExmID_S.UnFmtText) & "' AND '99999' "
            ElseIf Trim(fpExmID_S.UnFmtText) = " " And Trim(fpExmID_E.UnFmtText) > " " Then
                sStr = sStr & "         AND EXMID BETWEEN '00000' AND '" & Trim(fpExmID_E.UnFmtText) & "'"
            ElseIf Trim(fpExmID_S.UnFmtText) = " " And Trim(fpExmID_E.UnFmtText) = " " Then
                ' no action
            End If
            
            '>> ��/������ üũ
            If Trim(Right(cboExmType.Text, 30)) = "0" Then
                sStr = sStr & "         AND EXMTYPE = '0'"
            ElseIf Trim(Right(cboExmType.Text, 30)) = "1" Then
                sStr = sStr & "         AND EXMTYPE = '1'"
            End If
            
            '>> �迭
            Select Case Trim(basModule.SchCD)
                Case "K", "S", "P"
                    If Trim(Right(cboKaeyol.Text, 30)) <> "XX" Then
                        sStr = sStr & "     AND KAEYOL  = '" & Trim(Right(cboKaeyol.Text, 30)) & "'"
                    End If
                Case Else
                    Select Case Trim(Right(cboKaeyol, 30))
                        Case "XX"
                            ' no action
                        Case "01", "03", "11", "13"
                            sStr = sStr & "     AND SEL1 > ' ' "
                        Case "02", "12"
                            sStr = sStr & "     AND SEL3 > ' ' "
                        Case "04", "05", "06", "07", "08", "09", "10", "14", "15", "16"
                            sStr = sStr & "     AND KAEYOL = '" & Trim(Right(cboKaeyol.Text, 30)) & "'"
                    End Select
            End Select
            
            '>> �л���
            If Trim(txtStdNM.Text) > " " Then
                sStr = sStr & "         AND STDNM LIKE '%" & Trim(txtStdNM.Text) & "%'"
            End If
            '>> ���ͳ�/�п�
            If Trim(Right(cboinGbn.Text, 30)) = "INT" Then          '< ���ͳ� ����
                sStr = sStr & "         AND R_WAY = '2'"
            ElseIf Trim(Right(cboinGbn.Text, 30)) = "HAK" Then      '< �п� ����
                sStr = sStr & "         AND R_WAY IN ('1','3') "
            End If
            sStr = sStr & "             AND EXMID > ' ' "
            
            sStr = sStr & "             AND BIGO2 IS NULL"          '< 2008.12. ���ɺ� �л��� �⵵�� ���� �ƴϸ� NULL
            
            sStr = sStr & "          UNION ALL"
            '---------------------------------------------------------------------------- ��ü�л� ��ȸ END
            '---------------------------------------------------------------------------- �հ��� ��ȸ START
            sStr = sStr & "          SELECT *"
            sStr = sStr & "            From CLSTD01TB"
            sStr = sStr & "           WHERE (PASS1 = '" & Trim(basModule.SchCD) & "'" & " OR"
            sStr = sStr & "                  PASS2 = '" & Trim(basModule.SchCD) & "'" & " OR"
            sStr = sStr & "                  PASS3 = '" & Trim(basModule.SchCD) & "'" & " OR"
            sStr = sStr & "                  PASS4 = '" & Trim(basModule.SchCD) & "'" & " )"
            sStr = sStr & "             AND EXMID > ' ' "
            '>> ��/������ üũ
            If Trim(Right(cboExmType.Text, 30)) = "0" Then
                sStr = sStr & "         AND EXMTYPE = '0'"
            ElseIf Trim(Right(cboExmType.Text, 30)) = "1" Then
                sStr = sStr & "         AND EXMTYPE = '1'"
            End If
            '>> �迭
            Select Case Trim(basModule.SchCD)
                Case "K", "S", "P"
                    If Trim(Right(cboKaeyol.Text, 30)) <> "XX" Then
                        sStr = sStr & "     AND KAEYOL  = '" & Trim(Right(cboKaeyol.Text, 30)) & "'"
                    End If
                Case Else
                    Select Case Trim(Right(cboKaeyol, 30))
                        Case "XX"
                            ' no action
                        Case "01", "03", "11", "13"
                            sStr = sStr & "     AND SEL1 > ' ' "
                        Case "02", "12"
                            sStr = sStr & "     AND SEL3 > ' ' "
                        Case "04", "05", "06", "07", "08", "09", "10", "14", "15", "16"
                            sStr = sStr & "     AND KAEYOL = '" & Trim(Right(cboKaeyol.Text, 30)) & "'"
                    End Select
            End Select
            
            '>> �л���
            If Trim(txtStdNM.Text) > " " Then
                sStr = sStr & "         AND STDNM LIKE '%" & Trim(txtStdNM.Text) & "%'"
            End If
            '>> ���ͳ�/�п�
            If Trim(Right(cboinGbn.Text, 30)) = "INT" Then          '< ���ͳ� ����
                sStr = sStr & "         AND R_WAY = '2'"
            ElseIf Trim(Right(cboinGbn.Text, 30)) = "HAK" Then      '< �п� ����
                sStr = sStr & "         AND R_WAY IN ('1','3') "
            End If
            
            sStr = sStr & "             AND BIGO2 IS NULL"          '< 2008.12. ���ɺ� �л��� �⵵�� ���� �ƴϸ� NULL
    
            sStr = sStr & "          ) A, "
            
            
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
                    sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(2) & "', DECODE(SUB_NUM,'X',0, SUB_NUM), '53', DECODE(SUB_NUM, 'X',0, SUB_NUM), 0) AS J06,    /* ��Ž-" & constSatams(2) & "         , ��Ž-��������1             */"
                    sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(2) & "', DECODE(SUB_BAK,'X',0, SUB_BAK), '53', DECODE(SUB_BAK, 'X',0, SUB_BAK), 0) AS K06,    /* �����  ��Ž-" & constSatams(2) & "         , ��Ž-��������1     */"
                    sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(3) & "', DECODE(SUB_NUM,'X',0, SUB_NUM), '54', DECODE(SUB_NUM, 'X',0, SUB_NUM), 0) AS J07,    /* ��Ž-" & constSatams(3) & "   , ��Ž-��������1         */"
                    sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(3) & "', DECODE(SUB_BAK,'X',0, SUB_BAK), '54', DECODE(SUB_BAK, 'X',0, SUB_BAK), 0) AS K07,    /* �����  ��Ž-" & constSatams(3) & "   , ��Ž-��������1 */"
                    sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(4) & "', DECODE(SUB_NUM,'X',0, SUB_NUM), '55', DECODE(SUB_NUM, 'X',0, SUB_NUM), 0) AS J08,    /* ��Ž-" & constSatams(4) & "       , ��Ž-����2             */"
                    sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(4) & "', DECODE(SUB_BAK,'X',0, SUB_BAK), '55', DECODE(SUB_BAK, 'X',0, SUB_BAK), 0) AS K08,    /* �����  ��Ž-" & constSatams(4) & "       , ��Ž-����2     */"
                    sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(5) & "', DECODE(SUB_NUM,'X',0, SUB_NUM), '56', DECODE(SUB_NUM, 'X',0, SUB_NUM), 0) AS J09,    /* ��Ž-" & constSatams(5) & "     , ��Ž-ȭ��2             */"
                    sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(5) & "', DECODE(SUB_BAK,'X',0, SUB_BAK), '56', DECODE(SUB_BAK, 'X',0, SUB_BAK), 0) AS K09,    /* �����  ��Ž-" & constSatams(5) & "     , ��Ž-ȭ��2     */"
                    sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(6) & "', DECODE(SUB_NUM,'X',0, SUB_NUM), '57', DECODE(SUB_NUM, 'X',0, SUB_NUM), 0) AS J10,      /* ��Ž-" & constSatams(6) & "     , ��Ž-��������2           */"
                    sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(6) & "', DECODE(SUB_BAK,'X',0, SUB_BAK), '57', DECODE(SUB_BAK, 'X',0, SUB_BAK), 0) AS K10,      /* ����� ��Ž-" & constSatams(6) & "     , ��Ž-��������2    */"
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
            
            sStr = sStr & "   GROUP BY A.SCHNO"
            '---------------------------------------------------------------------------- �հ��� ��ȸ END
    
    sStr = sStr & "    ) "
    
    
    
    sStr = sStr & " WHERE SCHNO > ' ' "
    
    
    '>> �����ȣ
    If Trim(fpExmID_S.UnFmtText) > " " And Trim(fpExmID_E.UnFmtText) > " " Then
        sStr = sStr & " AND EXMID BETWEEN '" & Trim(fpExmID_S.UnFmtText) & "' AND '" & Trim(fpExmID_E.UnFmtText) & "'"
    ElseIf Trim(fpExmID_S.UnFmtText) > " " And Trim(fpExmID_E.UnFmtText) = " " Then
        sStr = sStr & " AND EXMID BETWEEN '" & Trim(fpExmID_S.UnFmtText) & "' AND '99999' "
    ElseIf Trim(fpExmID_S.UnFmtText) = " " And Trim(fpExmID_E.UnFmtText) > " " Then
        sStr = sStr & " AND EXMID BETWEEN '00000' AND '" & Trim(fpExmID_E.UnFmtText) & "'"
    ElseIf Trim(fpExmID_S.UnFmtText) = " " And Trim(fpExmID_E.UnFmtText) = " " Then
        ' no action
    End If
    sStr = sStr & " ORDER BY EXMID "
    
    'Text1.Text = sStr
    
    Set DBCmd = New ADODB.Command
    Set DBRec = New ADODB.Recordset
    Set DBParam = New ADODB.Parameter
    
    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
    DBCmd.CommandText = sStr
    DBCmd.CommandType = adCmdText
    DBCmd.CommandTimeout = 30
    
    
    
''>> �п�
'        sTmp = Trim(basModule.SchCD)
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
'
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
''>> �л���
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
            
            ReDim uSTD(0) As tSTD
            
            For Each UsrCtl In Me
                With UsrCtl
                     If UCase(TypeName(UsrCtl)) = "TEXTBOX" Then .Text = ""
                     If UCase(TypeName(UsrCtl)) = "LINE" Then .BorderColor = &H0
                     If UCase(TypeName(UsrCtl)) = "SHAPE" Then .BorderColor = &H0
                End With
            Next
            
            Set Photo.Picture = imgList.ListImages.Item(1).Picture
            
            MsgBox "�ش���ȸ����ڰ� �����ϴ�.", vbExclamation + vbOKOnly, "������� ��ȸ"
            
        ElseIf .RecordCount > 0 Then
            nTotRec = .RecordCount
            
            .MoveFirst
            
            ReDim uSTD(.RecordCount) As tSTD
            
            VScroll1.Max = .RecordCount
            VScroll1.Enabled = True
            
            For nRec = 1 To .RecordCount Step 1
            
                sTmp = "":      If IsNull(.Fields("SCHNO")) = False Then sTmp = .Fields("SCHNO")
                    uSTD(nRec).SCHNO = sTmp
                sTmp = "":      If IsNull(.Fields("ACID")) = False Then sTmp = .Fields("ACID")
                    uSTD(nRec).ACID = sTmp
                sTmp = "":      If IsNull(.Fields("EXMID")) = False Then sTmp = .Fields("EXMID")
                    uSTD(nRec).EXMID = sTmp
                sTmp = "":      If IsNull(.Fields("STDNM")) = False Then sTmp = .Fields("STDNM")
                    uSTD(nRec).STDNM = sTmp
                sTmp = "":      If IsNull(.Fields("Birth_ymd")) = False Then sTmp = .Fields("Birth_ymd")
                    uSTD(nRec).Birth_ymd = sTmp
                
                sTmp = "":      If IsNull(.Fields("EXMTYPE")) = False Then sTmp = .Fields("EXMTYPE")
                    uSTD(nRec).EXMTYPE = sTmp
                sTmp = "":      If IsNull(.Fields("KAEYOL")) = False Then sTmp = .Fields("KAEYOL")
                    uSTD(nRec).KAEYOL = sTmp
                
                sTmp = "":      If IsNull(.Fields("SEL1")) = False Then sTmp = .Fields("SEL1")
                    uSTD(nRec).SEL1 = sTmp
                sTmp = "":      If IsNull(.Fields("SEL2")) = False Then sTmp = .Fields("SEL2")
                    uSTD(nRec).SEL2 = sTmp
                sTmp = "":      If IsNull(.Fields("SEL3")) = False Then sTmp = .Fields("SEL3")
                    uSTD(nRec).SEL3 = sTmp
                sTmp = "":      If IsNull(.Fields("SEL4")) = False Then sTmp = .Fields("SEL4")
                    uSTD(nRec).SEL4 = sTmp
                sTmp = "":      If IsNull(.Fields("SEL5")) = False Then sTmp = .Fields("SEL5")
                    uSTD(nRec).SEL5 = sTmp
                
                
                nTmp = 0:      If IsNumeric(.Fields("K_NUM")) = True Then nTmp = .Fields("K_NUM")
                    uSTD(nRec).K_NUM = nTmp
                nTmp = 0:      If IsNumeric(.Fields("M_NUM")) = True Then nTmp = .Fields("M_NUM")
                    uSTD(nRec).M_NUM = nTmp
                nTmp = 0:      If IsNumeric(.Fields("E_NUM")) = True Then nTmp = .Fields("E_NUM")
                    uSTD(nRec).E_NUM = nTmp
                nTmp = 0:      If IsNumeric(.Fields("TOT_NUM")) = True Then nTmp = .Fields("TOT_NUM")
                    uSTD(nRec).TOT_NUM = nTmp
                
                
                sTmp = "":      If IsNull(.Fields("SEL1_SCH")) = False Then sTmp = .Fields("SEL1_SCH")
                    uSTD(nRec).SEL1_SCH = sTmp
                    
                    Select Case uSTD(nRec).SEL1_SCH
                        Case "N"
                            uSTD(nRec).SEL1_SCH = "�뷮��"
                        Case "K"
                            uSTD(nRec).SEL1_SCH = "����"
                        Case "S"
                            uSTD(nRec).SEL1_SCH = "����"
                        Case "P"
                            uSTD(nRec).SEL1_SCH = "���� M"
                        Case "M"
                            uSTD(nRec).SEL1_SCH = "���� M"
                            
                        Case "W"
                            uSTD(nRec).SEL1_SCH = "�ָ����Ǵ�"
                        Case "Q"
                            uSTD(nRec).SEL1_SCH = "�߰����Ǵ�"
                        
                        Case "J"
                            uSTD(nRec).SEL1_SCH = "����"
                        Case "B"
                            uSTD(nRec).SEL1_SCH = "�λ�"
                            
                    End Select
                
                
                sTmp = "":      If IsNull(.Fields("SEL2_SCH")) = False Then sTmp = .Fields("SEL2_SCH")
                    uSTD(nRec).SEL2_SCH = sTmp
                    
                    '<< 2008.01.10 : �뷮�� - ���� ������ >>
                    If Trim(basModule.SchCD) = "N" Then
                        Select Case uSTD(nRec).KAEYOL
                            Case "05"
                                uSTD(nRec).SEL2_SCH = "�ι�"
                            Case "06"
                                uSTD(nRec).SEL2_SCH = "�ڿ�"
                            
                            Case Else
                                uSTD(nRec).SEL2_SCH = ""
                        End Select
                    Else
                        Select Case uSTD(nRec).SEL2_SCH
                            Case "N"
                                uSTD(nRec).SEL2_SCH = "�뷮��"
                            Case "K"
                                uSTD(nRec).SEL2_SCH = "����"
                            Case "S"
                                uSTD(nRec).SEL2_SCH = "����"
                            Case "P"
                                uSTD(nRec).SEL2_SCH = "���� M"
                            Case "M"
                                uSTD(nRec).SEL2_SCH = "���� M"
                                
                            Case "W"
                                uSTD(nRec).SEL2_SCH = "�ָ����Ǵ�"
                            Case "Q"
                                uSTD(nRec).SEL2_SCH = "�߰����Ǵ�"
                                
                            Case "J"
                                uSTD(nRec).SEL2_SCH = "����"
                            Case "B"
                                uSTD(nRec).SEL2_SCH = "�λ�"
                                
                        End Select
                    End If
                
                sTmp = "":      If IsNull(.Fields("PASS1")) = False Then sTmp = .Fields("PASS1")
                    uSTD(nRec).PASS1 = sTmp
                sTmp = "":      If IsNull(.Fields("PASS2")) = False Then sTmp = .Fields("PASS2")
                    uSTD(nRec).PASS2 = sTmp
                sTmp = "":      If IsNull(.Fields("PASS3")) = False Then sTmp = .Fields("PASS3")
                    uSTD(nRec).PASS3 = sTmp
                sTmp = "":      If IsNull(.Fields("PASS4")) = False Then sTmp = .Fields("PASS4")
                    uSTD(nRec).PASS4 = sTmp
                    
                
                sTmp = "":      If IsNull(.Fields("CL_CLOSE")) = False Then sTmp = .Fields("CL_CLOSE")
                    uSTD(nRec).CL_CLOSE = sTmp
                sTmp = "":      If IsNull(.Fields("CY_ACNT")) = False Then sTmp = .Fields("CY_ACNT")
                    uSTD(nRec).CY_ACNT = sTmp
                nTmp = 0:       If IsNull(.Fields("TOT_AMT")) = False Then nTmp = .Fields("TOT_AMT")
                    uSTD(nRec).TOT_AMT = nTmp
                
                
                nTmp = 0:       If IsNull(.Fields("BASE_AMT1")) = False Then nTmp = .Fields("BASE_AMT1")
                    uSTD(nRec).BASE_AMT1 = nTmp
                nTmp = 0:       If IsNull(.Fields("BASE_AMT2")) = False Then nTmp = .Fields("BASE_AMT2")
                    uSTD(nRec).BASE_AMT2 = nTmp
                nTmp = 0:       If IsNull(.Fields("BASE_AMT3")) = False Then nTmp = .Fields("BASE_AMT3")
                    uSTD(nRec).BASE_AMT3 = nTmp
                nTmp = 0:       If IsNull(.Fields("BASE_AMT4")) = False Then nTmp = .Fields("BASE_AMT4")
                    uSTD(nRec).BASE_AMT4 = nTmp
                nTmp = 0:       If IsNull(.Fields("BASE_AMT5")) = False Then nTmp = .Fields("BASE_AMT5")
                    uSTD(nRec).BASE_AMT5 = nTmp
                nTmp = 0:       If IsNull(.Fields("BASE_AMT6")) = False Then nTmp = .Fields("BASE_AMT6")
                    uSTD(nRec).BASE_AMT6 = nTmp
                    
                    
                nTmp = 0:       If IsNull(.Fields("BASE_AMT7")) = False Then nTmp = .Fields("BASE_AMT7")
                    uSTD(nRec).BASE_AMT7 = nTmp
                nTmp = 0:       If IsNull(.Fields("BASE_AMT8")) = False Then nTmp = .Fields("BASE_AMT8")
                    uSTD(nRec).BASE_AMT8 = nTmp
                nTmp = 0:       If IsNull(.Fields("BASE_AMT9")) = False Then nTmp = .Fields("BASE_AMT9")
                    uSTD(nRec).BASE_AMT9 = nTmp
                nTmp = 0:       If IsNull(.Fields("BASE_AMT10")) = False Then nTmp = .Fields("BASE_AMT10")
                    uSTD(nRec).BASE_AMT10 = nTmp
                                                                                                                                                  
                nTmp = 0:       If IsNull(.Fields("TAMGU_AMT1")) = False Then nTmp = .Fields("TAMGU_AMT1")
                    uSTD(nRec).TAMGU_AMT1 = nTmp
                nTmp = 0:       If IsNull(.Fields("TAMGU_AMT2")) = False Then nTmp = .Fields("TAMGU_AMT2")
                    uSTD(nRec).TAMGU_AMT2 = nTmp
                nTmp = 0:       If IsNull(.Fields("TAMGU_AMT3")) = False Then nTmp = .Fields("TAMGU_AMT3")
                    uSTD(nRec).TAMGU_AMT3 = nTmp
                nTmp = 0:       If IsNull(.Fields("TAMGU_AMT4")) = False Then nTmp = .Fields("TAMGU_AMT4")
                    uSTD(nRec).TAMGU_AMT4 = nTmp
                nTmp = 0:       If IsNull(.Fields("TAMGU_AMT5")) = False Then nTmp = .Fields("TAMGU_AMT5")
                    uSTD(nRec).TAMGU_AMT5 = nTmp
                nTmp = 0:       If IsNull(.Fields("TAMGU_AMT6")) = False Then nTmp = .Fields("TAMGU_AMT6")
                    uSTD(nRec).TAMGU_AMT6 = nTmp
                nTmp = 0:       If IsNull(.Fields("TAMGU_AMT7")) = False Then nTmp = .Fields("TAMGU_AMT7")
                    uSTD(nRec).TAMGU_AMT7 = nTmp
                nTmp = 0:       If IsNull(.Fields("TAMGU_AMT8")) = False Then nTmp = .Fields("TAMGU_AMT8")
                    uSTD(nRec).TAMGU_AMT8 = nTmp
                nTmp = 0:       If IsNull(.Fields("TAMGU_AMT9")) = False Then nTmp = .Fields("TAMGU_AMT9")
                    uSTD(nRec).TAMGU_AMT9 = nTmp
                nTmp = 0:       If IsNull(.Fields("TAMGU_AMT10")) = False Then nTmp = .Fields("TAMGU_AMT10")
                    uSTD(nRec).TAMGU_AMT10 = nTmp
                nTmp = 0:       If IsNull(.Fields("TAMGU_AMT11")) = False Then nTmp = .Fields("TAMGU_AMT11")
                    uSTD(nRec).TAMGU_AMT11 = nTmp
                nTmp = 0:       If IsNull(.Fields("TAMGU_AMT12")) = False Then nTmp = .Fields("TAMGU_AMT12")
                    uSTD(nRec).TAMGU_AMT12 = nTmp
                                                                                                                                                  
                sTmp = "":      If IsNull(.Fields("SEX")) = False Then sTmp = .Fields("SEX")
                    uSTD(nRec).SEX = sTmp
                                                                                                                                                  
                sTmp = "":      If IsNull(.Fields("ZIP")) = False Then sTmp = .Fields("ZIP")
                    uSTD(nRec).ZIP = sTmp
                sTmp = "":      If IsNull(.Fields("ADDR1")) = False Then sTmp = .Fields("ADDR1")
                    uSTD(nRec).ADDR1 = sTmp
                sTmp = "":      If IsNull(.Fields("ADDR2")) = False Then sTmp = .Fields("ADDR2")
                    uSTD(nRec).ADDR2 = sTmp
                sTmp = "":      If IsNull(.Fields("TEL")) = False Then sTmp = .Fields("TEL")
                    uSTD(nRec).TEL = sTmp
                sTmp = "":      If IsNull(.Fields("CEL")) = False Then sTmp = .Fields("CEL")
                    uSTD(nRec).CEL = sTmp
                sTmp = "":      If IsNull(.Fields("EMAIL")) = False Then sTmp = .Fields("EMAIL")
                    uSTD(nRec).EMAIL = sTmp
                                                                                                                                                  
                sTmp = "":      If IsNull(.Fields("HIGH_SCH")) = False Then sTmp = .Fields("HIGH_SCH")
                    uSTD(nRec).HIGH_SCH = sTmp
                sTmp = "":      If IsNull(.Fields("GRADE_YEAR")) = False Then sTmp = .Fields("GRADE_YEAR")
                    uSTD(nRec).GRADE_YEAR = sTmp
                                                                                                                                                  
                sTmp = "":      If IsNull(.Fields("PRNT_NM")) = False Then sTmp = .Fields("PRNT_NM")
                    uSTD(nRec).PRNT_NM = sTmp
                sTmp = "":      If IsNull(.Fields("PRNT_RLTN")) = False Then sTmp = .Fields("PRNT_RLTN")
                    uSTD(nRec).PRNT_RLTN = sTmp
                                                                                                                                                  
                sTmp = "":      If IsNull(.Fields("PRNT_ZIP")) = False Then sTmp = .Fields("PRNT_ZIP")
                    uSTD(nRec).PRNT_ZIP = sTmp
                sTmp = "":      If IsNull(.Fields("PRNT_ADDR1")) = False Then sTmp = .Fields("PRNT_ADDR1")
                    uSTD(nRec).PRNT_ADDR1 = sTmp
                sTmp = "":      If IsNull(.Fields("PRNT_ADDR2")) = False Then sTmp = .Fields("PRNT_ADDR2")
                    uSTD(nRec).PRNT_ADDR2 = sTmp
                sTmp = "":      If IsNull(.Fields("PRNT_TEL")) = False Then sTmp = .Fields("PRNT_TEL")
                    uSTD(nRec).PRNT_TEL = sTmp
                sTmp = "":      If IsNull(.Fields("PRNT_CEL")) = False Then sTmp = .Fields("PRNT_CEL")
                    uSTD(nRec).PRNT_CEL = sTmp
                sTmp = "":      If IsNull(.Fields("PRNT_JOB")) = False Then sTmp = .Fields("PRNT_JOB")
                    uSTD(nRec).PRNT_JOB = sTmp
                sTmp = "":      If IsNull(.Fields("PRNT_W_TEL")) = False Then sTmp = .Fields("PRNT_W_TEL")
                    uSTD(nRec).PRNT_W_TEL = sTmp
                                                                                                                                                  
                sTmp = "":      If IsNull(.Fields("PHOTO_PATH")) = False Then sTmp = .Fields("PHOTO_PATH")
                    uSTD(nRec).PHOTO_PATH = sTmp

                sTmp = "":      If IsNull(.Fields("R_WAY")) = False Then sTmp = .Fields("R_WAY")
                    uSTD(nRec).R_WAY = sTmp
                    
                sTmp = "":      If IsNull(.Fields("PTS_SEL")) = False Then sTmp = .Fields("PTS_SEL")
                    uSTD(nRec).PTS_SEL = sTmp
                    
                
                sTmp = "":      If IsNull(.Fields("ORD_NO")) = False Then sTmp = .Fields("ORD_NO")
                    uSTD(nRec).ORD_NO = sTmp
                    
                sTmp = "":      If IsNull(.Fields("IMAGE_FILE")) = False Then sTmp = .Fields("IMAGE_FILE")
                    uSTD(nRec).IMAGE_FILE = sTmp
                    
                sTmp = "":      If IsNull(.Fields("WANT_ACID")) = False Then sTmp = .Fields("WANT_ACID")
                    uSTD(nRec).WANT_ACID = sTmp
                
                If uSTD(nRec).ORD_NO = "" Then          '< �п��������� ��� : ���� ���ε�
                    sFilePath = ""
                    Select Case Trim(basModule.SchCD)
                        Case "N"
                            sFilePath = "NDOC/dshw/noryangjin/register/ACC/"
                        Case "K", "W", "Q"
                            sFilePath = "NDOC/dshw/kangnam/register/ACC/"
                        Case "S"
                            sFilePath = "NDOC/dshw/songpa/register/ACC/"
                        Case "P"
                            sFilePath = "NDOC/dshw/msongpa/register/ACC/"
                        Case "M"
                            sFilePath = "NDOC/dshw/mkangnam/register/ACC/"
                        Case "J"
                            sFilePath = "NDOC/dshw/mgwanghaw/register/ACC/"
                        Case "B"
                            sFilePath = "NDOC/dshw/busan/register/ACC/"
                    End Select
                    
                    sFilePath = sFilePath & Trim(basModule.SchCD) & uSTD(nRec).EXMID & ".jpg"       '< image ��� : ORDNO �� ���� ���
                
                    uSTD(nRec).IMAGE_DIR = sFilePath
                Else
                    sTmp = "":      If IsNull(.Fields("IMAGE_DIR")) = False Then sTmp = .Fields("IMAGE_DIR")
                    uSTD(nRec).IMAGE_DIR = sTmp
                End If
                
                sTmp = "":      If IsNull(.Fields("GR")) = False Then sTmp = .Fields("GR")
                    uSTD(nRec).GR = sTmp
                
                
                
                
                
                '-------------------------------------------------------------------------------
                uSTD(nRec).JK_NUM = 0
                uSTD(nRec).KK_NUM = 0
                
                uSTD(nRec).JM_NUM = 0
                uSTD(nRec).KM_NUM = 0
                
                uSTD(nRec).JE_NUM = 0
                uSTD(nRec).KE_NUM = 0
                
                uSTD(nRec).JTOT_NUM = 0
                uSTD(nRec).KTOT_NUM = 0
                    
                Select Case Trim(basModule.SchCD)
                    Case "K", "W", "Q", "M"
                        Select Case uSTD(nRec).KAEYOL
                            Case "01", "04", "06", "11", "16"
                                If uSTD(nRec).PTS_SEL = "1" Then '��������
                                
                                    '// ���
                                    nTmp = 0:      If IsNumeric(.Fields("J01")) = True Then nTmp = .Fields("J01")
                                        uSTD(nRec).JK_NUM = nTmp
                                    nTmp = 0:      If IsNumeric(.Fields("K01")) = True Then nTmp = .Fields("K01")
                                        uSTD(nRec).KK_NUM = nTmp
                                    
                                    '// ���� (����)
                                    nTmp = 0:      If IsNumeric(.Fields("J02")) = True Then nTmp = .Fields("J02")
                                        uSTD(nRec).JM_NUM = nTmp
                                    nTmp = 0:      If IsNumeric(.Fields("K02")) = True Then nTmp = .Fields("K02")
                                        uSTD(nRec).KM_NUM = nTmp
                                     
                                    '// �ܱ���
                                    nTmp = 0:      If IsNumeric(.Fields("J03")) = True Then nTmp = .Fields("J03")
                                        uSTD(nRec).JE_NUM = nTmp
                                    nTmp = 0:      If IsNumeric(.Fields("K03")) = True Then nTmp = .Fields("K03")
                                        uSTD(nRec).KE_NUM = nTmp
                                    
                                ElseIf uSTD(nRec).PTS_SEL = "2" Then '��������
                                
                                    '// ���
                                    nTmp = 0:      If IsNumeric(.Fields("J01")) = True Then nTmp = .Fields("J01")
                                        uSTD(nRec).JK_NUM = nTmp
                                    nTmp = 0:      If IsNumeric(.Fields("K01")) = True Then nTmp = .Fields("K01")
                                        uSTD(nRec).KK_NUM = nTmp
                                    
                                    '// ���� (����)
                                    nTmp = 0:      If IsNumeric(.Fields("J18")) = True Then nTmp = .Fields("J18")
                                        uSTD(nRec).JM_NUM = nTmp
                                    nTmp = 0:      If IsNumeric(.Fields("K18")) = True Then nTmp = .Fields("K18")
                                        uSTD(nRec).KM_NUM = nTmp
                                     
                                    '// �ܱ���
                                    nTmp = 0:      If IsNumeric(.Fields("J03")) = True Then nTmp = .Fields("J03")
                                        uSTD(nRec).JE_NUM = nTmp
                                    nTmp = 0:      If IsNumeric(.Fields("K03")) = True Then nTmp = .Fields("K03")
                                        uSTD(nRec).KE_NUM = nTmp
                                    
                                Else
                                
                                    '// ���
                                    nTmp = 0:      If IsNumeric(.Fields("J01")) = True Then nTmp = .Fields("J01")
                                        uSTD(nRec).JK_NUM = nTmp
                                    nTmp = 0:      If IsNumeric(.Fields("K01")) = True Then nTmp = .Fields("K01")
                                        uSTD(nRec).KK_NUM = nTmp
                                        
'                                    2010.12.14 ���ѿ� CLSTD01TB�󿡼� SEL ���� ������ ����
'                                    If uSTD(nRec).SEL4 > " " Then
'                                        '// ���� (����)
'                                        nTmp = 0:      If IsNumeric(.Fields("J02")) = True Then nTmp = .Fields("J02")
'                                            uSTD(nRec).JM_NUM = nTmp
'                                        nTmp = 0:      If IsNumeric(.Fields("K02")) = True Then nTmp = .Fields("K02")
'                                            uSTD(nRec).KM_NUM = nTmp
'                                    Else
'                                        '// ���� (����)
'                                        nTmp = 0:      If IsNumeric(.Fields("J18")) = True Then nTmp = .Fields("J18")
'                                            uSTD(nRec).JM_NUM = nTmp
'                                        nTmp = 0:      If IsNumeric(.Fields("K18")) = True Then nTmp = .Fields("K18")
'                                            uSTD(nRec).KM_NUM = nTmp
'                                    End If
                                    
                                    '����
                                    If uSTD(nRec).SEL4 = "84" Then
                                        '// ���� (����)
                                        nTmp = 0:      If IsNumeric(.Fields("J18")) = True Then nTmp = .Fields("J18")
                                            uSTD(nRec).JM_NUM = nTmp
                                        nTmp = 0:      If IsNumeric(.Fields("K18")) = True Then nTmp = .Fields("K18")
                                            uSTD(nRec).KM_NUM = nTmp
                                    Else
                                        '// ���� (����)
                                        nTmp = 0:      If IsNumeric(.Fields("J02")) = True Then nTmp = .Fields("J02")
                                            uSTD(nRec).JM_NUM = nTmp
                                        nTmp = 0:      If IsNumeric(.Fields("K02")) = True Then nTmp = .Fields("K02")
                                            uSTD(nRec).KM_NUM = nTmp
                                        
                                    End If
                                    
                                    '// �ܱ���
                                    nTmp = 0:      If IsNumeric(.Fields("J03")) = True Then nTmp = .Fields("J03")
                                        uSTD(nRec).JE_NUM = nTmp
                                    nTmp = 0:      If IsNumeric(.Fields("K03")) = True Then nTmp = .Fields("K03")
                                        uSTD(nRec).KE_NUM = nTmp
                                    
                                End If
                            Case "02", "05", "07", "12", "17"
                                If uSTD(nRec).PTS_SEL = "2" Then
                                
                                    '// ���
                                    nTmp = 0:      If IsNumeric(.Fields("J01")) = True Then nTmp = .Fields("J01")
                                        uSTD(nRec).JK_NUM = nTmp
                                    nTmp = 0:      If IsNumeric(.Fields("K01")) = True Then nTmp = .Fields("K01")
                                        uSTD(nRec).KK_NUM = nTmp
                                    
                                    '// ���� (����)
                                    nTmp = 0:      If IsNumeric(.Fields("J18")) = True Then nTmp = .Fields("J18")
                                        uSTD(nRec).JM_NUM = nTmp
                                    nTmp = 0:      If IsNumeric(.Fields("K18")) = True Then nTmp = .Fields("K18")
                                        uSTD(nRec).KM_NUM = nTmp
                                     
                                    '// �ܱ���
                                    nTmp = 0:      If IsNumeric(.Fields("J03")) = True Then nTmp = .Fields("J03")
                                        uSTD(nRec).JE_NUM = nTmp
                                    nTmp = 0:      If IsNumeric(.Fields("K03")) = True Then nTmp = .Fields("K03")
                                        uSTD(nRec).KE_NUM = nTmp
                                
                                ElseIf uSTD(nRec).PTS_SEL = "1" Then
                                
                                    '// ���
                                    nTmp = 0:      If IsNumeric(.Fields("J01")) = True Then nTmp = .Fields("J01")
                                        uSTD(nRec).JK_NUM = nTmp
                                    nTmp = 0:      If IsNumeric(.Fields("K01")) = True Then nTmp = .Fields("K01")
                                        uSTD(nRec).KK_NUM = nTmp
                                        
'                                    2010.12.14 ���ѿ� CLSTD01TB�󿡼� SEL ���� ������ ����
'                                    If uSTD(nRec).SEL4 > " " Then
'                                        '// ���� (����)
'                                        nTmp = 0:      If IsNumeric(.Fields("J02")) = True Then nTmp = .Fields("J02")
'                                            uSTD(nRec).JM_NUM = nTmp
'                                        nTmp = 0:      If IsNumeric(.Fields("K02")) = True Then nTmp = .Fields("K02")
'                                            uSTD(nRec).KM_NUM = nTmp
'                                    Else
'                                        '// ���� (����)
'                                       nTmp = 0:      If IsNumeric(.Fields("J18")) = True Then nTmp = .Fields("J18")
'                                            uSTD(nRec).JM_NUM = nTmp
'                                       nTmp = 0:      If IsNumeric(.Fields("K18")) = True Then nTmp = .Fields("K18")
'                                           uSTD(nRec).KM_NUM = nTmp
'                                    End If
                                    
                                    '// ���� (����)
                                    nTmp = 0:      If IsNumeric(.Fields("J02")) = True Then nTmp = .Fields("J02")
                                        uSTD(nRec).JM_NUM = nTmp
                                    nTmp = 0:      If IsNumeric(.Fields("K02")) = True Then nTmp = .Fields("K02")
                                        uSTD(nRec).KM_NUM = nTmp
                                    
                                    '// �ܱ���
                                    nTmp = 0:      If IsNumeric(.Fields("J03")) = True Then nTmp = .Fields("J03")
                                        uSTD(nRec).JE_NUM = nTmp
                                    nTmp = 0:      If IsNumeric(.Fields("K03")) = True Then nTmp = .Fields("K03")
                                        uSTD(nRec).KE_NUM = nTmp
                                    
                                Else
                                
                                    '// ���
                                    nTmp = 0:      If IsNumeric(.Fields("J01")) = True Then nTmp = .Fields("J01")
                                        uSTD(nRec).JK_NUM = nTmp
                                    nTmp = 0:      If IsNumeric(.Fields("K01")) = True Then nTmp = .Fields("K01")
                                        uSTD(nRec).KK_NUM = nTmp
                                        
'                                    2010.12.14 ���ѿ� CLSTD01TB�󿡼� SEL ���� ������ ����
'                                    If uSTD(nRec).SEL4 > " " Then
'                                        '// ���� (����)
'                                        nTmp = 0:      If IsNumeric(.Fields("J02")) = True Then nTmp = .Fields("J02")
'                                            uSTD(nRec).JM_NUM = nTmp
'                                        nTmp = 0:      If IsNumeric(.Fields("K02")) = True Then nTmp = .Fields("K02")
'                                            uSTD(nRec).KM_NUM = nTmp
'                                    Else
'                                        '// ���� (����)
'                                        nTmp = 0:      If IsNumeric(.Fields("J18")) = True Then nTmp = .Fields("J18")
'                                            uSTD(nRec).JM_NUM = nTmp
'                                        nTmp = 0:      If IsNumeric(.Fields("K18")) = True Then nTmp = .Fields("K18")
'                                            uSTD(nRec).KM_NUM = nTmp
'                                    End If
                                    
                                    '����
                                    If uSTD(nRec).SEL4 = "84" Then
                                        '// ���� (����)
                                        nTmp = 0:      If IsNumeric(.Fields("J18")) = True Then nTmp = .Fields("J18")
                                            uSTD(nRec).JM_NUM = nTmp
                                        nTmp = 0:      If IsNumeric(.Fields("K18")) = True Then nTmp = .Fields("K18")
                                            uSTD(nRec).KM_NUM = nTmp
                                    Else
                                        '// ���� (����)
                                        nTmp = 0:      If IsNumeric(.Fields("J02")) = True Then nTmp = .Fields("J02")
                                            uSTD(nRec).JM_NUM = nTmp
                                        nTmp = 0:      If IsNumeric(.Fields("K02")) = True Then nTmp = .Fields("K02")
                                            uSTD(nRec).KM_NUM = nTmp
                                        
                                    End If
                                    
                                    '// �ܱ���
                                    nTmp = 0:      If IsNumeric(.Fields("J03")) = True Then nTmp = .Fields("J03")
                                        uSTD(nRec).JE_NUM = nTmp
                                    nTmp = 0:      If IsNumeric(.Fields("K03")) = True Then nTmp = .Fields("K03")
                                        uSTD(nRec).KE_NUM = nTmp
                                    
                                End If
                                
                        End Select
                End Select
                '-------------------------------------------------------------------------------
                
                
                
                
                .MoveNext
                
            Next nRec
            
            Call Get_STD_image              '<< �̹��� �ڷ� ��������
            
            Call Std_Data_Show(1)           '<< �л��ڷ� ȭ�� ���̱�
            Me.Tag = "LOAD"
                VScroll1.value = 1
                
                txtPage.Text = "1/" & Trim(CStr(nTotRec))
            Me.Tag = ""
            
        End If
    End With

    MsgBox "�л� ��ȸ�Ͽ����ϴ�.", vbInformation + vbOKOnly, "�л���ȸ"
    
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    VScroll1.Enabled = True
    
    Exit Sub
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    On Error GoTo 0
    MsgBox "�л���ȸ�� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "�л���ȸ"
End Sub


'>> scroll �̵�
Private Sub VScroll1_Change()
    If Me.Tag = "LOAD" Then Exit Sub
    
    VScroll1.Enabled = False
        Call Std_Data_Show(VScroll1.value)
        txtPage.Text = Trim(CStr(VScroll1.value)) & "/" & Trim(CStr(nTotRec))
    VScroll1.Enabled = True
End Sub

Private Sub lvRecvList_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If Me.Tag = "LOAD" Then Exit Sub
    
    Call Std_Data_Show(Item.Index)
    
End Sub

Private Sub Std_Data_Show(Index As Long)
    
    Dim nKME_Sum        As Integer
    
    If Me.Tag = "LOAD" Then Exit Sub
    
    If UBound(uSTD) < 1 Then Exit Sub
    If UBound(uSTD) < Index Then Exit Sub
    
    With uSTD(Index)
        
        Select Case Trim(.KAEYOL)   '<< �迭: 01,02,03-�ι�,�ڿ�,��ü   06,05-�����ι�,�ڿ�  06,07 -��������,�Ǵ�
            Case "01":  �����迭.Text = "�� �� ��"
                        �����迭2.Text = "��    ��"
            Case "02":  �����迭.Text = "�� �� ��"
                        �����迭2.Text = "��    ��"
            Case "03":  Select Case Trim(basModule.SchCD)
                               Case "N"
                                    �����迭.Text = "��.ü�ɰ�"
                                    �����迭2.Text = "��.ü�ɰ�"
                               Case "S", "P"                        '< 2008.02.15 : ����/ ����
                                    �����迭.Text = "��.ü�ɰ�"
                                    �����迭2.Text = "��.ü�ɰ�"
                               Case Else
                                    �����迭.Text = ""
                                    �����迭2.Text = ""
                        End Select
            Case "04":  Select Case Trim(basModule.SchCD)
                               Case "N"
                                    �����迭.Text = "����(��) �ڿ�"
                                    �����迭2.Text = "�� �� ��"
                               Case "K", "W"
                                    �����迭.Text = "�ָ�������"
                                    �����迭2.Text = "�� ��"
                               Case "S", "P"                       '< 2008.02.15 : ����/ ����
                                    �����迭.Text = "Ư���ڿ�"
                                    �����迭2.Text = "Ư���ڿ�"
                               Case Else
                                    �����迭.Text = ""
                                    �����迭2.Text = ""
                        End Select
            Case "05":  Select Case Trim(basModule.SchCD)
                               Case "N"
                                    �����迭.Text = "���� �ι�"
                                    �����迭2.Text = "��������"
                               Case "K", "W"
                                    �����迭.Text = "�ָ�������"
                                    �����迭2.Text = "�� ��"
                               Case "S"
                                    �����迭.Text = "�����ι�"
                                    �����迭2.Text = "�����ι�"
                               Case "B"
                                    �����迭.Text = "��ȭ�ι�"
                                    �����迭2.Text = "��ȭ�ι�"
                               Case Else
                                    �����迭.Text = ""
                                    �����迭2.Text = ""
                        End Select
            Case "06":  Select Case Trim(basModule.SchCD)
                               Case "N"
                                    �����迭.Text = "���� �ڿ�"
                                    �����迭2.Text = "��������"
                               Case "K"
                                    �����迭.Text = "�߰�������"
                                    �����迭2.Text = "�� ��"
                               Case "Q"
                                    �����迭.Text = "�߰������"
                                    �����迭2.Text = "�� ��"
                               Case "S"
                                    �����迭.Text = "�����ڿ�"
                                    �����迭2.Text = "�����ڿ�"
                               Case "B"
                                    �����迭.Text = "��ȭ�ڿ�"
                                    �����迭2.Text = "��ȭ�ڿ�"
                               Case Else
                                    �����迭.Text = ""
                                    �����迭2.Text = ""
                        End Select
            Case "07":  Select Case Trim(basModule.SchCD)
                               Case "K"
                                    �����迭.Text = "�߰�������"
                                    �����迭2.Text = "�� ��"
                               Case "Q"
                                    �����迭.Text = "�߰������"
                                    �����迭2.Text = "�� ��"
                               Case "N"
                                    �����迭.Text = "�ż��ι�"
                                    �����迭2.Text = "�ż��ι�"
                                Case "B"
                                    �����迭.Text = "�������ι�"
                                    �����迭2.Text = "�������ι�"
                               Case Else: �����迭.Text = ""
                                          �����迭2.Text = ""
                        End Select
            Case "08":  Select Case Trim(basModule.SchCD)
                               Case "N":  �����迭.Text = "�ż��ڿ�"
                                          �����迭2.Text = "�ż��ڿ�"
                               Case "S":  �����迭.Text = "��������"
                                          �����迭2.Text = "��������"
                               Case "B"
                                    �����迭.Text = "�������ڿ�"
                                    �����迭2.Text = "�������ڿ�"
                               Case Else: �����迭.Text = ""
                                          �����迭2.Text = ""
                        End Select
                        
            Case "09":  Select Case Trim(basModule.SchCD)
                               Case "N":  �����迭.Text = "�ż�����"
                                          �����迭2.Text = "��  ��"
                               Case "B"
                                    �����迭.Text = "��ȭ�ι�"
                                    �����迭2.Text = "��ȭ�ι�"
                               Case Else: �����迭.Text = ""
                                          �����迭2.Text = ""
                        End Select
            Case "10":  Select Case Trim(basModule.SchCD)
                               Case "N":  �����迭.Text = "�ż�����"
                                          �����迭2.Text = "��  ��"
                               Case "B"
                                    �����迭.Text = "��ȭ�ڿ�"
                                    �����迭2.Text = "��ȭ�ڿ�"
                               Case Else: �����迭.Text = ""
                                          �����迭2.Text = ""
                        End Select
                        
                        
            Case "11", "16":  Select Case Trim(basModule.SchCD)
                                     Case "N":  �����迭.Text = "��)�ι�"
                                                �����迭2.Text = "��    ��"
                                     Case "K":  �����迭.Text = "�������ι�"
                                                �����迭2.Text = "��    ��"
                                     Case "W":  �����迭.Text = "�������ι�"
                                                �����迭2.Text = "��    ��"
                                     Case "Q":  �����迭.Text = "�߰������"
                                                �����迭2.Text = "�����ι�"
                                     Case "S":  �����迭.Text = "�ż��ι�"
                                                �����迭2.Text = "�ż��ι�"
                                     Case "J":  �����迭.Text = "�ż��ι�"
                                                �����迭2.Text = "�ż��ι�"
                                     Case Else: �����迭.Text = ""
                                                �����迭2.Text = ""
                              End Select
            Case "12", "17":  Select Case Trim(basModule.SchCD)
                                     Case "N":  �����迭.Text = "��)�ڿ�"
                                                �����迭2.Text = "��    ��"
                                     Case "K":  �����迭.Text = "�������ڿ�"
                                                �����迭2.Text = "��    ��"
                                     Case "W":  �����迭.Text = "�������ڿ�"
                                                �����迭2.Text = "��    ��"
                                     Case "Q":  �����迭.Text = "�߰������"
                                                �����迭2.Text = "�����ڿ�"
                                     Case "S":  �����迭.Text = "�ż��ڿ�"
                                                �����迭2.Text = "�ż��ڿ�"
                                     Case "J":  �����迭.Text = "�ż��ڿ�"
                                                �����迭2.Text = "�ż��ڿ�"
                                     Case Else: �����迭.Text = ""
                                                �����迭2.Text = ""
                              End Select
                              
            Case "13":        Select Case Trim(basModule.SchCD)
                                     Case "N":  �����迭.Text = "��)��ü��"
                                                �����迭2.Text = "��ü��"
                              End Select
            Case "14":        Select Case Trim(basModule.SchCD)
                                     Case "N":  �����迭.Text = "��)����(��)"
                                                �����迭2.Text = "��    ��"
                              End Select
            Case "15":        Select Case Trim(basModule.SchCD)
                                     Case "N":  �����迭.Text = "��)�ι�����"
                                                �����迭2.Text = "��    ��"
                              End Select
            Case "16":        Select Case Trim(basModule.SchCD)
                                     Case "N":  �����迭.Text = "��)�ڿ�����"
                                                �����迭2.Text = "��    ��"
                              End Select
                         
            Case "18":        Select Case Trim(basModule.SchCD)
                                     Case "S":  �����迭.Text = "�ι������̾�"
                                                �����迭2.Text = "�ι������̾�"
                                     Case "J":  �����迭.Text = "�ι������̾�"
                                                �����迭2.Text = "�ι������̾�"
                              End Select
                              
            Case "19":        Select Case Trim(basModule.SchCD)
                                     Case "S":  �����迭.Text = "�ڿ������̾�"
                                                �����迭2.Text = "�ڿ������̾�"
                                     Case "J":  �����迭.Text = "�ڿ������̾�"
                                                �����迭2.Text = "�ڿ������̾�"
                              End Select
                              
            Case Else:  �����迭.Text = ""
        End Select
        
        
        
        �����ȣ.Text = .EXMID
        �л�����.Text = .STDNM:                 ���.Text = .GR
        ����.Text = .SEX
        �������.Text = .Birth_ymd
        �л�������ȣ.Text = "(" & .ZIP & ")"
        �л��ּ�1.Text = .ADDR1
        �л��ּ�2.Text = .ADDR2
        �л���Ű�.Text = .HIGH_SCH
        �����⵵.Text = .GRADE_YEAR
        �л��̸���.Text = .EMAIL
        �л�����ó_��.Text = .TEL
        �л�����ó_�޴���.Text = .CEL
        
        
        ��ȣ�ڼ���.Text = .PRNT_NM
        ��ȣ�ڰ���.Text = .PRNT_RLTN
        
        ��ȣ�ڿ���ó.Text = .PRNT_TEL
        ��ȣ�ڿ���ó_�޴���.Text = .PRNT_CEL
        ��ȣ�ڿ�����ȣ.Text = "(" & .PRNT_ZIP & ")"
        ��ȣ���ּ�1.Text = .PRNT_ADDR1
        ��ȣ���ּ�2.Text = .PRNT_ADDR2
        
        ��ȣ������.Text = .PRNT_JOB
        ��ȣ�ڿ���ó_����.Text = .PRNT_W_TEL
                             
        ����_��ȸŽ��.Text = " "
        ����_�ܱ���.Text = " "
        ����_��������.Text = " "
        ����_����Ž��.Text = " "
        
        ����_��ȸŽ��.Text = Div_Gwamok_NM("SEL1", .SEL1)
        ����_�ܱ���.Text = Div_Gwamok_NM("SEL2", .SEL2)
        
        Select Case Trim(basModule.SchCD)
            Case "W":
                       ����_����Ž��.Text = ""
                       ����_��������.Text = ""
            
            Case "Q":
                       ����_����Ž��.Text = ""
                       ����_��������.Text = ""
            Case Else
                       ����_����Ž��.Text = Div_Gwamok_NM("SEL3", .SEL3)
                       ����_��������.Text = Div_Gwamok_NM("SEL4", .SEL4)
            End Select
            
        ����_��ȸ����.Text = ""
        ����_�ڿ�����.Text = ""
        If Trim(.SEL1) > " " Then                                   '<<- ��ȸ����
            ����_��ȸ����.Text = Div_Gwamok_NM("SEL5", .SEL5)
        ElseIf Trim(.SEL3) > " " Then                               '<<- �ڿ�����
            ����_�ڿ�����.Text = Div_Gwamok_NM("SEL5", .SEL5)
        End If
        
        ���.Text = ""
        ����.Text = ""
        ����.Text = ""
        �������.Text = ""
        
        ���2.Text = ""
        ����2.Text = ""
        ����2.Text = ""
        �������2.Text = ""
        
        
        
        ������_����.Text = ""
        ������_����.Text = ""
        ������_����.Text = ""
        
        Select Case Trim(.EXMTYPE)
            Case "0"
                
                nKME_Sum = 0
                If IsNumeric(Trim(.K_NUM)) = True Then nKME_Sum = nKME_Sum + CInt(.K_NUM)
                If IsNumeric(Trim(.M_NUM)) = True Then nKME_Sum = nKME_Sum + CInt(.M_NUM)
                If IsNumeric(Trim(.E_NUM)) = True Then nKME_Sum = nKME_Sum + CInt(.E_NUM)
                
                If nKME_Sum > 100 Then      'ǥ������
                    ���.Text = .K_NUM
                    ����.Text = .M_NUM
                    ����.Text = .E_NUM
                    
                    �������.Text = Format(nKME_Sum, "##0")
                Else                        '���
                    ���2.Text = .K_NUM
                    ����2.Text = .M_NUM
                    ����2.Text = .E_NUM
                    
                    �������2.Text = Format(nKME_Sum, "##0")
                End If
                
                
                Select Case Trim(basModule.SchCD)
                    Case "M"
                        
                        '## ǥ������ ������ �κ��� Ʋ����.
                        ���.Text = .JK_NUM
                        ����.Text = .JM_NUM
                        ����.Text = .JE_NUM
                        
                        �������.Text = .JK_NUM + .JM_NUM + .JE_NUM
                        
                        '> ��޸� ����
                        ���2.Text = .K_NUM
                        ����2.Text = .M_NUM
                        ����2.Text = .E_NUM
                        
                        �������2.Text = Format(nKME_Sum, "##0")
                    Case "K", "W", "Q"
                        '2010.12.13 ���ѿ� ��� ������ ����
                        '2011-01-10 ���ѿ� Ȳ���� ���� ��û���� ����� �� CLSTD03TB
                            ���2.Text = .KK_NUM
                            ����2.Text = .KM_NUM
                            ����2.Text = .KE_NUM
                End Select
                
            Case "1"
            
                Select Case Trim(basModule.SchCD)
                    Case "M"
                        ������_����.Text = .M_NUM
                        ������_����.Text = .E_NUM
                        ������_����.Text = .TOT_NUM
                        
                    Case "K"
                        ���.Text = .JK_NUM
                        ����.Text = .JM_NUM
                        ����.Text = .JE_NUM
                        
                        �������.Text = .JK_NUM + .JM_NUM + .JE_NUM
                        
                        ���2.Text = ""
                        ����2.Text = ""
                        ����2.Text = ""
                        
                        �������2.Text = ""
                        
                        If Format(Now, "yyyymmdd") > "20110202" Then
                            ������_����.Text = .M_NUM
                            ������_����.Text = .E_NUM
                            ������_����.Text = .TOT_NUM
                        Else
                            ������_����.Text = ""
                            ������_����.Text = ""
                            ������_����.Text = ""
                        End If
                        
                    Case "W"
                        ���.Text = .JK_NUM
                        ����.Text = .JM_NUM
                        ����.Text = .JE_NUM
                        
                        �������.Text = .JK_NUM + .JM_NUM + .JE_NUM
                        
                        ���2.Text = ""
                        ����2.Text = ""
                        ����2.Text = ""
                        
                        �������2.Text = ""
                        
                        If Format(Now, "yyyymmdd") > "20110220" Then
                            ������_����.Text = .M_NUM
                            ������_����.Text = .E_NUM
                            ������_����.Text = .TOT_NUM
                        Else
                            ������_����.Text = ""
                            ������_����.Text = ""
                            ������_����.Text = ""
                        End If
                        
                    Case "Q"
                    
                        ���.Text = .JK_NUM
                        ����.Text = .JM_NUM
                        ����.Text = .JE_NUM
                        
                        �������.Text = .JK_NUM + .JM_NUM + .JE_NUM
                        
                        ���2.Text = ""
                        ����2.Text = ""
                        ����2.Text = ""
                        
                        �������2.Text = ""
                        
                        If Format(Now, "yyyymmdd") > "20110211" Then
                            ������_����.Text = .M_NUM
                            ������_����.Text = .E_NUM
                            ������_����.Text = .TOT_NUM
                        Else
                            ������_����.Text = ""
                            ������_����.Text = ""
                            ������_����.Text = ""
                        End If
                        
                End Select
            
                
                
        End Select
        
        '>> �ι��� ����, �ڿ��� ����
        
        Select Case Trim(basModule.SchCD)
            Case "K", "W", "Q"
                Select Case Trim(.KAEYOL)
                    Case "01", "04", "06", "11", "16"
                        If Trim(.PTS_SEL) = "1" Then
                            ��������.Caption = "����[��]"
                        ElseIf Trim(.PTS_SEL) = "2" Then
                            ��������.Caption = "����[��]"
                        Else
                            ��������.Caption = IIf(Trim(.SEL4) > " ", "����[��]", "����[��]")
                        End If
                    Case "02", "05", "07", "12", "17"
                        If Trim(.PTS_SEL) = "2" Then
                            ��������.Caption = "����[��]"
                        ElseIf Trim(.PTS_SEL) = "1" Then
                            ��������.Caption = "����[��]"
                        Else
                            ��������.Caption = IIf(Trim(.SEL4) > " ", "����[��]", "����[��]")
                        End If
                    Case Else
                        ��������.Caption = ""
                End Select
            Case "S", "P"               '< 2009.01.12 : ����/ ����
                Select Case Trim(.KAEYOL)
                    Case "01", "03", "05", "18"
                        ��������.Caption = "����[��]"
                    Case "02", "04", "06", "19"
                        ��������.Caption = "����[��]"
                    Case Else
                        ��������.Caption = ""
                End Select
            Case Else
                Select Case Trim(.KAEYOL)
                    Case "01", "02", "04", "05", "06", "07", "08", "09", "10", "11", "12", "14", "15", "16"
                        '2011-01-10 ���ѿ� ���� ���� PTS_SEL 1:���� 2:����
                        '��������.Caption = IIf(Trim(.SEL4) > " ", "����[��]", "����[��]")
                        If Trim(.PTS_SEL) = "1" Then
                            ��������.Caption = "����[��]"
                        ElseIf Trim(.PTS_SEL) = "2" Then
                           ��������.Caption = "����[��]"
                        End If
'                    Case "04"
'                        ��������.Caption = "����[��]"
                    Case "03", "13"
                        ��������.Caption = "����"                                   '<< ��ü��
                    Case Else
                        ��������.Caption = ""
                End Select
        End Select
        
        
        �л�����.Tag = .SCHNO
        �����ȣ.Tag = .ORD_NO
        �п�����.Text = .R_WAY
        �����п�.Text = .WANT_ACID
        
        Set Photo.Picture = CheckJPG(sSavePath & "\" & .IMAGE_FILE & ".jpg")
        
    End With
    
End Sub


Private Function Div_Gwamok_NM(ByVal aGbn As String, ByVal aGwamok As String) As String
    Dim sDiv()      As String
    Dim ni          As Integer
    Dim sTmp        As String
    
    Dim sGwamok     As String
    
    sDiv = Split(aGwamok, "|", -1, vbTextCompare)
    
    sTmp = "":  sGwamok = ""
    For ni = 0 To UBound(sDiv) - 1 Step 1
        
        sTmp = sDiv(ni)
        
        Select Case aGbn
            Case "SEL1"
                Select Case sTmp
                    Case constSatamCodes(0)
                        sTmp = constSatams(0)
                    Case constSatamCodes(1)
                        sTmp = constSatams(1)
                    Case constSatamCodes(2)
                        sTmp = constSatams(2)
                    Case constSatamCodes(3)
                        sTmp = constSatams(3)
                    Case constSatamCodes(4)
                        sTmp = constSatams(4)
                    Case constSatamCodes(5)
                        sTmp = constSatams(5)
                    Case constSatamCodes(6)
                        sTmp = constSatams(6)
                    Case constSatamCodes(7)
                        sTmp = constSatams(7)
                    Case constSatamCodes(8)
                        sTmp = constSatams(8)
                    Case constSatamCodes(9)
                        sTmp = constSatams(9)
'                    Case "11"
'                        sTmp = "��������"
                End Select
            Case "SEL2"
                Select Case sTmp
                    Case "31"
                        sTmp = "����"
                    Case "32"
                        sTmp = "�Ͼ�"
                    Case "33"
                        sTmp = "�����ĳľ�"
                    Case "34"
                        sTmp = "�Ҿ�"
                    Case "35"
                        sTmp = "�߱���"
                    Case "36"
                        sTmp = "�ѹ�"
                    
                    Case "37"
                        sTmp = "���"
                    Case "38"
                        sTmp = "����"
                    Case "39"
                        sTmp = "����"
                    Case "40"
                        sTmp = "�����"
                    Case "41"
                        sTmp = "��������"
                    Case "42"
                        sTmp = "�ƶ���"
                End Select
            Case "SEL3"
                Select Case sTmp
                    Case "51"
                        sTmp = "����1"
                    Case "52"
                        sTmp = "ȭ��1"
                    Case "53"
                        sTmp = "��������1"
                    Case "54"
                        sTmp = "��������1"
                    Case "55"
                        sTmp = "����2"
                    Case "56"
                        sTmp = "ȭ��2"
                    Case "57"
                        sTmp = "��������2"
                    Case "58"
                        sTmp = "��������2"
                End Select
            Case "SEL4"
                Select Case sTmp
                    Case "81"
                        sTmp = "������"
                    Case "82"
                        sTmp = "�̻����"
                    Case "83"
                        sTmp = "Ȯ�����"
                    Case "84"
                        sTmp = "��������"
                End Select
            Case "SEL5"
                Select Case sTmp
                    Case "91"
                        sTmp = "���"
                    Case "92"
                        sTmp = "����"
                    Case "93"
                        sTmp = "�ܱ���"         '< ����
                    Case "94"
                        sTmp = ""               '< ����
                End Select
            Case Else
                sTmp = ""
        End Select
        
        If ni > 0 Then sGwamok = sGwamok & ", "
        sGwamok = sGwamok & sTmp
        
    Next ni
    
    If sGwamok = "" Then
        Div_Gwamok_NM = ""
    Else
        Div_Gwamok_NM = sGwamok
    End If
    
End Function

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
    If (header(0) > 0) Then
    'If (header(0) = 255 And header(1) = 216) And _
       (tailer(0) = 255 And tailer(1) >= 209) Then
        Set CheckJPG = LoadPicture(fileName)
    Else
        Set CheckJPG = imgList.ListImages.Item(1).Picture       '<< no-image
    End If

End Function


'## ������ �̹��� ��������
Private Sub Get_STD_image()
    
    Dim bData()     As Byte
    Dim f           As Integer
    Dim nRec        As Long

    Dim sLocalFile  As String
    Dim sSourceUrl  As String
    
    Dim MaxSize     As Long

    On Error Resume Next

    f = FreeFile()
    
    For nRec = 1 To UBound(uSTD) Step 1
    
        sLocalFile = sSavePath & "\" & uSTD(nRec).IMAGE_FILE & ".jpg"       '<< unique key : �п�+�����ȣ
        
        If Dir(sLocalFile) > " " Then
            Open sLocalFile For Binary As #f
                On Error Resume Next
                MaxSize = LOF(f)
            Close f
            
            If MaxSize = 0 Then
                Kill sLocalFile
            End If
        End If
        
        If Dir(sLocalFile) = "" Then                                        '<< �л� �̹��� ���� �͸� ����
        
            Select Case Trim(basModule.SchCD)
                Case "B"        '<< �λ�����
                    sSourceUrl = "http://www.dsbusan.com" & uSTD(nRec).PHOTO_PATH           '<< ������ �̹��� ���
                Case Else
                    sSourceUrl = "http://www.dshw.co.kr" & uSTD(nRec).PHOTO_PATH            '<< ������ �̹��� ���
            End Select
            
            bData() = Inet1.OpenURL(sSourceUrl, icByteArray)
            
            If UBound(bData) > 0 Then
                Open sLocalFile For Binary Access Write As #f
                Put #f, , bData()
            
                DoEvents
                Close #f
            End If
        End If
        
    Next nRec
    
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
        
        
        Call Std_Data_Show(nRec)                                '<< �л��ڷ� ȭ�� ���̱�
        Me.Tag = "LOAD"
            VScroll1.value = nRec
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


'## ���� ���ε�
Private Sub Photo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sFileLocation   As String
    Dim sSchNO          As String
    Dim sOrdNO          As String
    Dim sExmID          As String
    Dim simageFile      As String

    Dim bRet            As String
    
    Dim sDiv()          As String
    Dim nS              As Long
    Dim sLocalFile      As String
    
    
    If Button <> vbRightButton Then
        Exit Sub
    End If

    If �л�����.Tag = "" Then
        MsgBox "�л��� ��ȸ�Ͻʽÿ�.", vbExclamation + vbOKOnly, "���� ���ε�"
        Exit Sub
    End If
    If UBound(uSTD) < 1 Then
        MsgBox "�л��� ��ȸ�Ͻʽÿ�.", vbExclamation + vbOKOnly, "���� ���ε�"
        Exit Sub
    End If
    
    '�����ȣ.tag
    
    With uSTD(VScroll1.value)
        sOrdNO = .ORD_NO
        sSchNO = .SCHNO
        sExmID = .EXMID
        sFileLocation = .IMAGE_DIR
        simageFile = .IMAGE_FILE
        
        bRet = ""
        If Trim(sOrdNO) = "" Then        '< �̹����� ���� ��쿣 ������ ����
            bRet = Make_image_Path(sSchNO, sExmID, simageFile)
            
            If bRet = "" Then
                MsgBox "��� ������ ������ �ֽ��ϴ�." & vbCrLf & _
                       "�����ڿ��� �����Ͻʽÿ�.", vbExclamation + vbOKOnly, "���� ���ε�"
                Exit Sub
            Else
                sFileLocation = bRet
            End If
        End If
    End With
    
    '<< ���� ����� >>
    If Trim(txtPage) > " " Then
        sDiv = Split(txtPage.Text, "/", -1, vbTextCompare)
        
        nS = CLng(sDiv(0))
        sLocalFile = sSavePath & "\" & uSTD(nS).IMAGE_FILE & ".jpg"       '<< unique key : �п�+�����ȣ
        If Dir(sLocalFile) > " " Then
            Kill sLocalFile
        End If
    End If
    
    '���� �ֱ�
    Load INT900
    Call INT900.Save_Photo(sFileLocation, sSchNO)
    INT900.Show
    
End Sub


'## �̹��� ���� ��� ������ ����
Private Function Make_image_Path(ByVal aSchNO As String, ByVal aExmID As String, ByVal aimageFile As String) As String
    Dim sFilePath       As String
    
    Dim sStr            As String
    Dim DBCmd           As ADODB.Command
    Dim DBParam         As ADODB.Parameter
    
    Dim ni              As Long
    Dim sLocalFile      As String
    Dim nExe            As Integer
    Dim f               As Integer
    Dim MaxSize         As Long
    
    sFilePath = ""
    Select Case Trim(basModule.SchCD)
        Case "N"
            sFilePath = "/NDOC/dshw/noryangjin/register/ACC/"
        Case "K", "W", "Q"
            sFilePath = "/NDOC/dshw/kangnam/register/ACC/"
        Case "S"
            sFilePath = "/NDOC/dshw/songpa/register/ACC/"
        Case "P"
            sFilePath = "/NDOC/dshw/msongpa/register/ACC/"
        Case "M"
            sFilePath = "/NDOC/dshw/mkangnam/register/ACC/"
        Case "J"
            sFilePath = "/NDOC/dshw/mgwanghwa/register/ACC/"
        Case "B"
            sFilePath = "/NDOC/dshw/busan/register/ACC/"
    End Select
    
    sFilePath = sFilePath & Trim(basModule.SchCD) & aExmID & ".jpg"
    
    On Error GoTo ErrStmt
    
    basDataBase.DBConn.BeginTrans
    
    Set DBCmd = New ADODB.Command
    Set DBParam = New ADODB.Parameter
    
    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
    
    
    
    '<< UPDATE
    sStr = ""
    sStr = sStr & " Update CLSTD01TB"
    sStr = sStr & "    SET PHOTO_PATH = '" & sFilePath & "'"
    sStr = sStr & "  WHERE SCHNO = '" & Trim(aSchNO) & "'"
            
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
        
        f = FreeFile()
        sLocalFile = sSavePath & "\" & aimageFile & ".jpg"               '<< unique key : �п�+�����ȣ
        If Dir(sLocalFile) > " " Then
            Open sLocalFile For Binary As #f
                On Error Resume Next
                MaxSize = LOF(f)
            Close f
            
            Kill sLocalFile
            
        End If
    
        Make_image_Path = sFilePath
    Else
        basDataBase.DBConn.RollbackTrans
        
        Set DBCmd = Nothing
        Set DBParam = Nothing
    
        Make_image_Path = ""
    End If
    
    Exit Function
    
ErrStmt:
    basDataBase.DBConn.RollbackTrans
    
    Set DBCmd = Nothing
    Set DBParam = Nothing
    
    Make_image_Path = ""
End Function










































