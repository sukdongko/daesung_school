VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form MAT010 
   Caption         =   "���л��� >> ���п��� ��� >> ���� ���� Ŭ����"
   ClientHeight    =   10740
   ClientLeft      =   4470
   ClientTop       =   2430
   ClientWidth     =   14130
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   10740
   ScaleWidth      =   14130
   Begin VB.Frame Frame2 
      BackColor       =   &H0082C8E8&
      BorderStyle     =   0  '����
      Caption         =   "Frame2"
      Height          =   495
      Left            =   30
      TabIndex        =   39
      Top             =   0
      Width           =   14085
      Begin VB.Frame Frame1 
         BackColor       =   &H00D2EAF5&
         BorderStyle     =   0  '����
         Caption         =   "Frame1"
         Height          =   435
         Left            =   30
         TabIndex        =   40
         Top             =   30
         Width           =   14025
         Begin VB.CommandButton cmdPrintAll 
            Caption         =   "��ü������ ���"
            Height          =   375
            Left            =   10740
            TabIndex        =   48
            Top             =   30
            Width           =   1515
         End
         Begin VB.CommandButton cmdPrint 
            Caption         =   "���������� ���"
            Height          =   375
            Left            =   9150
            TabIndex        =   47
            Top             =   30
            Width           =   1515
         End
         Begin VB.ComboBox cboKaeyol 
            Height          =   300
            Left            =   870
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   46
            Top             =   67
            Width           =   1155
         End
         Begin VB.CommandButton cmdFind 
            Caption         =   "�л� ��ȸ"
            Height          =   375
            Left            =   7200
            TabIndex        =   45
            Top             =   30
            Width           =   1515
         End
         Begin VB.TextBox txtStdNM 
            Height          =   285
            Left            =   2730
            TabIndex        =   44
            Text            =   "txtStdNM"
            Top             =   75
            Width           =   945
         End
         Begin VB.TextBox txtPage 
            Enabled         =   0   'False
            Height          =   375
            Left            =   12840
            TabIndex        =   43
            Text            =   "txtPage"
            Top             =   30
            Width           =   735
         End
         Begin VB.CommandButton cmdShiftLeft 
            Caption         =   "��"
            Height          =   375
            Left            =   12390
            TabIndex        =   42
            Top             =   30
            Width           =   405
         End
         Begin VB.CommandButton cmdShiftRight 
            Caption         =   "��"
            Height          =   375
            Left            =   13590
            TabIndex        =   41
            Top             =   30
            Width           =   405
         End
         Begin EditLib.fpMask fpExmID_S 
            Height          =   285
            Left            =   4650
            TabIndex        =   49
            Top             =   75
            Width           =   795
            _Version        =   196608
            _ExtentX        =   1402
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
            Left            =   5940
            TabIndex        =   50
            Top             =   75
            Width           =   795
            _Version        =   196608
            _ExtentX        =   1402
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
            Left            =   480
            TabIndex        =   53
            Top             =   120
            Width           =   945
         End
         Begin VB.Label NonPrintLbl 
            BackStyle       =   0  '����
            Caption         =   "�����ȣ          ����          ����"
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
            Left            =   3900
            TabIndex        =   52
            Top             =   120
            Width           =   3285
         End
         Begin VB.Label NonPrintLbl 
            BackStyle       =   0  '����
            Caption         =   "�л���"
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
            Left            =   2160
            TabIndex        =   51
            Top             =   120
            Width           =   945
         End
      End
   End
   Begin VB.PictureBox pReportControl 
      Height          =   9915
      Left            =   0
      ScaleHeight     =   9855
      ScaleWidth      =   14040
      TabIndex        =   0
      Top             =   540
      Width           =   14100
      Begin VB.TextBox �����ȣ 
         Alignment       =   1  '������ ����
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
         Left            =   11400
         TabIndex        =   59
         Text            =   "N12501"
         Top             =   750
         Width           =   1050
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   9885
         Left            =   13800
         TabIndex        =   38
         Top             =   0
         Width           =   255
      End
      Begin VB.PictureBox pReportViewer 
         Appearance      =   0  '���
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   9870
         Left            =   0
         ScaleHeight     =   9840
         ScaleWidth      =   13770
         TabIndex        =   1
         Top             =   0
         Width           =   13800
         Begin VB.TextBox ��ȣ���ּ�1 
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
            Left            =   3300
            TabIndex        =   70
            Text            =   "��ȣ���ּ�1"
            Top             =   6630
            Width           =   3135
         End
         Begin VB.TextBox �����迭 
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
            Index           =   2
            Left            =   1335
            TabIndex        =   66
            Text            =   "�����迭2"
            Top             =   2100
            Width           =   1515
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
            Index           =   1
            Left            =   11160
            TabIndex        =   60
            Text            =   "��.ü�ɰ�"
            Top             =   330
            Width           =   1530
         End
         Begin VB.TextBox ������� 
            BorderStyle     =   0  '����
            Height          =   225
            Left            =   7710
            TabIndex        =   58
            Text            =   "�������"
            Top             =   3570
            Width           =   2895
         End
         Begin VB.TextBox �г�1 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  '����
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3840
            TabIndex        =   55
            Text            =   "�г�1"
            Top             =   7830
            Width           =   1545
         End
         Begin VB.TextBox ��� 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  '����
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   870
            TabIndex        =   54
            Text            =   "6�� �򰡿�"
            Top             =   8790
            Width           =   1395
         End
         Begin VB.TextBox �г� 
            Appearance      =   0  '���
            BorderStyle     =   0  '����
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2460
            TabIndex        =   19
            Text            =   "�г�"
            Top             =   2970
            Width           =   660
         End
         Begin VB.TextBox �л����� 
            Appearance      =   0  '���
            BorderStyle     =   0  '����
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2460
            TabIndex        =   18
            Text            =   "ȫ�浿"
            Top             =   3570
            Width           =   1545
         End
         Begin VB.TextBox �����迭 
            Appearance      =   0  '���
            BorderStyle     =   0  '����
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   7710
            TabIndex        =   17
            Text            =   "�迭"
            Top             =   2940
            Width           =   1860
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
            Left            =   7710
            TabIndex        =   16
            Text            =   "iiiboss_12345@mail.naver.com"
            Top             =   5460
            Width           =   4095
         End
         Begin VB.TextBox �л������ȣ 
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
            Height          =   165
            Left            =   2460
            TabIndex        =   15
            Text            =   "(100-100)"
            Top             =   4110
            Width           =   1005
         End
         Begin VB.TextBox ��ȣ�ڿ����ȣ 
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
            Height          =   165
            Left            =   2460
            TabIndex        =   14
            Text            =   "(100-100)"
            Top             =   6630
            Width           =   1005
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
            Left            =   2460
            TabIndex        =   13
            Text            =   "ȫ�浿"
            Top             =   6150
            Width           =   1545
         End
         Begin VB.TextBox ��ȣ���ּ�2 
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
            Left            =   2460
            TabIndex        =   12
            Text            =   "���� �߱� �Ŵ絿 ��������..................."
            Top             =   6870
            Width           =   3870
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
            Left            =   7710
            TabIndex        =   11
            Text            =   "��ȣ�����ֽ�ȸ��"
            Top             =   6780
            Width           =   2325
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
            Left            =   7710
            TabIndex        =   10
            Text            =   "011-9490-8607"
            Top             =   6090
            Width           =   1965
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
            Left            =   10830
            TabIndex        =   9
            Text            =   "02-2104-8600"
            Top             =   6750
            Width           =   1470
         End
         Begin VB.TextBox �л��ּ�1 
            Appearance      =   0  '���
            BorderStyle     =   0  '����
            Height          =   195
            Left            =   2460
            TabIndex        =   8
            Text            =   "���� ���ı� ������"
            Top             =   4350
            Width           =   3915
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
            Left            =   2460
            TabIndex        =   7
            Text            =   "���緿����"
            Top             =   5490
            Width           =   3990
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
            Left            =   7710
            TabIndex        =   6
            Text            =   "02-2104-8600"
            Top             =   4230
            Width           =   2955
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
            Left            =   7710
            TabIndex        =   5
            Text            =   "011-9490-8607"
            Top             =   4800
            Width           =   2955
         End
         Begin VB.TextBox �л��ּ�2 
            Appearance      =   0  '���
            BorderStyle     =   0  '����
            Height          =   195
            Left            =   2460
            TabIndex        =   4
            Text            =   "53-21 �ֿ���� ���� 201ȣ "
            Top             =   4890
            Width           =   3960
         End
         Begin VB.TextBox �������� 
            Appearance      =   0  '���
            BorderStyle     =   0  '����
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2610
            TabIndex        =   3
            Text            =   "��������"
            Top             =   8340
            Width           =   1080
         End
         Begin VB.TextBox �������� 
            Appearance      =   0  '���
            BorderStyle     =   0  '����
            BeginProperty Font 
               Name            =   "����ü"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2610
            TabIndex        =   2
            Text            =   "��������"
            Top             =   8820
            Width           =   2130
         End
         Begin VB.Line Line8 
            BorderColor     =   &H00FF0000&
            X1              =   4350
            X2              =   4350
            Y1              =   5880
            Y2              =   6510
         End
         Begin VB.Line Line7 
            BorderColor     =   &H00FF0000&
            X1              =   5220
            X2              =   5220
            Y1              =   5880
            Y2              =   6510
         End
         Begin VB.Label Label9 
            BackStyle       =   0  '����
            Caption         =   "��   ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   4470
            TabIndex        =   79
            Top             =   6090
            Width           =   705
         End
         Begin VB.Label ���� 
            BackStyle       =   0  '����
            Caption         =   "����"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   5610
            TabIndex        =   78
            Top             =   6090
            Width           =   525
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '����
            Caption         =   "*�޴���"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   19
            Left            =   6660
            TabIndex        =   77
            Top             =   5940
            Width           =   885
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '����
            Caption         =   "�޴���"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   14
            Left            =   6690
            TabIndex        =   76
            Top             =   4920
            Width           =   645
         End
         Begin VB.Label Label4 
            BackStyle       =   0  '����
            Caption         =   "*"
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
            Left            =   6660
            TabIndex        =   75
            Top             =   4650
            Width           =   195
         End
         Begin VB.Label Label5 
            BackStyle       =   0  '����
            Caption         =   "(�л���)"
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
            Left            =   6840
            TabIndex        =   74
            Top             =   4680
            Width           =   645
         End
         Begin VB.Label Label6 
            Alignment       =   2  '��� ����
            BackStyle       =   0  '����
            Caption         =   "(SMS ����"
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
            Left            =   6660
            TabIndex        =   73
            Top             =   6150
            Width           =   915
         End
         Begin VB.Label Label10 
            BackStyle       =   0  '����
            Caption         =   "��ȣ)"
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
            Left            =   6840
            TabIndex        =   72
            Top             =   6330
            Width           =   585
         End
         Begin VB.Label Label11 
            BackStyle       =   0  '����
            Caption         =   "@ ������ �ȿ��� �����Ͻÿ�. (*�� �ʼ������̰� �� �ܿ��� ���������Դϴ�.)"
            BeginProperty Font 
               Name            =   "����"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   7140
            TabIndex        =   71
            Top             =   2400
            Width           =   5865
         End
         Begin VB.Shape FillBOXs 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  '�������� ����
            BorderStyle     =   0  '����
            Height          =   480
            Index           =   3
            Left            =   2310
            Top             =   7680
            Width           =   4665
         End
         Begin VB.Line Line6 
            X1              =   2280
            X2              =   2280
            Y1              =   7650
            Y2              =   9150
         End
         Begin VB.Line Line5 
            X1              =   780
            X2              =   6960
            Y1              =   8655
            Y2              =   8655
         End
         Begin VB.Line Line4 
            X1              =   780
            X2              =   6960
            Y1              =   8160
            Y2              =   8160
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            BorderStyle     =   3  '��
            Index           =   8
            X1              =   810
            X2              =   5700
            Y1              =   7440
            Y2              =   7440
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            BorderStyle     =   3  '��
            Index           =   7
            X1              =   8010
            X2              =   12990
            Y1              =   7440
            Y2              =   7440
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
            Left            =   5850
            TabIndex        =   69
            Top             =   7380
            Width           =   2055
         End
         Begin VB.Shape Boxs 
            Height          =   1515
            Index           =   3
            Left            =   780
            Top             =   7650
            Width           =   6195
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
            Left            =   810
            TabIndex        =   68
            Top             =   1290
            Width           =   4065
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   52
            X1              =   3060
            X2              =   5730
            Y1              =   2550
            Y2              =   2550
         End
         Begin VB.Label Labels 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  '����
            Caption         =   "�й�:"
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
            Left            =   3120
            TabIndex        =   67
            Top             =   2190
            Width           =   765
         End
         Begin VB.Shape Boxs 
            BorderColor     =   &H00FF0000&
            Height          =   585
            Index           =   0
            Left            =   1200
            Top             =   1980
            Width           =   1755
         End
         Begin VB.Line Line3 
            BorderColor     =   &H00FF0000&
            X1              =   10830
            X2              =   10830
            Y1              =   2730
            Y2              =   5250
         End
         Begin VB.Label ���� 
            BackStyle       =   0  '����
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   5640
            TabIndex        =   65
            Top             =   3570
            Width           =   525
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00FF0000&
            X1              =   5250
            X2              =   5250
            Y1              =   3360
            Y2              =   3990
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00FF0000&
            X1              =   4380
            X2              =   4380
            Y1              =   3360
            Y2              =   3990
         End
         Begin VB.Label Label8 
            BackStyle       =   0  '����
            Caption         =   "��   ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   4500
            TabIndex        =   64
            Top             =   3570
            Width           =   705
         End
         Begin VB.Label Label7 
            BackStyle       =   0  '����
            Caption         =   "��   ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   6690
            TabIndex        =   63
            Top             =   6630
            Width           =   705
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            BorderStyle     =   3  '��
            Index           =   6
            X1              =   60
            X2              =   13680
            Y1              =   1140
            Y2              =   1140
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '����
            Caption         =   "���� ���� Ŭ���� ���п���"
            BeginProperty Font 
               Name            =   "����"
               Size            =   27.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   705
            Index           =   23
            Left            =   780
            TabIndex        =   62
            Top             =   270
            Width           =   7665
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   5
            X1              =   10770
            X2              =   12990
            Y1              =   690
            Y2              =   690
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
            Index           =   0
            Left            =   9540
            TabIndex        =   61
            Top             =   510
            Width           =   1275
         End
         Begin VB.Shape Boxs 
            BorderColor     =   &H00FF0000&
            Height          =   795
            Index           =   1
            Left            =   9420
            Top             =   240
            Width           =   3555
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   3
            X1              =   10740
            X2              =   10740
            Y1              =   240
            Y2              =   1020
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '����
            Caption         =   "���� ����"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   870
            TabIndex        =   57
            Top             =   8310
            Width           =   1275
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '����
            Caption         =   "  ��   ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   990
            TabIndex        =   56
            Top             =   7830
            Width           =   1065
         End
         Begin VB.Image Photo 
            Height          =   2355
            Left            =   10920
            Picture         =   "MAT010.frx":0000
            Stretch         =   -1  'True
            Top             =   2790
            Width           =   1995
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '����
            Caption         =   "��   ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   6720
            TabIndex        =   37
            Top             =   2940
            Width           =   645
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '����
            Caption         =   "��   ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   16
            Left            =   1500
            TabIndex        =   36
            Top             =   3600
            Width           =   645
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '����
            Caption         =   "�������"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   6
            Left            =   6630
            TabIndex        =   35
            Top             =   3630
            Width           =   840
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '����
            Caption         =   "��   ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   25
            Left            =   1530
            TabIndex        =   34
            Top             =   2940
            Width           =   645
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '����
            Caption         =   "��   ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   22
            Left            =   1485
            TabIndex        =   33
            Top             =   4500
            Width           =   645
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '����
            Caption         =   "��   ȭ"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   15
            Left            =   6660
            TabIndex        =   32
            Top             =   4230
            Width           =   795
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '����
            Caption         =   "E-mail"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   13
            Left            =   6705
            TabIndex        =   31
            Top             =   5460
            Width           =   645
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '����
            Caption         =   "���б�"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   21
            Left            =   1485
            TabIndex        =   30
            Top             =   5340
            Width           =   645
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '����
            Caption         =   "(��ű�)"
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
            Index           =   10
            Left            =   1500
            TabIndex        =   29
            Top             =   5595
            Width           =   615
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '����
            Caption         =   "���� ��ȭ"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   20
            Left            =   11070
            TabIndex        =   28
            Top             =   6120
            Width           =   990
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   16
            X1              =   10260
            X2              =   10260
            Y1              =   5880
            Y2              =   7170
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '����
            Caption         =   "��   ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   18
            Left            =   1500
            TabIndex        =   27
            Top             =   6180
            Width           =   645
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '����
            Caption         =   "��   ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   17
            Left            =   1470
            TabIndex        =   26
            Top             =   6750
            Width           =   645
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   4
            X1              =   6540
            X2              =   10830
            Y1              =   4620
            Y2              =   4620
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            BorderStyle     =   3  '��
            Index           =   13
            X1              =   2340
            X2              =   6540
            Y1              =   4620
            Y2              =   4620
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   0
            X1              =   1320
            X2              =   12930
            Y1              =   6510
            Y2              =   6510
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '����
            Caption         =   "(�ٹ�ó)"
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
            Left            =   6690
            TabIndex        =   25
            Top             =   6900
            Width           =   615
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '����
            Caption         =   "ȣ"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   56
            Left            =   975
            TabIndex        =   24
            Top             =   6360
            Width           =   195
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '����
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   57
            Left            =   975
            TabIndex        =   23
            Top             =   6720
            Width           =   195
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '����
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   58
            Left            =   975
            TabIndex        =   22
            Top             =   6030
            Width           =   195
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '����
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   59
            Left            =   945
            TabIndex        =   21
            Top             =   4560
            Width           =   195
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '����
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   60
            Left            =   945
            TabIndex        =   20
            Top             =   3255
            Width           =   195
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   42
            X1              =   7500
            X2              =   7500
            Y1              =   3450
            Y2              =   7170
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   44
            X1              =   2280
            X2              =   2280
            Y1              =   2730
            Y2              =   7170
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   45
            X1              =   7500
            X2              =   7500
            Y1              =   2730
            Y2              =   3495
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   46
            X1              =   1320
            X2              =   1320
            Y1              =   2730
            Y2              =   7170
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   50
            X1              =   1350
            X2              =   12990
            Y1              =   5250
            Y2              =   5250
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            BorderWidth     =   2
            Index           =   1
            X1              =   810
            X2              =   12990
            Y1              =   5880
            Y2              =   5880
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   2
            X1              =   1320
            X2              =   10830
            Y1              =   3360
            Y2              =   3360
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   11
            X1              =   1320
            X2              =   10800
            Y1              =   3990
            Y2              =   3990
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   14
            X1              =   6540
            X2              =   6540
            Y1              =   2730
            Y2              =   7170
         End
         Begin VB.Shape Boxs 
            BorderColor     =   &H00FF0000&
            BorderWidth     =   2
            Height          =   4455
            Index           =   2
            Left            =   810
            Top             =   2730
            Width           =   12195
         End
         Begin VB.Image Image1 
            Height          =   435
            Left            =   10830
            Picture         =   "MAT010.frx":1406
            Stretch         =   -1  'True
            Top             =   9210
            Width           =   2040
         End
         Begin VB.Shape FillBOXs 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  '�������� ����
            BorderStyle     =   0  '����
            Height          =   660
            Index           =   4
            Left            =   10260
            Top             =   5880
            Width           =   2715
         End
         Begin VB.Shape FillBOXs 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  '�������� ����
            BorderStyle     =   0  '����
            Height          =   4455
            Index           =   1
            Left            =   6540
            Top             =   2730
            Width           =   960
         End
         Begin VB.Shape FillBOXs 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  '�������� ����
            BorderStyle     =   0  '����
            Height          =   675
            Index           =   0
            Left            =   4380
            Top             =   3360
            Width           =   885
         End
         Begin VB.Shape FillBOXs 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  '�������� ����
            BorderStyle     =   0  '����
            Height          =   4440
            Index           =   2
            Left            =   1320
            Top             =   2730
            Width           =   960
         End
         Begin VB.Shape FillBOXs 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  '�������� ����
            BorderStyle     =   0  '����
            Height          =   795
            Index           =   6
            Left            =   9420
            Top             =   240
            Width           =   1320
         End
         Begin VB.Shape Shape2 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  '�������� ����
            BorderColor     =   &H00000000&
            BorderStyle     =   0  '����
            Height          =   1455
            Left            =   810
            Top             =   7680
            Width           =   1485
         End
         Begin VB.Shape FillBOXs 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  '�������� ����
            BorderStyle     =   0  '����
            Height          =   645
            Index           =   5
            Left            =   4350
            Top             =   5880
            Width           =   885
         End
      End
   End
   Begin MSComDlg.CommonDialog dlgPrint 
      Left            =   3420
      Top             =   10590
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   1860
      Top             =   10530
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   2490
      Top             =   10530
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
            Picture         =   "MAT010.frx":2372
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "MAT010"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################################################
'   �� �� ��  �� : �뼺�п� ���л���, �ݹ��� & �ð�ǥ ���α׷�
'   ����ý��۸� :
'   ��   ��   �� : MAT010
'   �� ��  �� �� : ���� ���� Ŭ���� ���п���
'
'   ��   ��   �� : 2009/11/26
'   ��   ��   �� : ���ϱ�
' --------------------------------------------------------------------------------------------------------------
' --------------------------------------------------------------------------------------------------------------
'                 ��     ��     ��     ��
' --------------------------------------------------------------------------------------------------------------
'   1. ������ :
'   2. ��  �� :
'################################################################################################################

Option Explicit


Private Type tSTD
    ORD_NO      As String
    EX_YN       As String
    
    SU_NO       As String
    SCHNO       As String
    ACID        As String
    EXMID       As String
    STDNM       As String
    Birth       As String
    
    PTS1        As String
    MATJUM      As String
    
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
    
    K_LEV       As String
    M_LEV       As String
    E_LEV       As String
    
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
    
    HAKYUN      As String
    E_SUKCHA    As String
    M_SUKCHA    As String
    
    ETC1        As String
    
End Type
Private uSTD() As tSTD

Private sSavePath   As String       '<< image ���
Private nTotRec     As Long


Private Sub Form_Load()
    Me.Top = 0
    Me.Left = 0
    Me.Width = 14550
    Me.Height = 10900
    
    Me.Tag = "LOAD"
        nTotRec = 0
        Call Clear_Form_Control
        
        sSavePath = App.Path & "\MPHOTO"
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
        
        '>> �迭
        With cboKaeyol
            .Clear
            .AddItem "�ι�" & Space(30) & "11"
            .AddItem "�ڿ�" & Space(30) & "12"
            .AddItem "��ü" & Space(30) & "XX"
            .ListIndex = 2
        End With
        
        ReDim uSTD(0) As tSTD
        
    Me.Tag = ""
    
    ����.Caption = ""
    ���.Text = "2013 ����"
    
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
    
    'Height = 3990
    'Width = 4890   ' ���̿� �ʺ� �����մϴ�.
    Set Photo.Picture = imgList.ListImages.Item(1).Picture
        
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
    
'    Select Case Trim(basModule.SchCD)
'        Case "N"
'
'        Case Else
'            MsgBox "�뷮�� �뼺�п��� �ƴ� ��� ��¹��� �ٸ� �� �ֽ��ϴ�.", vbExclamation + vbOKOnly, "�л���ȸ"
'    End Select
    
    On Error GoTo ErrStmt
    ReDim uSTD(0) As tSTD
    
    cmdFind.Enabled = False
        Call Get_STD_Data
        
    cmdFind.Enabled = True
    
    Exit Sub
ErrStmt:
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
    
    
    '<< �ʱ� �۾� : ��������
    '..
    
    
    On Error GoTo ErrStmt
    
    sStr = ""
    sStr = sStr & "  SELECT ROWNUM AS ID, "
    sStr = sStr & "         ORD_NO, ACACD, EXMROUND, EMAIL, USERNM, SU_NO,"
    sStr = sStr & "         HOPE_ACACD, SEX, KEYOL, Birth, "
    sStr = sStr & "         SEL1, SEL2, SEL3, SEL4, SEL5, PTS_SEL, PTS1, PTS2, "
    sStr = sStr & "         GRADE_KOR, GRADE_MAT, GRADE_ENG, GTOT, "
    sStr = sStr & "         ZIP, ADR1, ADR2,"
    sStr = sStr & "         TEL,"
    sStr = sStr & "         CEL,"
    sStr = sStr & "         HAKCD, GYEAR,"
    sStr = sStr & "         D_UNIVCD, D_MAJORCD, "
    sStr = sStr & "         FILENM, "
    sStr = sStr & "         PRTNM, PRTREL, PZIPCODE, PADR1, PADR2, PJOB,"
    sStr = sStr & "         PTEL,"
    sStr = sStr & "         JTEL,"
    sStr = sStr & "         REG_DATE,"
    sStr = sStr & "         BIGO, ACC_NO, AMNT,"
    sStr = sStr & "         MOD_REG_DATE, RECSMS, GRADE_TAM1, GRADE_TAM2, GRADE_TAM1_SELECT, GRADE_TAM2_SELECT,ETC1"
    sStr = sStr & "    FROM (SELECT ORD_NO, ACACD, EXMROUND, EMAIL, USERNM, SU_NO,"
    sStr = sStr & "                 HOPE_ACACD, SEX, NVL(KEYOL,'1') AS KEYOL, SUBSTR(birth, 1, 4)||'-'||SUBSTR(birth, 5, 2) ||'-'||SUBSTR(birth, 7, 2)  AS Birth, "
    sStr = sStr & "                 SEL1, SEL2, SEL3, SEL4, SEL5, PTS_SEL, PTS1, PTS2, "
    sStr = sStr & "                 GRADE_KOR, GRADE_MAT, GRADE_ENG, 0 AS GTOT,"
    sStr = sStr & "                 SUBSTR(ZIPCODE,1,3)||'-'||SUBSTR(ZIPCODE,4,3) AS ZIP, ADDR2 AS ADR1, ADDR AS ADR2,"
    sStr = sStr & "                 TEL1||'-'||TEL2||'-'||TEL3 AS TEL,"
    sStr = sStr & "                 CEL1||'-'||CEL2||'-'||CEL3 AS CEL,"
    sStr = sStr & "                 GET_SCHOOLNM(HAKCD) AS HAKCD, GYEAR,"
    sStr = sStr & "                 D_UNIVCD, D_MAJORCD, "
    sStr = sStr & "                 FILENM, "
    sStr = sStr & "                 PRTNM, PRTREL, "
    sStr = sStr & "                 SUBSTR(PZIPCODE,1,3)||'-'||SUBSTR(PZIPCODE,4,3) AS PZIPCODE, PADDR2 AS PADR1, PADDR AS PADR2, PJOB,"
    sStr = sStr & "                 PTEL1||'-'||PTEL2||'-'||PTEL3 AS PTEL,"
    sStr = sStr & "                 JTEL1||'-'||JTEL2||'-'||JTEL3 AS JTEL,"
    sStr = sStr & "                 REG_DATE,"
    sStr = sStr & "                 BIGO, ACC_NO, AMNT,"
    sStr = sStr & "                 MOD_REG_DATE, RECSMS, GRADE_TAM1, GRADE_TAM2, GRADE_TAM1_SELECT, GRADE_TAM2_SELECT,ETC1"
    sStr = sStr & "            FROM HWSIN01TB_WINTER"
    sStr = sStr & "           WHERE EXMROUND LIKE "
    
    Select Case Trim(SchCD)
        Case "N"
            sStr = sStr & "         'NR081126%'"
        Case "K"
            sStr = sStr & "         'KN081126%'"
        Case "S"
            sStr = sStr & "         'SP081126%'"
        Case "P"
            sStr = sStr & "         'MK081126%'"
        Case "M"
            sStr = sStr & "         'NR081126%'"
            
        Case "W"
            sStr = sStr & "         'KN081126%'"
        Case "Q"
            sStr = sStr & "         'KN081126%'"
            
        Case "J"
            sStr = sStr & "         'YJ081126%'"
        Case "B"
            sStr = sStr & "         'BS081126%'"
        
        Case Else
            sStr = sStr & "         'BS081126%'"
    End Select
    
    
'>> �迭
    Select Case Trim(Right(cboKaeyol, 30))
        Case "XX"
            sStr = sStr & "     AND KEYOL IN ('11','12')"
        Case "11", "13"
            sStr = sStr & "     AND KEYOL = '11' "
        Case "12"
            sStr = sStr & "     AND KEYOL = '12' "
    End Select
    
'>> �����ȣ
'    If Trim(fpExmID_S.UnFmtText) > " " And Trim(fpExmID_E.UnFmtText) > " " Then
'        sStr = sStr & "         AND ORD_NO BETWEEN '" & Trim(fpExmID_S.UnFmtText) & "' AND '" & Trim(fpExmID_E.UnFmtText) & "'"
'    ElseIf Trim(fpExmID_S.UnFmtText) > " " And Trim(fpExmID_E.UnFmtText) = " " Then
'        sStr = sStr & "         AND ORD_NO BETWEEN '" & Trim(fpExmID_S.UnFmtText) & "' AND '999999' "
'    ElseIf Trim(fpExmID_S.UnFmtText) = " " And Trim(fpExmID_E.UnFmtText) > " " Then
'        sStr = sStr & "         AND ORD_NO BETWEEN '000000' AND " & Trim(fpExmID_E.UnFmtText)
'    ElseIf Trim(fpExmID_S.UnFmtText) = " " And Trim(fpExmID_E.UnFmtText) = " " Then
'        ' no action
'    End If
    
    If Trim(fpExmID_S.UnFmtText) > " " And Trim(fpExmID_E.UnFmtText) > " " Then
        sStr = sStr & "         AND SU_NO BETWEEN '" & Trim(fpExmID_S.UnFmtText) & "' AND '" & Trim(fpExmID_E.UnFmtText) & "'"
    ElseIf Trim(fpExmID_S.UnFmtText) > " " And Trim(fpExmID_E.UnFmtText) = " " Then
        sStr = sStr & "         AND SU_NO BETWEEN '" & Trim(fpExmID_S.UnFmtText) & "' AND '999999' "
    ElseIf Trim(fpExmID_S.UnFmtText) = " " And Trim(fpExmID_E.UnFmtText) > " " Then
        sStr = sStr & "         AND SU_NO BETWEEN '000000' AND " & Trim(fpExmID_E.UnFmtText)
    ElseIf Trim(fpExmID_S.UnFmtText) = " " And Trim(fpExmID_E.UnFmtText) = " " Then
        ' no action
    End If
    
'>> �л���
    If Trim(txtStdNM.Text) > " " Then
        sStr = sStr & "         AND USERNM LIKE '" & Trim(txtStdNM.Text) & "%'"
    End If
    
    sStr = sStr & "           ORDER BY ORD_NO "
    sStr = sStr & "          ) "
    sStr = sStr & "    WHERE ORD_NO > 0 "
    sStr = sStr & "      AND KEYOL <> '3' "
    
'    Text1.Text = sStr
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
        If .RecordCount > 0 Then
            nTotRec = .RecordCount
            
            .MoveFirst
            
            ReDim uSTD(.RecordCount) As tSTD
            
            VScroll1.Max = .RecordCount
            VScroll1.Enabled = True
            
            For nRec = 1 To .RecordCount Step 1
            
                If IsNull(.Fields("SU_NO")) = False Then uSTD(nRec).SU_NO = .Fields("SU_NO")
                If IsNull(.Fields("ORD_NO")) = False Then uSTD(nRec).SCHNO = .Fields("ORD_NO")
                If IsNull(.Fields("ORD_NO")) = False Then uSTD(nRec).ORD_NO = .Fields("ORD_NO")
                
                
                If IsNull(.Fields("PTS1")) = False Then uSTD(nRec).PTS1 = .Fields("PTS1")
                If IsNull(.Fields("GRADE_MAT")) = False Then uSTD(nRec).MATJUM = .Fields("GRADE_MAT")
                
                
                If IsNull(.Fields("ACACD")) = False Then uSTD(nRec).ACID = .Fields("ACACD")
                If IsNull(.Fields("EXMROUND")) = False Then uSTD(nRec).EXMID = .Fields("EXMROUND")
                If IsNull(.Fields("USERNM")) = False Then uSTD(nRec).STDNM = .Fields("USERNM")
                If IsNull(.Fields("Birth")) = False Then uSTD(nRec).Birth = .Fields("Birth")
                
                If IsNull(.Fields("EXMROUND")) = False Then
                    If Right(.Fields("EXMROUND"), 1) = "1" Then
                        uSTD(nRec).EX_YN = "������"
                    ElseIf uSTD(nRec).EX_YN = "2" Then
                        uSTD(nRec).EX_YN = "������"
                    End If
                End If
                
                If IsNull(.Fields("KEYOL")) = False Then uSTD(nRec).KAEYOL = .Fields("KEYOL")
                
                If IsNull(.Fields("SEL1")) = False Then uSTD(nRec).SEL1 = .Fields("SEL1")
                If IsNull(.Fields("SEL2")) = False Then uSTD(nRec).SEL2 = .Fields("SEL2")
                If IsNull(.Fields("SEL3")) = False Then uSTD(nRec).SEL3 = .Fields("SEL3")
                If IsNull(.Fields("SEL4")) = False Then uSTD(nRec).SEL4 = .Fields("SEL4")
                If IsNull(.Fields("SEL5")) = False Then uSTD(nRec).SEL5 = .Fields("SEL5")
                
                If IsNull(.Fields("GRADE_KOR")) = False Then uSTD(nRec).K_LEV = .Fields("GRADE_KOR")
                If IsNull(.Fields("GRADE_MAT")) = False Then uSTD(nRec).M_LEV = .Fields("GRADE_MAT")
                If IsNull(.Fields("GRADE_ENG")) = False Then uSTD(nRec).E_LEV = .Fields("GRADE_ENG")
                'If IsNull(.Fields("GTOT")) = False Then uSTD(nRec).TOT_NUM = .Fields("GTOT")
                
                '## �����п� - WINTER�� �ʿ����.
                If IsNull(.Fields("ACACD")) = False Then uSTD(nRec).ACID = .Fields("ACACD")
                    Select Case Trim(.Fields("ACACD"))
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
                            
                    End Select
                
                If IsNull(.Fields("ACACD")) = False Then uSTD(nRec).ACID = .Fields("ACACD")
                    Select Case Trim(.Fields("ACACD"))
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
                            
                    End Select
                
                'If IsNull(.Fields("PASS1")) = False Then uSTD(nRec).PASS1 = .Fields("PASS1")
                'If IsNull(.Fields("PASS2")) = False Then uSTD(nRec).PASS2 = .Fields("PASS2")
                'If IsNull(.Fields("PASS3")) = False Then uSTD(nRec).PASS3 = .Fields("PASS3")
                'If IsNull(.Fields("PASS4")) = False Then uSTD(nRec).PASS4 = .Fields("PASS4")
                
                'If IsNull(.Fields("CL_CLOSE")) = False Then uSTD(nRec).CL_CLOSE = .Fields("CL_CLOSE")
                'If IsNull(.Fields("CY_ACNT")) = False Then uSTD(nRec).CY_ACNT = .Fields("CY_ACNT")
                If IsNull(.Fields("AMNT")) = False Then uSTD(nRec).TOT_AMT = .Fields("AMNT")
                
                If IsNull(.Fields("SEX")) = False Then uSTD(nRec).SEX = .Fields("SEX")
                
                If IsNull(.Fields("ZIP")) = False Then uSTD(nRec).ZIP = .Fields("ZIP")
                If IsNull(.Fields("ADR1")) = False Then uSTD(nRec).ADDR1 = .Fields("ADR1")
                If IsNull(.Fields("ADR2")) = False Then uSTD(nRec).ADDR2 = .Fields("ADR2")
                
                If IsNull(.Fields("TEL")) = False Then uSTD(nRec).TEL = .Fields("TEL")
                If IsNull(.Fields("CEL")) = False Then uSTD(nRec).CEL = .Fields("CEL")
                If IsNull(.Fields("EMAIL")) = False Then uSTD(nRec).EMAIL = .Fields("EMAIL")
                
                If IsNull(.Fields("HAKCD")) = False Then uSTD(nRec).HIGH_SCH = .Fields("HAKCD")
                If IsNull(.Fields("GYEAR")) = False Then uSTD(nRec).GRADE_YEAR = .Fields("GYEAR")
                
                If IsNull(.Fields("PRTNM")) = False Then uSTD(nRec).PRNT_NM = .Fields("PRTNM")
                If IsNull(.Fields("PRTREL")) = False Then uSTD(nRec).PRNT_RLTN = .Fields("PRTREL")
                
                If IsNull(.Fields("PZIPCODE")) = False Then uSTD(nRec).PRNT_ZIP = .Fields("PZIPCODE")
                If IsNull(.Fields("PADR1")) = False Then uSTD(nRec).PRNT_ADDR1 = .Fields("PADR1")
                If IsNull(.Fields("PADR2")) = False Then uSTD(nRec).PRNT_ADDR2 = .Fields("PADR2")
                If IsNull(.Fields("PTEL")) = False Then uSTD(nRec).PRNT_CEL = .Fields("PTEL")
                If IsNull(.Fields("JTEL")) = False Then uSTD(nRec).PRNT_TEL = .Fields("JTEL")
                If IsNull(.Fields("PJOB")) = False Then uSTD(nRec).PRNT_JOB = .Fields("PJOB")
                'If IsNull(.Fields("PRNT_W_TEL")) = False Then uSTD(nRec).PRNT_W_TEL = .Fields("PRNT_W_TEL")
                
                If IsNull(.Fields("FILENM")) = False Then uSTD(nRec).PHOTO_PATH = .Fields("FILENM")
                
                If IsNull(.Fields("BIGO")) = False Then uSTD(nRec).HAKYUN = .Fields("BIGO")
                
                'If IsNull(.Fields("E_SUKCHA")) = False Then uSTD(nRec).E_SUKCHA = .Fields("E_SUKCHA")
                'If IsNull(.Fields("M_SUKCHA")) = False Then uSTD(nRec).M_SUKCHA = .Fields("M_SUKCHA")
                If IsNull(.Fields("ETC1")) = False Then uSTD(nRec).ETC1 = .Fields("ETC1")
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


'## ������ �̹��� ��������
Private Sub Get_STD_image()
    
    Dim bData()     As Byte
    Dim f           As Integer
    Dim nRec        As Long

    Dim sLocalFile  As String
    Dim sSourceUrl  As String

    On Error Resume Next

    f = FreeFile()
    
    For nRec = 1 To UBound(uSTD) Step 1
    
        '2010.12.20 ���ѿ� �뷮��,����,���� ��� ���������� �����ȣ�� ����
        
        Select Case Trim(SchCD)
            Case "N"
                sLocalFile = sSavePath & "\" & uSTD(nRec).SU_NO & ".jpg"                    '�����ȣ
            Case "S"
                sLocalFile = sSavePath & "\" & uSTD(nRec).SU_NO & ".jpg"                    '�����ȣ
            Case "J"
                sLocalFile = sSavePath & "\" & uSTD(nRec).SU_NO & ".jpg"                    '�����ȣ
            Case Else
                sLocalFile = sSavePath & "\" & uSTD(nRec).ORD_NO & ".jpg"                   '<< unique key : ORD_NO
        End Select
        
        If Dir(sLocalFile, vbNormal) = "" Then                                                '<< �л� �̹��� ���� �͸� ����
            If uSTD(nRec).PHOTO_PATH > " " Then
            
            
                Select Case Trim(basModule.SchCD)
                    Case "B"
                        sSourceUrl = "http://www.dsnschool.net" & uSTD(nRec).PHOTO_PATH        '<< ������ �̹��� ���
                        
                    Case Else
                        sSourceUrl = "http://www.dshw.co.kr" & uSTD(nRec).PHOTO_PATH        '<< ������ �̹��� ���
                        
                End Select
                
                bData() = Inet1.OpenURL(sSourceUrl, icByteArray)
                
                If UBound(bData) > 0 Then
                    Open sLocalFile For Binary Access Write As #f
                    Put #f, , bData()
                
                    DoEvents
                    Close #f
                End If
            End If
        End If
    Next nRec
    
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
    Dim sTmp        As String
    
    If Me.Tag = "LOAD" Then Exit Sub
    
    If UBound(uSTD) < 1 Then Exit Sub
    If UBound(uSTD) < Index Then Exit Sub
    
    With uSTD(Index)
    
        If Trim(.HAKYUN) = "4" Then
            �г�.Text = "�����"
            �г�1.Text = "�����"
        Else
            �г�.Text = .HAKYUN & "�г�"
            �г�1.Text = .HAKYUN & "�г�"
        End If
        
        
        Select Case Trim(.KAEYOL)   '<< �迭: 01,02,03-�ι�,�ڿ�,��ü   06,05-�����ι�,�ڿ�  06,07 -��������,�Ǵ�
            Case "11"
                    �����迭(0).Text = "�� �� ��"
                    �����迭(1).Text = "�� �� ��"
                    �����迭(2).Text = "�� �� ��"
            Case "12"
                    �����迭(0).Text = "�� �� ��"
                    �����迭(1).Text = "�� �� ��"
                    �����迭(2).Text = "�� �� ��"
            Case Else
                    �����迭(0).Text = ""
                    �����迭(1).Text = ""
                    �����迭(2).Text = ""
        End Select
        
        
        If Trim(.PTS1) = "1" Then
            ��������.Text = "����"
        ElseIf Trim(.PTS1) = "2" Then
            ��������.Text = "����"
        End If
        
        '�������� 3�г� ���ϸ� ������, ������̸� ǥ������
        If CInt(.HAKYUN) < 4 Then
            ��������.Text = "���� ������"   '1,2,3�г�
        Else
            ��������.Text = "���� ǥ������" '�����
        End If
        
         If Trim(.MATJUM) = "0" Or Trim(.MATJUM) = "" Then
            ��������.Text = "X"
        Else
            ��������.Text = ��������.Text & " " & Trim(.MATJUM) & " ��"
        End If
        
        
        '���
        Select Case Trim(.ETC1)
            Case "1"
                ���.Text = "2013 ����"
            Case "2"
                ���.Text = "6�� �򰡿�"
            Case "3"
                ���.Text = "9�� �򰡿�"
            Case "4"
                ���.Text = "��2 �뼺���ǰ��"
            Case "5"
                ���.Text = "��2 ����û���ǰ��"
            Case "9"
                ���.Text = "���ŵ��"
            Case Else
                ���.Text = "X"
        End Select
        
        
'        '2011�� ���� Ŭ���� ���� ���� ���� �뷮�� ���� ����
'        '���ѿ� ����
'        Select Case Trim(SchCD)
'            Case "N"
'            Select Case Trim(.ETC1)
'                Case "1"
'                        M_Title(0).Text = "12�г⵵����"        '2012.11.22�� ���� ���� ���� 3
'                Case "2"
'                        M_Title(0).Text = "6�� �򰡿�"          '2012.11.22�� ���� ���� ���� 1
'                Case "3"
'                        M_Title(0).Text = "9�� �򰡿�"          '2012.11.22�� ���� ���� ���� 2
'                Case Else
'                        M_Title(0).Text = "������"
'            End Select
'
'            Case "S"
'            Select Case Trim(.ETC1)
'                Case "1"
'                        M_Title(0).Text = "12�г⵵����"
'                Case "2"
'                        M_Title(0).Text = "6�� �򰡿�"
'                Case "3"
'                        M_Title(0).Text = "9�� �򰡿�"
'                Case Else
'                        M_Title(0).Text = "������"
'            End Select
'
'            Case "J"
'            Select Case Trim(.ETC1)
'                Case "1"
'                        M_Title(0).Text = "12�г⵵����"
'                Case "2"
'                        M_Title(0).Text = "6�� �򰡿�"
'                Case "3"
'                        M_Title(0).Text = "9�� �򰡿�"
'                Case Else
'                        M_Title(0).Text = "������"
'            End Select
'        '���ѿ� ��
'        End Select
        
       
    
        
        �����ȣ.Text = .SU_NO
        �л�����.Text = .STDNM
        �������.Text = .Birth
        ����.Caption = IIf(.SEX = 1, "��", "��")
        �л������ȣ.Text = "(" & .ZIP & ")"
        �л��ּ�1.Text = .ADDR1
        �л��ּ�2.Text = .ADDR2
        
        �л���Ű�.Text = .HIGH_SCH
        �л��̸���.Text = .EMAIL
        �л�����ó_��.Text = .TEL
        �л�����ó_�޴���.Text = .CEL
        
        ��ȣ�ڼ���.Text = .PRNT_NM
        ����.Caption = IIf(.PRNT_RLTN = 1, "��", "��")
        ��ȣ�ڿ���ó_�޴���.Text = .PRNT_CEL
        ��ȣ�ڿ����ȣ.Text = "(" & .PRNT_ZIP & ")"
        ��ȣ���ּ�1.Text = .PRNT_ADDR1
        ��ȣ���ּ�2.Text = .PRNT_ADDR2
        
        
        ��ȣ������.Text = .PRNT_JOB
        ��ȣ�ڿ���ó_����.Text = .PRNT_TEL
                       
        
        '<< ����
        Call Div_Gwamok_NM("SEL1", .SEL1)
        Call Div_Gwamok_NM("SEL4", .SEL4)
        
        '<< ����
        sTmp = ""
        sTmp = sTmp & "����(" & Trim(.E_LEV) & "), "
        sTmp = sTmp & "����(" & Trim(.M_LEV) & ")"
        '�б�����.Text = sTmp
        
        '��2����.Text = .SEL2_SCH
        
        '2010.12.20 ���ѿ� �뷮��,����,������ ��� �����ȣ�� �� ���� ǥ��
        Select Case Trim(SchCD)
            Case "N"
                Set Photo.Picture = CheckJPG(sSavePath & "\" & .SU_NO & ".jpg")
            Case "S"
                Set Photo.Picture = CheckJPG(sSavePath & "\" & .SU_NO & ".jpg")
            Case "J"
                Set Photo.Picture = CheckJPG(sSavePath & "\" & .SU_NO & ".jpg")
            Case Else
                Set Photo.Picture = CheckJPG(sSavePath & "\" & .SCHNO & ".jpg")
        End Select
        
    End With
    
End Sub

'<< ����ֱ� : �迭�� �Ǿ������� ������ ��!!
Private Sub Div_Gwamok_NM(ByVal aGbn As String, ByVal aGwamok As String)
    Dim sDiv()      As String
    Dim ni          As Integer
    
    Dim sTmp        As String
    
    On Error Resume Next
    
    sDiv = Split(aGwamok, "|", -1, vbTextCompare)
    
    For ni = 0 To UBound(sDiv) Step 1
        
        Select Case aGbn
            Case "SEL1"
            
                sTmp = ""
                Select Case Trim(sDiv(ni))
                    Case "1"
                        sTmp = "����"
                    Case "2"
                        sTmp = "����"
                    Case "3"
                        sTmp = "����"
                    Case "4"
                        sTmp = "�ѱ�������"
                    Case "5"
                        sTmp = "�����"
                    Case "6"
                        sTmp = "��������"
                    Case "7"
                        sTmp = "�ѱ�����"
                    Case "8"
                        sTmp = "��ġ"
                    Case "9"
                        sTmp = "��ȸ��ȭ"
                    Case "10"
                        sTmp = "������ȸ"
                    Case "11"
                        sTmp = "��������"
                End Select

                
            Case "SEL4"
            
                sTmp = ""
                Select Case Trim(sDiv(ni))
                    Case "1"
                        sTmp = "����"
                    Case "2"
                        sTmp = "ȭ��"
                    Case "3"
                        sTmp = "�������"
                    Case "4"
                        sTmp = "��������"
                End Select

                
            Case Else
                ' skip
                
                
        End Select
    Next ni
    
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
    If (header(0) = 255 And header(1) = 216) And _
       (tailer(0) = 255 And tailer(1) >= 209) Then
        Set CheckJPG = LoadPicture(fileName)
    Else
        Set CheckJPG = imgList.ListImages.Item(1).Picture       '<< no-image
    End If
    'Set CheckJPG = LoadPicture(fileName)
End Function




'## ��ü ���
Private Sub cmdPrintAll_Click()

    Dim nRec        As Long
    Dim bChk        As Boolean

    If UBound(uSTD) < 1 Then
        Exit Sub
    End If
    
    On Error GoTo ErrPrint
    
    
    
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
                '********************************************************************
                '  �׵θ� ���� �簢 �ڽ��� ����� ���λ��� ĥ�Ѵ�.
                '********************************************************************
                 Printer.DrawWidth = 1                   ' ���� ����
                 Printer.FillStyle = vbFSTransparent     ' �ܻ�
                 Printer.FillColor = &HC1F1FF            ' ���� ĥ�ϱ�
                 PrintFilledBox .Left * pRate, .Top * pRate, .Width * pRate, .Height * pRate, &HC1F1FF
             End If
        End With
    Next

    For Each UsrCtl In Me
        With UsrCtl
             If (UCase(TypeName(UsrCtl)) = "SHAPE" And UCase(UsrCtl.Name) = "BOXS") Then
                '********************************************************************
                '  line�� �̿��� box�����(�⺻������ shape�� ��½� line�� �̿��Ѵ�)
                '********************************************************************
                 Printer.DrawWidth = 12
                 PrintBox .Left * pRate, .Top * pRate, .Width * pRate, .Height * pRate
             End If
        End With
    Next


    For Each UsrCtl In Me
        With UsrCtl
             Select Case UCase(TypeName(UsrCtl))
                    Case "LINE"
                         '********************************************************************
                         '  �ڽ�/line�� �ߴ´�.
                         '********************************************************************
                          Printer.DrawStyle = IIf(UsrCtl.BorderStyle = 3, 2, UsrCtl.BorderStyle)
                          Printer.DrawWidth = IIf(UsrCtl.BorderStyle = 3, 1, UsrCtl.BorderWidth * 4)
                          Printer.FillStyle = vbFSTransparent
                          PrintLine .X1 * pRate, .Y1 * pRate, .X2 * pRate, .Y2 * pRate

                    Case "LABEL"
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

                    Case "TEXTBOX"
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
                         
                    Case "IMAGE"
                          '********************************************************************
                          '  �������
                          '********************************************************************
                          If (Photo.Picture <> 0) Then
                              Printer.FontTransparent = True
                              iBKMode = SetBkMode(Printer.hDC, OPAQUE)
                              ' iBKMode = SetBkMode(Printer.hDC, TRANSPARENT)
                              PrintPicture .Picture, .Left * pRate, .Top * pRate, .Width * pRate, .Height * pRate
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
    Dim sPhotoPath      As String

    Dim bRet            As String
    
    Dim sDiv()          As String
    Dim nS              As Long
    Dim sLocalFile      As String
    Dim sFilePath       As String
    
    If Button <> vbRightButton Then
        Exit Sub
    End If

    If �л�����.Text = "" Then
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
        
        
        simageFile = ""
        sPhotoPath = ""
        
        bRet = ""
        If Trim(sPhotoPath) = "" Then           '< �̹����� ���� ��쿣 ������ ����
            bRet = Make_image_Path(sOrdNO, sExmID, simageFile)
            
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
        sLocalFile = sSavePath & "\" & uSTD(nS).ORD_NO & ".jpg"       '<< unique key : ord_no
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
Private Function Make_image_Path(ByVal aOrdNO As String, ByVal aExmID As String, ByVal aimageFile As String) As String
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
            sFilePath = "/NDOC/dshw/noryangjin/register/ETC/"
        Case "K", "W", "Q"
            sFilePath = "/NDOC/dshw/kangnam/register/ETC/"
        Case "S"
            sFilePath = "/NDOC/dshw/songpa/register/ETC/"
        Case "P"
            sFilePath = "/NDOC/dshw/msongpa/register/ETC/"
        Case "M"
            sFilePath = "/NDOC/dshw/mkangnam/register/ETC/"
        Case "J"
            sFilePath = "/NDOC/dshw/mgwanghwa/register/ETC/"
        Case "B"
            sFilePath = "/NDOC/dshw/busan/register/ETC/"
    End Select
    
    sFilePath = sFilePath & aOrdNO & ".jpg"
    
    On Error GoTo ErrStmt
    
    basDataBase.DBConn.BeginTrans
    
    Set DBCmd = New ADODB.Command
    Set DBParam = New ADODB.Parameter
    
    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
            
            
    
    '<< UPDATE
    sStr = ""
    sStr = sStr & " Update HWSIN01TB_WINTER"
    sStr = sStr & "    SET FILENM = '" & sFilePath & "'"
    sStr = sStr & "  WHERE ORD_NO = '" & Trim(aOrdNO) & "'"
            
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
        sLocalFile = sSavePath & "\" & aOrdNO & ".jpg"              '<< unique key : ORD_NO
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





