VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form TMR020 
   Caption         =   "�ð�ǥ ����� >> �̵����� �ð�ǥ ���"
   ClientHeight    =   8940
   ClientLeft      =   1140
   ClientTop       =   1785
   ClientWidth     =   13815
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8940
   ScaleWidth      =   13815
   Begin VB.Frame Frame11 
      BackColor       =   &H0082C8E8&
      BorderStyle     =   0  '����
      Caption         =   "Frame11"
      Height          =   3435
      Left            =   120
      TabIndex        =   42
      Top             =   10290
      Width           =   13305
      Begin VB.Frame Frame10 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  '����
         Caption         =   "Frame10"
         Height          =   3375
         Left            =   30
         TabIndex        =   43
         Top             =   30
         Width           =   13245
         Begin FPSpread.vaSpread sprCopyLsn 
            Height          =   1605
            Left            =   510
            TabIndex        =   44
            Top             =   60
            Width           =   12675
            _Version        =   393216
            _ExtentX        =   22357
            _ExtentY        =   2831
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
            SpreadDesigner  =   "TMR020.frx":0000
         End
         Begin FPSpread.vaSpread sprBaseLsn 
            Height          =   1485
            Left            =   510
            TabIndex        =   46
            Top             =   1800
            Width           =   12675
            _Version        =   393216
            _ExtentX        =   22357
            _ExtentY        =   2619
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
            SpreadDesigner  =   "TMR020.frx":49EE
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '����
            Caption         =   "�۾� ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   780
            Left            =   30
            TabIndex        =   47
            Top             =   1890
            Width           =   585
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '����
            Caption         =   "�ӽ��۾� ȭ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   780
            Left            =   30
            TabIndex        =   45
            Top             =   60
            Width           =   585
         End
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00C6AD84&
      BorderStyle     =   0  '����
      Caption         =   "Frame6"
      Height          =   4455
      Left            =   90
      TabIndex        =   36
      Top             =   5790
      Width           =   13365
      Begin VB.Frame Frame7 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  '����
         Caption         =   "Frame7"
         Height          =   4395
         Left            =   30
         TabIndex        =   37
         Top             =   30
         Width           =   13305
         Begin VB.CheckBox chkOKNot 
            BackColor       =   &H00C0FFFF&
            Caption         =   "�Ϸ�����"
            Height          =   435
            Left            =   1830
            TabIndex        =   24
            Top             =   60
            Width           =   675
         End
         Begin VB.CommandButton cmdOrdGwamok_View 
            Caption         =   "�л���û���� ��ģ���� ����"
            Height          =   405
            Left            =   10800
            TabIndex        =   27
            Top             =   60
            Width           =   2505
         End
         Begin VB.Frame Frame9 
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   0  '����
            Caption         =   "Frame9"
            Height          =   555
            Left            =   2550
            TabIndex        =   38
            Top             =   0
            Width           =   8265
            Begin VB.Frame Frame8 
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  '����
               Caption         =   "Frame8"
               Height          =   495
               Left            =   30
               TabIndex        =   39
               Top             =   30
               Width           =   8205
               Begin VB.ComboBox cboLsnin 
                  Height          =   300
                  Left            =   6780
                  Style           =   2  '��Ӵٿ� ���
                  TabIndex        =   48
                  Top             =   90
                  Width           =   1425
               End
               Begin EditLib.fpDateTime fpCL_Close 
                  Height          =   315
                  Left            =   3930
                  TabIndex        =   41
                  Top             =   90
                  Width           =   795
                  _Version        =   196608
                  _ExtentX        =   1402
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
                  Text            =   "2007-11"
                  DateCalcMethod  =   0
                  DateTimeFormat  =   5
                  UserDefinedFormat=   "yyyy-mm"
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
                  Appearance      =   2
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
               Begin VB.CommandButton cmdStdGwamokSave 
                  Caption         =   "����ó������ ����ϱ�"
                  Height          =   405
                  Left            =   4740
                  TabIndex        =   40
                  Top             =   30
                  Width           =   1995
               End
               Begin VB.CommandButton cmdStdGwamokChk 
                  Caption         =   "�������ɿ��� Ȯ��"
                  Height          =   405
                  Left            =   30
                  TabIndex        =   25
                  Top             =   30
                  Width           =   1725
               End
               Begin VB.CommandButton cmdStdGwamokChk_Show 
                  Caption         =   "�������� ó������ ����"
                  Height          =   405
                  Left            =   1800
                  TabIndex        =   26
                  Top             =   30
                  Width           =   2085
               End
            End
         End
         Begin VB.CommandButton cmdFind_LSN_in_STD 
            Caption         =   "���ù��� �л���ȸ"
            Height          =   405
            Left            =   60
            TabIndex        =   23
            Top             =   60
            Width           =   1725
         End
         Begin FPSpread.vaSpread sprSTD 
            Height          =   3795
            Left            =   60
            TabIndex        =   28
            Top             =   570
            Width           =   13155
            _Version        =   393216
            _ExtentX        =   23204
            _ExtentY        =   6694
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
            MaxCols         =   23
            SpreadDesigner  =   "TMR020.frx":93DC
         End
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H0082C8E8&
      BorderStyle     =   0  '����
      Caption         =   "Frame5"
      Height          =   2775
      Left            =   60
      TabIndex        =   32
      Top             =   60
      Width           =   11895
      Begin VB.Frame Frame4 
         BackColor       =   &H00D2EAF5&
         BorderStyle     =   0  '����
         Caption         =   "Frame4"
         Height          =   2715
         Left            =   30
         TabIndex        =   33
         Top             =   30
         Width           =   11835
         Begin VB.CheckBox chkAll 
            BackColor       =   &H00D2EAF5&
            Caption         =   "�۾�����"
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
            Left            =   1860
            TabIndex        =   3
            Top             =   510
            Width           =   1095
         End
         Begin VB.ComboBox cboLsnType 
            Height          =   300
            Left            =   3570
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   1
            Top             =   90
            Width           =   1755
         End
         Begin VB.ComboBox cboKaeyol 
            Height          =   300
            Left            =   990
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   0
            Top             =   90
            Width           =   1755
         End
         Begin VB.CommandButton cmdFind_STD_Subj 
            Caption         =   "�ݺ� ���� ��û���� ��ȸ"
            Height          =   465
            Left            =   6240
            TabIndex        =   2
            Top             =   0
            Width           =   2445
         End
         Begin FPSpread.vaSpread sprLsn 
            Height          =   2145
            Left            =   90
            TabIndex        =   4
            Top             =   510
            Width           =   11655
            _Version        =   393216
            _ExtentX        =   20558
            _ExtentY        =   3784
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
            SpreadDesigner  =   "TMR020.frx":B255
         End
         Begin VB.Label Label1 
            Alignment       =   1  '������ ����
            BackStyle       =   0  '����
            Caption         =   "������"
            Height          =   210
            Left            =   2580
            TabIndex        =   35
            Top             =   165
            Width           =   975
         End
         Begin VB.Label Label5 
            Alignment       =   1  '������ ����
            BackStyle       =   0  '����
            Caption         =   "�迭"
            Height          =   210
            Left            =   -30
            TabIndex        =   34
            Top             =   165
            Width           =   975
         End
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C6AD84&
      BorderStyle     =   0  '����
      Caption         =   "Frame2"
      Height          =   2865
      Left            =   60
      TabIndex        =   29
      Top             =   2880
      Width           =   13395
      Begin VB.Frame Frame3 
         BackColor       =   &H00F7EFE7&
         BorderStyle     =   0  '����
         Caption         =   "Frame3"
         Height          =   2355
         Left            =   12090
         TabIndex        =   31
         Top             =   480
         Width           =   1275
         Begin VB.CommandButton cmdFindLsn 
            Caption         =   "��ϳ�����ȸ"
            Height          =   435
            Left            =   0
            TabIndex        =   22
            Top             =   1290
            Width           =   1275
         End
         Begin VB.CommandButton cmdDeleteLsn 
            Caption         =   "�̵��� ����"
            Height          =   435
            Left            =   0
            TabIndex        =   21
            Top             =   660
            Width           =   1275
         End
         Begin VB.CommandButton cmdinSertLsn 
            Caption         =   "�̵��� ���"
            Height          =   435
            Left            =   0
            TabIndex        =   20
            Top             =   180
            Width           =   1275
         End
      End
      Begin VB.TextBox txtLsnType 
         Height          =   300
         Left            =   2040
         TabIndex        =   18
         Text            =   "txtLsnType"
         Top             =   300
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.TextBox txtKaeyol 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   270
         TabIndex        =   17
         Text            =   "txtKaeyol"
         Top             =   300
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00F7EFE7&
         BorderStyle     =   0  '����
         Caption         =   "Frame1"
         Height          =   435
         Left            =   30
         TabIndex        =   30
         Top             =   30
         Width           =   13335
         Begin VB.OptionButton optTamgu 
            BackColor       =   &H00FFFFFF&
            Caption         =   "���񳻿�"
            Height          =   240
            Index           =   0
            Left            =   0
            TabIndex        =   5
            Top             =   60
            Width           =   1125
         End
         Begin VB.OptionButton optTamgu 
            BackColor       =   &H00C0C0FF&
            Caption         =   "Option1"
            Height          =   240
            Index           =   1
            Left            =   1170
            TabIndex        =   6
            Top             =   60
            Width           =   1065
         End
         Begin VB.OptionButton optTamgu 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Option1"
            Height          =   240
            Index           =   2
            Left            =   2280
            TabIndex        =   7
            Top             =   60
            Width           =   1065
         End
         Begin VB.OptionButton optTamgu 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Option1"
            Height          =   240
            Index           =   3
            Left            =   3390
            TabIndex        =   8
            Top             =   60
            Width           =   1065
         End
         Begin VB.OptionButton optTamgu 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Option1"
            Height          =   240
            Index           =   4
            Left            =   4500
            TabIndex        =   9
            Top             =   60
            Width           =   1065
         End
         Begin VB.OptionButton optTamgu 
            BackColor       =   &H0000C0C0&
            Caption         =   "Option1"
            Height          =   240
            Index           =   5
            Left            =   5610
            TabIndex        =   10
            Top             =   60
            Width           =   1065
         End
         Begin VB.OptionButton optTamgu 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Option1"
            Height          =   240
            Index           =   6
            Left            =   6720
            TabIndex        =   11
            Top             =   60
            Width           =   1065
         End
         Begin VB.OptionButton optTamgu 
            BackColor       =   &H00FF80FF&
            Caption         =   "Option1"
            Height          =   240
            Index           =   7
            Left            =   7830
            TabIndex        =   12
            Top             =   60
            Width           =   1065
         End
         Begin VB.OptionButton optTamgu 
            BackColor       =   &H00FFFF00&
            Caption         =   "Option1"
            Height          =   240
            Index           =   8
            Left            =   8940
            TabIndex        =   13
            Top             =   60
            Width           =   1065
         End
         Begin VB.OptionButton optTamgu 
            BackColor       =   &H0000C000&
            Caption         =   "Option1"
            Height          =   240
            Index           =   9
            Left            =   10050
            TabIndex        =   14
            Top             =   60
            Width           =   1065
         End
         Begin VB.OptionButton optTamgu 
            BackColor       =   &H000000FF&
            Caption         =   "Option1"
            Height          =   240
            Index           =   10
            Left            =   11160
            TabIndex        =   15
            Top             =   60
            Width           =   1065
         End
         Begin VB.OptionButton optTamgu 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Option1"
            Height          =   240
            Index           =   11
            Left            =   12270
            TabIndex        =   16
            Top             =   60
            Width           =   1065
         End
      End
      Begin FPSpread.vaSpread sprGwamok 
         Height          =   2325
         Left            =   30
         TabIndex        =   19
         Top             =   510
         Width           =   12045
         _Version        =   393216
         _ExtentX        =   21246
         _ExtentY        =   4101
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
         SpreadDesigner  =   "TMR020.frx":CDB0
      End
   End
End
Attribute VB_Name = "TMR020"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################################################
'   �� �� ��  �� : �뼺�п� ���л���, �ݹ��� & �ð�ǥ ���α׷�
'   ����ý��۸� :
'   ��   ��   �� : TRM020
'   �� ��  �� �� :
'
'   ��   ��   �� : 2007/11/06
'   ��   ��   �� : ���ϱ�
' --------------------------------------------------------------------------------------------------------------
' --------------------------------------------------------------------------------------------------------------
'                 ��     ��     ��     ��
' --------------------------------------------------------------------------------------------------------------
'   1. ������ :
'   2. ��  �� :
'################################################################################################################

Option Explicit

Private Const nRowHeight = 14


Private Type tLsnSetting
    LSNCD       As String       '< ��
    LSN_FM      As String       '< ����/�̵���
    SUBJCD      As String       '< �����ڵ�
End Type
Private uLsnSetting()       As tLsnSetting



Private Sub Form_Unload(Cancel As Integer)
    Unload TMR021
    Unload TMR022
    
End Sub

Private Sub Form_Load()
    
    Me.Move 0, 0, 15700, 10670
    
    Me.Tag = "LOAD"
        With sprLsn
            .ShadowColor = basModule.ShadowColor1
            .ShadowDark = basModule.ShadowDark1
            .ShadowText = basModule.ShadowText1
            .GridColor = basModule.GridColor1
            .GrayAreaBackColor = basModule.GrayAreaBackColor1
            
            .MaxRows = 0
            
            chkAll.BackColor = basModule.ShadowColor1
        End With
        
        With sprSTD
            .ShadowColor = basModule.ShadowColor2
            .ShadowDark = basModule.ShadowDark2
            .ShadowText = basModule.ShadowText2
            .GridColor = basModule.GridColor2
            .GrayAreaBackColor = basModule.GrayAreaBackColor2
            
            .MaxRows = 0
        End With
        
        With sprGwamok
            .ShadowColor = basModule.ShadowColor2
            .ShadowDark = basModule.ShadowDark2
            .ShadowText = basModule.ShadowText2
            .GridColor = basModule.GridColor2
            .GrayAreaBackColor = basModule.GrayAreaBackColor2
            
            .MaxRows = 4
            .MaxCols = 0
            
            .Col = SpreadHeader:        .ColWidth(.Col) = 14
                .Row = 1:       .Text = "����1":    .RowHeight(.Row) = nRowHeight + 5
                .Row = 2:       .Text = "����2":    .RowHeight(.Row) = nRowHeight + 5
                .Row = 3:       .Text = "�̵�1":    .RowHeight(.Row) = nRowHeight + 5
                .Row = 4:       .Text = "�̵�2":    .RowHeight(.Row) = nRowHeight + 5
        End With
        
        With cboKaeyol
            .Clear
            .AddItem "�ι�" & Space(30) & "01"
            .AddItem "�ڿ�" & Space(30) & "02"
            
            .ListIndex = 0
        End With
        
        With cboLsnType
            .Clear
            .AddItem "��ü" & Space(30) & "ALL"
            .AddItem "A type" & Space(30) & "A"
            .AddItem "B type" & Space(30) & "B"
            .AddItem "C type" & Space(30) & "C"
            
            .ListIndex = 1
        End With
        
        With sprCopyLsn
            .ShadowColor = basModule.ShadowColor2
            .ShadowDark = basModule.ShadowDark2
            .ShadowText = basModule.ShadowText2
            .GridColor = basModule.GridColor2
            .GrayAreaBackColor = basModule.GrayAreaBackColor2
            
            .MaxCols = 0
            .MaxRows = 0
        End With
        
        With sprBaseLsn
            .ShadowColor = basModule.ShadowColor1
            .ShadowDark = basModule.ShadowDark1
            .ShadowText = basModule.ShadowText1
            .GridColor = basModule.GridColor1
            .GrayAreaBackColor = basModule.GrayAreaBackColor1
            
            .MaxCols = 0
            .MaxRows = 0
        End With
        
        With cboLsnin
            .Clear
            .AddItem "�ݵ��" & Space(30) & "[T]IN"
            
            .ListIndex = 0
        End With
        
        
        '## �ʱ�ȭ
        sprCopyLsn.MaxRows = 0
        chkOKNot.Value = 0
        fpCL_Close.Text = Format(Now, "yyyy-mm")
        
    Me.Tag = ""
    
End Sub

Private Sub cboLsnType_Click()
    txtLsnType.Text = Trim(Right(cboLsnType.Text, 30))
    
End Sub

Private Sub cboKaeyol_Click()

    txtKaeyol.Text = Trim(Right(cboKaeyol.Text, 30))
    Select Case Trim(Right(cboKaeyol.Text, 30))
        Case "01", "03"         '<< �ι�
            With sprLsn
                .Row = SpreadHeader:        .RowHeight(.Row) = nRowHeight
                .MaxCols = 16
                
                .Col = 1:           .Text = "��":           .ColWidth(.Col) = 8 ':        .ColHidden = True
                .Col = .Col + 1:    .Text = "�ݸ�":         .ColWidth(.Col) = 10
                .Col = .Col + 1:    .Text = "�۾�����":     .ColWidth(.Col) = 9
                .Col = .Col + 1:    .Text = "�����":       .ColWidth(.Col) = 7
                .Col = .Col + 1:    .Text = "���ο�":       .ColWidth(.Col) = 7
                
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
                
                .MaxRows = 0
            End With
            
            optTamgu(0).Caption = "����/����"
            optTamgu(1).Caption = "����"
            optTamgu(2).Caption = "����"
            optTamgu(3).Caption = "����"
            optTamgu(4).Caption = "�ѱ�"
            optTamgu(5).Caption = "�����"
            optTamgu(6).Caption = "����"
            optTamgu(7).Caption = "����"
            optTamgu(8).Caption = "��ġ"
            optTamgu(9).Caption = "�繮":       optTamgu(9).Visible = True
            optTamgu(10).Caption = "����":      optTamgu(10).Visible = True
            optTamgu(11).Caption = "����":      optTamgu(11).Visible = True
                        
            optTamgu(0).Value = True            '�⺻����
            
        Case "02"       '<< �ڿ�
            With sprLsn
                .Row = SpreadHeader:        .RowHeight(.Row) = nRowHeight
                .MaxCols = 13
                
                .Col = 1:           .Text = "��":           .ColWidth(.Col) = 8 ':        .ColHidden = True
                .Col = .Col + 1:    .Text = "�ݸ�":         .ColWidth(.Col) = 10
                .Col = .Col + 1:    .Text = "�۾�����":     .ColWidth(.Col) = 9
                .Col = .Col + 1:    .Text = "�����":       .ColWidth(.Col) = 7
                .Col = .Col + 1:    .Text = "���ο�":       .ColWidth(.Col) = 7
                
                .Col = .Col + 1:    .Text = "��1":          .ColWidth(.Col) = 5
                .Col = .Col + 1:    .Text = "ȭ1":          .ColWidth(.Col) = 5
                .Col = .Col + 1:    .Text = "��1":          .ColWidth(.Col) = 5
                .Col = .Col + 1:    .Text = "��1":          .ColWidth(.Col) = 5
                .Col = .Col + 1:    .Text = "��2":          .ColWidth(.Col) = 5
                .Col = .Col + 1:    .Text = "ȭ2":          .ColWidth(.Col) = 5
                .Col = .Col + 1:    .Text = "��2":          .ColWidth(.Col) = 5
                .Col = .Col + 1:    .Text = "��2":          .ColWidth(.Col) = 5
                
                .MaxRows = 0
            End With
            
            optTamgu(0).Caption = "����/����"
            optTamgu(1).Caption = "��1"
            optTamgu(2).Caption = "ȭ1"
            optTamgu(3).Caption = "��1"
            optTamgu(4).Caption = "��1"
            optTamgu(5).Caption = "��2"
            optTamgu(6).Caption = "ȭ2"
            optTamgu(7).Caption = "��2"
            optTamgu(8).Caption = "��2"
            
            optTamgu(9).Caption = "":       optTamgu(9).Visible = False
            optTamgu(10).Caption = "":      optTamgu(10).Visible = False
            optTamgu(11).Caption = "":      optTamgu(11).Visible = False

            optTamgu(0).Value = True            '�⺻����
            
    End Select
        
End Sub


'>> �ݺ� ���� ��û���� ��ȸ
Private Sub cmdFind_STD_Subj_Click()
    Dim nRow        As Long
    Dim nCol        As Long
    
    Dim sLsnCD      As String
    Dim sLsnNM      As String
    
    On Error GoTo ErrStmt
    
    sprLsn.MaxRows = 0
    chkAll.Value = 0
    sprGwamok.MaxCols = 0
    
    txtKaeyol.Text = Trim(Right(cboKaeyol.Text, 30))        '<< �迭
    txtLsnType.Text = Trim(Right(cboLsnType.Text, 30))      '<< �� ����
    
    Select Case Find_Lsn_To_STD_TOT                 '<< �ݺ� �����û���� �հ��ο�
        Case 0
            MsgBox "��ȸ�� �Ϸ��Ͽ����ϴ�.", vbInformation + vbOKOnly, "�ݺ� ������û���� ��ȸ"
        Case Is > 0
            
            
            '�ݺ� ��û�ο� ��ϵ� �ο�
            '-----------------------------------------------
                Call View_SaveBase_LsnCopySpread

            
            If txtLsnType.Text = "ALL" Then
                MsgBox "�ݺ� ��û������ ��ȸ�Ͽ�����," & vbCrLf & _
                       "�������� ���Ͻø� �����¸� �����Ͻʽÿ�.", vbExclamation + vbOKOnly, "�̵����� �ð�ǥ ��ȸ"
                Exit Sub
            End If
            
            
            '## ��ϵ� ������ �ִ��� ��ȸ
            If Last_Save_Chk_Gwamok = False Then            '<< ���°��
                sprGwamok.MaxCols = sprLsn.MaxRows
                sprGwamok.ColHeaderRows = 2
                
                For nRow = 1 To sprLsn.MaxRows Step 1
                    sprLsn.Row = nRow
                    sprLsn.Col = 1:         sLsnCD = Trim(sprLsn.Text)
                    sprLsn.Col = 2:         sLsnNM = Trim(sprLsn.Text)
                    
                    '>> sprGwamok Header ����.
                    sprGwamok.Col = nRow            '<< sprLsn ���� sprGwamok�������� �ٲ�.
                        sprGwamok.Row = SpreadHeader:           sprGwamok.Text = sLsnCD:        sprGwamok.RowHeight(sprGwamok.Row) = nRowHeight:        sprGwamok.RowHidden = True
                        sprGwamok.Row = SpreadHeader + 1:       sprGwamok.Text = sLsnNM:        sprGwamok.RowHeight(sprGwamok.Row) = nRowHeight + 2
                Next nRow
                        
                MsgBox "��ȸ�Ͽ����ϴ�.", vbInformation + vbOKOnly, "�̵����� �ð�ǥ ��ȸ"
            
            Else            '<< �ִ°��
            
                Call cmdFindLsn_Click                       '<< ���񳻿� ��ȸ
                
            End If
    End Select
    
    With sprLsn
        cboLsnin.Clear
        cboLsnin.AddItem "�ݵ��" & Space(30) & "[T]IN"
        
        If .MaxRows > 0 Then
            For nRow = 1 To .MaxRows Step 1
                .Row = nRow
                .Col = 2:       sLsnNM = Trim(.Text)
                .Col = 1:       sLsnCD = Trim(.Text)
                
                cboLsnin.AddItem sLsnNM & Space(30) & "[T]" & sLsnCD
                
            Next nRow
        End If
        
        cboLsnin.AddItem "�ݻ���" & Space(30) & "[T]OUT"
        cboLsnin.ListIndex = 0
    End With
    
    
    Exit Sub
ErrStmt:
    MsgBox "��ȸ�� ������ �߻��Ͽ����ϴ�." & vbCrLf & _
           Trim(CStr(Err.Number)) & ":" & Trim(Err.Description), vbCritical + vbOKOnly, "�̵����� �ð�ǥ ��ȸ"
    On Error GoTo 0
End Sub

'-----------------------------------------------
'�ݺ� ��û�ο� ��ϵ� �ο�
'-----------------------------------------------
Private Sub View_SaveBase_LsnCopySpread()

    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    
    Dim nLength     As Long
    Dim sStr        As String
    Dim ni          As Integer
    Dim nRec        As Long
    
    Dim sTmp        As String
    Dim nTmp        As Double
    
    Dim nCol        As Integer
    Dim siTem       As String
    
    Dim nRet        As Long
    Dim nCol_Lsn    As Long
    Dim sGwamok     As String
   
    On Error GoTo ErrStmt
    
    
    
    sprBaseLsn.MaxRows = 0
    If sprLsn.MaxCols = 0 Then Exit Sub
        sprBaseLsn.MaxCols = sprLsn.MaxCols - 2         '< �۾�����/ ó���ο��� ����
        sprBaseLsn.Col = 1

    For nCol_Lsn = 1 To sprLsn.MaxCols Step 1
        sprLsn.Row = SpreadHeader
        sprLsn.Col = nCol_Lsn:      sTmp = Trim(sprLsn.Text)

        sprBaseLsn.Row = SpreadHeader

            Select Case nCol_Lsn
                Case 1, 2, 5

                    sprBaseLsn.Text = sTmp:     sprBaseLsn.ColWidth(sprBaseLsn.Col) = 7
                    sprBaseLsn.Col = sprBaseLsn.Col + 1
                    
                Case 6 To sprLsn.MaxCols
                    Select Case sTmp
                        Case "����"
                            sGwamok = "01"
                        Case "����"
                            sGwamok = "02"
                        Case "����"
                            sGwamok = "03"
                        Case "�ѱ�"
                            sGwamok = "04"
                        Case "�����"
                            sGwamok = "05"
                        Case "����"
                            sGwamok = "06"
                        Case "����"
                            sGwamok = "07"
                        Case "��ġ"
                            sGwamok = "08"
                        Case "�繮"
                            sGwamok = "09"
                        Case "����"
                            sGwamok = "10"
                        Case "����"
                            sGwamok = "11"
                        Case "��1"
                            sGwamok = "51"
                        Case "ȭ1"
                            sGwamok = "52"
                        Case "��1"
                            sGwamok = "53"
                        Case "��1"
                            sGwamok = "54"
                        Case "��2"
                            sGwamok = "55"
                        Case "ȭ2"
                            sGwamok = "56"
                        Case "��2"
                            sGwamok = "57"
                        Case "��2"
                            sGwamok = "58"
                    End Select
                    sprBaseLsn.Text = sGwamok:  sprBaseLsn.ColWidth(sprBaseLsn.Col) = 5
                    
                    sprBaseLsn.Col = sprBaseLsn.Col + 1
                    
            End Select
    Next nCol_Lsn
    
    
    
    
    nRet = 0
    
    sStr = ""
    sStr = sStr & "  SELECT LSNCD, LSNNM, "
    sStr = sStr & "         S_LSN,"
    sStr = sStr & "         SEL1 ,"
    sStr = sStr & "         SEL2 ,"
    sStr = sStr & "         SEL3 ,"
    sStr = sStr & "         SEL4 ,"
    sStr = sStr & "         SEL5 ,"
    sStr = sStr & "         SEL6 ,"
    sStr = sStr & "         SEL7 ,"
    sStr = sStr & "         SEL8 ,"
    sStr = sStr & "         SEL9 ,"
    sStr = sStr & "         SEL10,"
    sStr = sStr & "         SEL11,"
    sStr = sStr & "         KAEYOL"
    
    sStr = sStr & "    FROM (SELECT LSNCD,"
    sStr = sStr & "                 GET_LSNNM(LSNCD) AS LSNNM,"
    
    sStr = sStr & "                 0 AS S_LSN,"
    sStr = sStr & "                 SUM(SEL1 ) AS SEL1 ,"
    sStr = sStr & "                 SUM(SEL2 ) AS SEL2 ,"
    sStr = sStr & "                 SUM(SEL3 ) AS SEL3 ,"
    sStr = sStr & "                 SUM(SEL4 ) AS SEL4 ,"
    sStr = sStr & "                 SUM(SEL5 ) AS SEL5 ,"
    sStr = sStr & "                 SUM(SEL6 ) AS SEL6 ,"
    sStr = sStr & "                 SUM(SEL7 ) AS SEL7 ,"
    sStr = sStr & "                 SUM(SEL8 ) AS SEL8 ,"
    sStr = sStr & "                 SUM(SEL9 ) AS SEL9 ,"
    sStr = sStr & "                 SUM(SEL10) AS SEL10,"
    sStr = sStr & "                 SUM(SEL11) AS SEL11,"
    sStr = sStr & "                 MAX(GAEYUL_CD) AS KAEYOL"
    
    sStr = sStr & "           FROM (SELECT LSNCD, "
    sStr = sStr & "                        GAEYUL_CD,"
    
    sStr = sStr & "                        SEL1 ,"
    sStr = sStr & "                        SEL2 ,"
    sStr = sStr & "                        SEL3 ,"
    sStr = sStr & "                        SEL4 ,"
    sStr = sStr & "                        SEL5 ,"
    sStr = sStr & "                        SEL6 ,"
    sStr = sStr & "                        SEL7 ,"
    sStr = sStr & "                        SEL8 ,"
    sStr = sStr & "                        SEL9 ,"
    sStr = sStr & "                        SEL10,"
    sStr = sStr & "                        SEL11,"
    sStr = sStr & "                        CL_CLOSE "
    
    sStr = sStr & "                  FROM (SELECT SEL_CLASS AS LSNCD,"
    sStr = sStr & "                               CASE WHEN SEL1 > ' ' THEN"
    sStr = sStr & "                                  '01'"
    sStr = sStr & "                               ELSE CASE WHEN SEL3 > ' ' THEN"
    sStr = sStr & "                                  '02'"
    sStr = sStr & "                               END END GAEYUL_CD,"
    
    sStr = sStr & "                        /* ��Ž, ��Ž �и� */"
    sStr = sStr & "                               CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'01|') > 0 THEN          /* ��Ž-���� */"
    sStr = sStr & "                                  1"
    sStr = sStr & "                               ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'51|') > 0 THEN     /* ��Ž-����1 */"
    sStr = sStr & "                                  1"
    sStr = sStr & "                               ELSE"
    sStr = sStr & "                                  0"
    sStr = sStr & "                               END END SEL1,"
    sStr = sStr & "                               CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'02|') > 0 THEN          /* ��Ž-���� */"
    sStr = sStr & "                                  1"
    sStr = sStr & "                               ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'52|') > 0 THEN     /* ��Ž-ȭ��1 */"
    sStr = sStr & "                                  1"
    sStr = sStr & "                               ELSE"
    sStr = sStr & "                                  0"
    sStr = sStr & "                               END END SEL2,"
    sStr = sStr & "                               CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'03|') > 0 THEN          /* ��Ž-���� */"
    sStr = sStr & "                                  1"
    sStr = sStr & "                               ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'53|') > 0 THEN     /* ��Ž-����1 */"
    sStr = sStr & "                                  1"
    sStr = sStr & "                               ELSE"
    sStr = sStr & "                                  0"
    sStr = sStr & "                               END END SEL3,"
    sStr = sStr & "                               CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'04|') > 0 THEN          /* ��Ž-�ѱ������� */"
    sStr = sStr & "                                  1"
    sStr = sStr & "                               ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'54|') > 0 THEN     /* ��Ž-��������1 */"
    sStr = sStr & "                                  1"
    sStr = sStr & "                               ELSE"
    sStr = sStr & "                                  0"
    sStr = sStr & "                               END END SEL4,"
    sStr = sStr & "                               CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'05|') > 0 THEN          /* ��Ž-����� */"
    sStr = sStr & "                                  1"
    sStr = sStr & "                               ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'55|') > 0 THEN     /* ��Ž-����2 */"
    sStr = sStr & "                                  1"
    sStr = sStr & "                               ELSE"
    sStr = sStr & "                                  0"
    sStr = sStr & "                               END END SEL5,"
    sStr = sStr & "                               CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'06|') > 0 THEN          /* ��Ž-�������� */"
    sStr = sStr & "                                  1"
    sStr = sStr & "                               ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'56|') > 0 THEN     /* ��Ž-ȭ��2 */"
    sStr = sStr & "                                  1"
    sStr = sStr & "                               ELSE"
    sStr = sStr & "                                  0"
    sStr = sStr & "                               END END SEL6,"
    sStr = sStr & "                               CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'07|') > 0 THEN          /* ��Ž-�ѱ����� */"
    sStr = sStr & "                                   1"
    sStr = sStr & "                               ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'57|') > 0 THEN     /* ��Ž-����2 */"
    sStr = sStr & "                                   1"
    sStr = sStr & "                               ELSE"
    sStr = sStr & "                                   0"
    sStr = sStr & "                               END END SEL7,"
    sStr = sStr & "                               CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'08|') > 0 THEN          /* ��Ž-��ġ */"
    sStr = sStr & "                                   1"
    sStr = sStr & "                               ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'58|') > 0 THEN     /* ��Ž-��������2 */"
    sStr = sStr & "                                   1"
    sStr = sStr & "                               ELSE"
    sStr = sStr & "                                   0"
    sStr = sStr & "                               END END SEL8,"
    sStr = sStr & "                               CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'09|') > 0 THEN          /* ��Ž-��ȸ��ȭ */"
    sStr = sStr & "                                   1"
    sStr = sStr & "                               ELSE"
    sStr = sStr & "                                   0"
    sStr = sStr & "                               END SEL9,"
    sStr = sStr & "                               CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'10|') > 0 THEN          /* ��Ž-������ȸ */"
    sStr = sStr & "                                   1"
    sStr = sStr & "                               ELSE"
    sStr = sStr & "                                   0"
    sStr = sStr & "                               END SEL10,"
    sStr = sStr & "                               CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'11|') > 0 THEN          /* ��Ž-�������� */"
    sStr = sStr & "                                   1"
    sStr = sStr & "                               ELSE"
    sStr = sStr & "                                   0"
    sStr = sStr & "                               END SEL11, "
    sStr = sStr & "                               CL_CLOSE "
    
    sStr = sStr & "                          FROM CLTTL01TB"
    sStr = sStr & "                         WHERE ACID   = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "                           AND CL_CLOSE IS NULL "
    
    sStr = sStr & "                        )"
    sStr = sStr & "                    WHERE GAEYUL_CD = '" & Trim(txtKaeyol.Text) & "'"
    sStr = sStr & "                   )"
    sStr = sStr & "              GROUP BY LSNCD"
    sStr = sStr & "              HAVING LSNCD"
    sStr = sStr & "                  IN (SELECT LSNCD"
    sStr = sStr & "                        FROM SDLSN01TB"
    sStr = sStr & "                       WHERE ACID    = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "                         AND KAEYOL  = '" & Trim(txtKaeyol.Text) & "'"
    If Trim(Right(cboLsnType.Text, 30)) <> "ALL" Then
        sStr = sStr & "                     AND LSNTYPE = '" & Trim(txtLsnType.Text) & "'"
    End If
    sStr = sStr & "                     )"
    sStr = sStr & "           )"
    sStr = sStr & "      ORDER BY LSNNM "
    
    
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
    
'    '>> �п�
'        sTmp = Trim(basModule.SchCD)
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
'    '>> �迭
'        sTmp = Trim(Right(cboKaeyol.Text, 30))
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
'    '>> �п�
'        sTmp = Trim(basModule.SchCD)
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
'    '>> �迭
'        sTmp = Trim(Right(cboKaeyol.Text, 30))
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
'    '>> �� ����
'        If Trim(Right(cboLsnType.Text, 30)) <> "ALL" Then
'            sTmp = Trim(Right(cboKaeyol.Text, 30))
'            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'                Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
'        End If
        
        
    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic���·� ���ԵǸ� record count�� �� �� ����.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop
    
    With DBRec
        If .RecordCount > 0 Then
            .MoveFirst
            
            
            For nRec = 1 To .RecordCount Step 1
            
                nRet = nRet + 1
                
                sprBaseLsn.MaxRows = sprBaseLsn.MaxRows + 1
                sprBaseLsn.Row = sprBaseLsn.MaxRows
                
                sprBaseLsn.Col = 1
                    sTmp = " ": If IsNull(.Fields("LSNCD")) = False Then sTmp = Trim(.Fields("LSNCD"))
                        Call basFunction.Set_SprType_Text(sprBaseLsn, "CENTER", "LEFT", LenB(sTmp), sTmp)
                sprBaseLsn.Col = sprBaseLsn.Col + 1
                    sTmp = " ": If IsNull(.Fields("LSNNM")) = False Then sTmp = Trim(.Fields("LSNNM"))
                        Call basFunction.Set_SprType_Text(sprBaseLsn, "CENTER", "LEFT", LenB(sTmp), sTmp)
                    
                '## ���ο�
                sprBaseLsn.Col = sprBaseLsn.Col + 1:    nTmp = 0
                    If IsNull(.Fields("S_LSN")) = False Then
                        nTmp = CDbl(.Fields("S_LSN"))
                    End If
                    Call basFunction.Set_SprType_Numeric(sprBaseLsn, 0, 0, 99999, ",", nTmp)
                    
                    
                '<< �ι��ڿ� ���� : 8 ���� >>
                For nCol = 1 To 8 Step 1
                    sprBaseLsn.Col = sprBaseLsn.Col + 1:    nTmp = 0
                    siTem = "SEL" & Trim(CStr(nCol))
                    
                    If IsNull(.Fields(siTem)) = False Then nTmp = CDbl(.Fields(siTem))
                    Call basFunction.Set_SprType_Numeric(sprBaseLsn, 0, 0, 99999, "", nTmp)
                Next nCol
                
                
                Select Case Trim(.Fields("KAEYOL"))
                    Case "01", "03"
                        '��Ž�� 9~11
                        For nCol = 9 To 11 Step 1
                            sprBaseLsn.Col = sprBaseLsn.Col + 1:    nTmp = 0
                            siTem = "SEL" & Trim(CStr(nCol))
                            
                            If IsNull(.Fields(siTem)) = False Then nTmp = CDbl(.Fields(siTem))
                            Call basFunction.Set_SprType_Numeric(sprBaseLsn, 0, 0, 99999, "", nTmp)
                            
                        Next nCol
                        
                    Case "02"
                        '��Ž�� SKIP
                End Select
                
                
                .MoveNext       '<< �����׸�
                
            Next nRec
            
            sprBaseLsn.Row = 1:       sprBaseLsn.Row2 = sprBaseLsn.MaxRows
            sprBaseLsn.Col = 1:       sprBaseLsn.Col2 = sprBaseLsn.MaxCols
            sprBaseLsn.BlockMode = True
                sprBaseLsn.BackColor = basModule.WhiteColor
                sprBaseLsn.BackColorStyle = BackColorStyleUnderGrid
            sprBaseLsn.BlockMode = False

        '>> spread lock
            sprBaseLsn.Row = 1:       sprBaseLsn.Row2 = sprBaseLsn.MaxRows
            sprBaseLsn.Col = 1:       sprBaseLsn.Col2 = sprBaseLsn.MaxCols
            sprBaseLsn.BlockMode = True
                sprBaseLsn.Lock = True
                sprBaseLsn.Protect = True
            sprBaseLsn.BlockMode = False
            
        End If
    End With
    
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    Exit Sub
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    On Error GoTo 0
    MsgBox "�� �۾��� ���� ��ȸ�� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "�۾��� ���� ��ȸ"
    
End Sub


'## ���� ���񳻿� ��ϵ� ������ �ִ��� üũ��.
Private Function Last_Save_Chk_Gwamok() As Boolean
    
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    
    Dim nLength     As Long
    Dim sStr        As String
    Dim ni          As Integer
    Dim nRec        As Long
    
    Dim sTmp        As String
    Dim nTmp        As Double
    
    Dim bChk        As Boolean
      
    bChk = False
    On Error GoTo ErrStmt
    
    sStr = ""
    sStr = sStr & "  SELECT COUNT(SUBJCD) AS SUBJCD "
    sStr = sStr & "    From SDTRX20TB"
    sStr = sStr & "   WHERE ACID    = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "     AND KAEYOL  = '" & Trim(txtKaeyol.Text) & "'"
    sStr = sStr & "     AND LSNTYPE = '" & Trim(txtLsnType.Text) & "'"
    
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
    
'    '>> �п�
'        sTmp = Trim(basModule.SchCD)
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
'    '>> �迭
'        sTmp = Trim(txtKaeyol.Text)
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
'    '>> �� ����
'        sTmp = Trim(txtLsnType.Text)
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
       
        
    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic���·� ���ԵǸ� record count�� �� �� ����.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop
    
    With DBRec
        If .RecordCount > 0 Then
            .MoveFirst
            
            If IsNull(.Fields("SUBJCD")) = False Then
                If CLng(.Fields("SUBJCD")) > 0 Then
                    bChk = True
                End If
            End If
            
            .MoveNext       '<< �����׸�
        End If
    End With
    
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    Last_Save_Chk_Gwamok = bChk
    
    Exit Function
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    On Error GoTo 0
    MsgBox "���񳻿� ��ȸ�� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "�̵�����ǥ ��ȸ"
    
    Last_Save_Chk_Gwamok = bChk
    
End Function









'## �ݺ� �����û���� �հ��ο�
Private Function Find_Lsn_To_STD_TOT() As Long

    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    
    Dim nLength     As Long
    Dim sStr        As String
    Dim ni          As Integer
    Dim nRec        As Long
    
    Dim sTmp        As String
    Dim nTmp        As Double
    
    Dim nCol        As Integer
    Dim siTem       As String
    
    Dim nRet        As Long
    
    On Error GoTo ErrStmt
    
    nRet = 0
    
    sStr = ""
    sStr = sStr & "  SELECT LSNCD, LSNNM, INWON_STAT, "
    sStr = sStr & "         S_LSN,"
    sStr = sStr & "         SEL1 ,"
    sStr = sStr & "         SEL2 ,"
    sStr = sStr & "         SEL3 ,"
    sStr = sStr & "         SEL4 ,"
    sStr = sStr & "         SEL5 ,"
    sStr = sStr & "         SEL6 ,"
    sStr = sStr & "         SEL7 ,"
    sStr = sStr & "         SEL8 ,"
    sStr = sStr & "         SEL9 ,"
    sStr = sStr & "         SEL10,"
    sStr = sStr & "         SEL11,"
    sStr = sStr & "         KAEYOL"
    
    sStr = sStr & "    FROM (SELECT LSNCD,"
    sStr = sStr & "                 GET_LSNNM(LSNCD) AS LSNNM,"
    
    sStr = sStr & "                 COUNT(CL_CLOSE) AS INWON_STAT,                      /* �۾��Ϸ� �� �л� */"
    
    sStr = sStr & "                 COUNT(LSNCD) AS S_LSN,"
    sStr = sStr & "                 SUM(SEL1 ) AS SEL1 ,"
    sStr = sStr & "                 SUM(SEL2 ) AS SEL2 ,"
    sStr = sStr & "                 SUM(SEL3 ) AS SEL3 ,"
    sStr = sStr & "                 SUM(SEL4 ) AS SEL4 ,"
    sStr = sStr & "                 SUM(SEL5 ) AS SEL5 ,"
    sStr = sStr & "                 SUM(SEL6 ) AS SEL6 ,"
    sStr = sStr & "                 SUM(SEL7 ) AS SEL7 ,"
    sStr = sStr & "                 SUM(SEL8 ) AS SEL8 ,"
    sStr = sStr & "                 SUM(SEL9 ) AS SEL9 ,"
    sStr = sStr & "                 SUM(SEL10) AS SEL10,"
    sStr = sStr & "                 SUM(SEL11) AS SEL11,"
    sStr = sStr & "                 MAX(GAEYUL_CD) AS KAEYOL"
    
    sStr = sStr & "           FROM (SELECT LSNCD, "
    sStr = sStr & "                        GAEYUL_CD,"
    
    sStr = sStr & "                        SEL1 ,"
    sStr = sStr & "                        SEL2 ,"
    sStr = sStr & "                        SEL3 ,"
    sStr = sStr & "                        SEL4 ,"
    sStr = sStr & "                        SEL5 ,"
    sStr = sStr & "                        SEL6 ,"
    sStr = sStr & "                        SEL7 ,"
    sStr = sStr & "                        SEL8 ,"
    sStr = sStr & "                        SEL9 ,"
    sStr = sStr & "                        SEL10,"
    sStr = sStr & "                        SEL11,"
    sStr = sStr & "                        CL_CLOSE "
    
    sStr = sStr & "                  FROM (SELECT SEL_CLASS AS LSNCD,"
    sStr = sStr & "                               CASE WHEN SEL1 > ' ' THEN"
    sStr = sStr & "                                  '01'"
    sStr = sStr & "                               ELSE CASE WHEN SEL3 > ' ' THEN"
    sStr = sStr & "                                  '02'"
    sStr = sStr & "                               END END GAEYUL_CD,"
    
    sStr = sStr & "                        /* ��Ž, ��Ž �и� */"
    sStr = sStr & "                               CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'01|') > 0 THEN          /* ��Ž-���� */"
    sStr = sStr & "                                  1"
    sStr = sStr & "                               ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'51|') > 0 THEN     /* ��Ž-����1 */"
    sStr = sStr & "                                  1"
    sStr = sStr & "                               ELSE"
    sStr = sStr & "                                  0"
    sStr = sStr & "                               END END SEL1,"
    sStr = sStr & "                               CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'02|') > 0 THEN          /* ��Ž-���� */"
    sStr = sStr & "                                  1"
    sStr = sStr & "                               ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'52|') > 0 THEN     /* ��Ž-ȭ��1 */"
    sStr = sStr & "                                  1"
    sStr = sStr & "                               ELSE"
    sStr = sStr & "                                  0"
    sStr = sStr & "                               END END SEL2,"
    sStr = sStr & "                               CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'03|') > 0 THEN          /* ��Ž-���� */"
    sStr = sStr & "                                  1"
    sStr = sStr & "                               ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'53|') > 0 THEN     /* ��Ž-����1 */"
    sStr = sStr & "                                  1"
    sStr = sStr & "                               ELSE"
    sStr = sStr & "                                  0"
    sStr = sStr & "                               END END SEL3,"
    sStr = sStr & "                               CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'04|') > 0 THEN          /* ��Ž-�ѱ������� */"
    sStr = sStr & "                                  1"
    sStr = sStr & "                               ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'54|') > 0 THEN     /* ��Ž-��������1 */"
    sStr = sStr & "                                  1"
    sStr = sStr & "                               ELSE"
    sStr = sStr & "                                  0"
    sStr = sStr & "                               END END SEL4,"
    sStr = sStr & "                               CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'05|') > 0 THEN          /* ��Ž-����� */"
    sStr = sStr & "                                  1"
    sStr = sStr & "                               ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'55|') > 0 THEN     /* ��Ž-����2 */"
    sStr = sStr & "                                  1"
    sStr = sStr & "                               ELSE"
    sStr = sStr & "                                  0"
    sStr = sStr & "                               END END SEL5,"
    sStr = sStr & "                               CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'06|') > 0 THEN          /* ��Ž-�������� */"
    sStr = sStr & "                                  1"
    sStr = sStr & "                               ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'56|') > 0 THEN     /* ��Ž-ȭ��2 */"
    sStr = sStr & "                                  1"
    sStr = sStr & "                               ELSE"
    sStr = sStr & "                                  0"
    sStr = sStr & "                               END END SEL6,"
    sStr = sStr & "                               CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'07|') > 0 THEN          /* ��Ž-�ѱ����� */"
    sStr = sStr & "                                   1"
    sStr = sStr & "                               ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'57|') > 0 THEN     /* ��Ž-����2 */"
    sStr = sStr & "                                   1"
    sStr = sStr & "                               ELSE"
    sStr = sStr & "                                   0"
    sStr = sStr & "                               END END SEL7,"
    sStr = sStr & "                               CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'08|') > 0 THEN          /* ��Ž-��ġ */"
    sStr = sStr & "                                   1"
    sStr = sStr & "                               ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'58|') > 0 THEN     /* ��Ž-��������2 */"
    sStr = sStr & "                                   1"
    sStr = sStr & "                               ELSE"
    sStr = sStr & "                                   0"
    sStr = sStr & "                               END END SEL8,"
    sStr = sStr & "                               CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'09|') > 0 THEN          /* ��Ž-��ȸ��ȭ */"
    sStr = sStr & "                                   1"
    sStr = sStr & "                               ELSE"
    sStr = sStr & "                                   0"
    sStr = sStr & "                               END SEL9,"
    sStr = sStr & "                               CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'10|') > 0 THEN          /* ��Ž-������ȸ */"
    sStr = sStr & "                                   1"
    sStr = sStr & "                               ELSE"
    sStr = sStr & "                                   0"
    sStr = sStr & "                               END SEL10,"
    sStr = sStr & "                               CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'11|') > 0 THEN          /* ��Ž-�������� */"
    sStr = sStr & "                                   1"
    sStr = sStr & "                               ELSE"
    sStr = sStr & "                                   0"
    sStr = sStr & "                               END SEL11, "
    sStr = sStr & "                               CL_CLOSE "
    
    sStr = sStr & "                          FROM CLTTL01TB"
    sStr = sStr & "                         WHERE ACID   = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "                        )"
    sStr = sStr & "                    WHERE GAEYUL_CD = '" & Trim(Right(cboKaeyol.Text, 30)) & "'"
    sStr = sStr & "                   )"
    sStr = sStr & "              GROUP BY LSNCD"
    sStr = sStr & "              HAVING LSNCD"
    sStr = sStr & "                  IN (SELECT LSNCD"
    sStr = sStr & "                        FROM SDLSN01TB"
    sStr = sStr & "                       WHERE ACID    = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "                         AND KAEYOL  = '" & Trim(Right(cboKaeyol.Text, 30)) & "'"
    If Trim(Right(cboLsnType.Text, 30)) <> "ALL" Then
        sStr = sStr & "                     AND LSNTYPE = '" & Trim(Right(cboLsnType.Text, 30)) & "'"
    End If
    sStr = sStr & "                     )"
    sStr = sStr & "           )"
    sStr = sStr & "      ORDER BY LSNNM "
    
    
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
    
'    '>> �п�
'        sTmp = Trim(basModule.SchCD)
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
'    '>> �迭
'        sTmp = Trim(Right(cboKaeyol.Text, 30))
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
'    '>> �п�
'        sTmp = Trim(basModule.SchCD)
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
'    '>> �迭
'        sTmp = Trim(Right(cboKaeyol.Text, 30))
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
'    '>> �� ����
'        If Trim(Right(cboLsnType.Text, 30)) <> "ALL" Then
'            sTmp = Trim(Right(cboKaeyol.Text, 30))
'            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'                Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
'        End If
        
        
    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic���·� ���ԵǸ� record count�� �� �� ����.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop
    
    With DBRec
        If .RecordCount > 0 Then
            .MoveFirst
            
            
            For nRec = 1 To .RecordCount Step 1
            
                nRet = nRet + 1
                
                sprLsn.MaxRows = sprLsn.MaxRows + 1
                sprLsn.Row = sprLsn.MaxRows
                
                sprLsn.Col = 1
                    sTmp = " ": If IsNull(.Fields("LSNCD")) = False Then sTmp = Trim(.Fields("LSNCD"))
                        Call basFunction.Set_SprType_Text(sprLsn, "CENTER", "LEFT", LenB(sTmp), sTmp)
                sprLsn.Col = sprLsn.Col + 1
                    sTmp = " ": If IsNull(.Fields("LSNNM")) = False Then sTmp = Trim(.Fields("LSNNM"))
                        Call basFunction.Set_SprType_Text(sprLsn, "CENTER", "LEFT", LenB(sTmp), sTmp)
                
                sprLsn.Col = sprLsn.Col + 1
                    Call basFunction.Set_SprType_ChkBox(sprLsn):        sprLsn.Value = 0
                
                '## ó���ο�
                sprLsn.Col = sprLsn.Col + 1:    nTmp = 0
                    If IsNull(.Fields("INWON_STAT")) = False Then
                        nTmp = CDbl(.Fields("INWON_STAT"))
                    End If
                    Call basFunction.Set_SprType_Numeric(sprLsn, 0, 0, 99999, ",", nTmp)
                    
                '## ���ο�
                sprLsn.Col = sprLsn.Col + 1:    nTmp = 0
                    If IsNull(.Fields("S_LSN")) = False Then
                        nTmp = CDbl(.Fields("S_LSN"))
                    End If
                    Call basFunction.Set_SprType_Numeric(sprLsn, 0, 0, 99999, ",", nTmp)
                    
                    
                '<< �ι��ڿ� ���� : 8 ���� >>
                For nCol = 1 To 8 Step 1
                    sprLsn.Col = sprLsn.Col + 1:    nTmp = 0
                    siTem = "SEL" & Trim(CStr(nCol))
                    
                    If IsNull(.Fields(siTem)) = False Then nTmp = CDbl(.Fields(siTem))
                    Call basFunction.Set_SprType_Numeric(sprLsn, 0, 0, 99999, "", nTmp)
                Next nCol
                
                
                Select Case Trim(.Fields("KAEYOL"))
                    Case "01", "03"
                        '��Ž�� 9~11
                        For nCol = 9 To 11 Step 1
                            sprLsn.Col = sprLsn.Col + 1:    nTmp = 0
                            siTem = "SEL" & Trim(CStr(nCol))
                            
                            If IsNull(.Fields(siTem)) = False Then nTmp = CDbl(.Fields(siTem))
                            Call basFunction.Set_SprType_Numeric(sprLsn, 0, 0, 99999, "", nTmp)
                            
                        Next nCol
                        
                    Case "02"
                        '��Ž�� SKIP
                End Select
                
                
                .MoveNext       '<< �����׸�
                
            Next nRec
            
            sprLsn.Row = 1:       sprLsn.Row2 = sprLsn.MaxRows
            sprLsn.Col = 1:       sprLsn.Col2 = sprLsn.MaxCols
            sprLsn.BlockMode = True
                sprLsn.BackColor = basModule.WhiteColor
                sprLsn.BackColorStyle = BackColorStyleUnderGrid
            sprLsn.BlockMode = False

            sprLsn.ColsFrozen = 5
            
        '>> spread lock
            sprLsn.Row = 1:       sprLsn.Row2 = sprLsn.MaxRows
            sprLsn.Col = 1:       sprLsn.Col2 = sprLsn.MaxCols
            sprLsn.BlockMode = True
                sprLsn.Lock = True
                sprLsn.Protect = True
            sprLsn.BlockMode = False
            
        End If
    End With
    
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    Find_Lsn_To_STD_TOT = nRet
    Exit Function
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    On Error GoTo 0
    MsgBox "�ݺ� ������û���� ��ȸ�� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "�ݺ� ������û���� ��ȸ"
    
    Find_Lsn_To_STD_TOT = nRet
End Function


Private Sub chkAll_Click()
    Dim nRow        As Long
    
    With sprLsn
        For nRow = 1 To .MaxRows Step 1
            
            .Row = nRow
            .Col = 3
            If chkAll.Value = 1 Then
                .Value = 1
            Else
                .Value = 0
            End If
        Next nRow
    End With
    
End Sub




Private Sub sprLsn_Click(ByVal Col As Long, ByVal Row As Long)
    If Row < 1 Then Exit Sub
    If Col < 1 Then Exit Sub
    
    With sprLsn
        If Trim(.Tag) = "" Then .Tag = "1"
        
        .Row = CLng(.Tag):  .Row2 = .Row
        .Col = 1:           .Col2 = .MaxCols
        .BlockMode = True
            .BackColor = basModule.WhiteColor
            .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
        
        
        .Row = Row:     .Row2 = .Row
        .Col = 1:       .Col2 = .MaxCols
        .BlockMode = True
            .BackColor = basModule.BackColor1
            .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
        
        .Tag = Trim(CStr(Row))
        
        If Col = 3 Then
            .Row = Row
            .Col = Col
            
            If .Value = 1 Then
                .Value = 0
            Else
                .Value = 1
            End If
        End If
        
    End With
End Sub












'## ���ù��� �л���ȸ
Private Sub cmdFind_LSN_in_STD_Click()
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    
    Dim nLength     As Long
    Dim sStr        As String
    Dim ni          As Long
    Dim nRec        As Long
    
    Dim nRow        As Long
    Dim bChk        As Boolean
    Dim sAddSql     As String           ' sql���� ���� : ��
    
    Dim sTmp        As String
    Dim nTmp        As Double
    
    Dim sDiv()      As String
    
    Dim sTam1       As String
    Dim sTam2       As String
    Dim sTam3       As String
    Dim sTam4       As String
    
    Dim sGbn        As String
    
    On Error GoTo ErrStmt
    
    sprSTD.MaxRows = 0
    
    bChk = False
    sAddSql = ""
    
    With sprLsn
        For nRow = 1 To .MaxRows Step 1
            .Row = nRow
            .Col = 3
            If .Value = 1 Then
                If sAddSql > " " Then sAddSql = sAddSql & ", "
                
                .Col = 1
                    sAddSql = sAddSql & "'" & Trim(.Text) & "'"
                
                bChk = True
            End If
        Next nRow
        
        If bChk = False Then
            MsgBox "ó���� ���� �����ϼ���.", vbExclamation + vbOKOnly, "���ù� �л� ��ȸ"
            Exit Sub
        End If
    End With
    
    sStr = ""
    sStr = sStr & "          SELECT SCHNO, EXMID, STDNM, SEL_CLASS, GET_LSNNM(SEL_CLASS) AS CLASSNM, CL_CLOSE, "
    Select Case Trim(Right(cboKaeyol.Text, 30))
        Case "01", "03"
            sStr = sStr & "         SEL1 AS TAMGU, "
        Case "02"
            sStr = sStr & "         SEL3 AS TAMGU, "
    End Select
    sStr = sStr & "                 CL_CLOSE,"
    sStr = sStr & "                 GWA_BAN1, GWA_BAN2, GWA_BAN3, GWA_BAN4, "
    sStr = sStr & "                 GET_LSNNM(GWA_BAN1) AS GWA_BANNM1, "
    sStr = sStr & "                 GET_LSNNM(GWA_BAN2) AS GWA_BANNM2, "
    sStr = sStr & "                 GET_LSNNM(GWA_BAN3) AS GWA_BANNM3, "
    sStr = sStr & "                 GET_LSNNM(GWA_BAN4) AS GWA_BANNM4 "
    sStr = sStr & "            FROM CLTTL01TB"
    sStr = sStr & "           WHERE ACID      = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "             AND SEL_CLASS"
    sStr = sStr & "              IN ( " & sAddSql & " )"
    If chkOKNot.Value = 0 Then
        sStr = sStr & "         AND CL_CLOSE IS NULL "
    End If
    sStr = sStr & "           ORDER BY EXMID, SEL_CLASS, STDNM "
    
    
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
                sprSTD.MaxRows = sprSTD.MaxRows + 1
                sprSTD.Row = sprSTD.MaxRows
                
                sprSTD.Col = 1
                    sTmp = " ": If IsNull(.Fields("SCHNO")) = False Then sTmp = Trim(.Fields("SCHNO"))
                        Call basFunction.Set_SprType_Text(sprSTD, "CENTER", "LEFT", LenB(sTmp), sTmp)
                sprSTD.Col = sprSTD.Col + 1
                    sTmp = " ": If IsNull(.Fields("EXMID")) = False Then sTmp = Trim(.Fields("EXMID"))
                        Call basFunction.Set_SprType_Text(sprSTD, "CENTER", "LEFT", LenB(sTmp), sTmp)
                sprSTD.Col = sprSTD.Col + 1
                    sTmp = " ": If IsNull(.Fields("STDNM")) = False Then sTmp = Trim(.Fields("STDNM"))
                        Call basFunction.Set_SprType_Text(sprSTD, "CENTER", "LEFT", LenB(sTmp), sTmp)
                
                
                sprSTD.SetCellBorder sprSTD.Col, sprSTD.Row, sprSTD.Col, sprSTD.Row, 2, basModule.SectionColor2, CellBorderStyleSolid
                
                
                sprSTD.Col = sprSTD.Col + 1
                    Call basFunction.Set_SprType_ChkBox(sprSTD):    sprSTD.Value = 0
                
                
                sprSTD.SetCellBorder sprSTD.Col, sprSTD.Row, sprSTD.Col, sprSTD.Row, 2, basModule.SectionColor2, CellBorderStyleSolid
                
                
                sprSTD.Col = sprSTD.Col + 1
                    sTmp = " ": If IsNull(.Fields("SEL_CLASS")) = False Then sTmp = Trim(.Fields("SEL_CLASS"))
                        Call basFunction.Set_SprType_Text(sprSTD, "CENTER", "LEFT", LenB(sTmp), sTmp)
                sprSTD.Col = sprSTD.Col + 1
                    sTmp = " ": If IsNull(.Fields("CLASSNM")) = False Then sTmp = Trim(.Fields("CLASSNM"))
                        Call basFunction.Set_SprType_Text(sprSTD, "CENTER", "LEFT", LenB(sTmp), sTmp)
                
                
                If IsNull(.Fields("TAMGU")) = False Then
                    sTmp = .Fields("TAMGU")
                    
                    sTam1 = "": sTam2 = "": sTam3 = "": sTam4 = ""
                        
                    sDiv() = Split(sTmp, "|", -1, vbTextCompare)
                    Select Case UBound(sDiv)
                        Case 0
                            
                        Case 1
                            sprSTD.Col = sprSTD.Col + 1
                                sTmp = sDiv(0)
                                If LenB(sTmp) > 0 Then Call basFunction.Set_SprType_Text(sprSTD, "CENTER", "LEFT", LenB(sTmp), sTmp)
                                sTam1 = sTmp
                            sprSTD.Col = sprSTD.Col + 1
                            sprSTD.Col = sprSTD.Col + 1
                            sprSTD.Col = sprSTD.Col + 1
                            
                        Case 2
                            sprSTD.Col = sprSTD.Col + 1
                                sTmp = sDiv(0)
                                If LenB(sTmp) > 0 Then Call basFunction.Set_SprType_Text(sprSTD, "CENTER", "LEFT", LenB(sTmp), sTmp)
                                sTam1 = sTmp
                            sprSTD.Col = sprSTD.Col + 1
                                sTmp = sDiv(1)
                                If LenB(sTmp) > 0 Then Call basFunction.Set_SprType_Text(sprSTD, "CENTER", "LEFT", LenB(sTmp), sTmp)
                                sTam2 = sTmp
                            sprSTD.Col = sprSTD.Col + 1
                            sprSTD.Col = sprSTD.Col + 1
                            
                        Case 3
                            sprSTD.Col = sprSTD.Col + 1
                                sTmp = sDiv(0)
                                If LenB(sTmp) > 0 Then Call basFunction.Set_SprType_Text(sprSTD, "CENTER", "LEFT", LenB(sTmp), sTmp)
                                sTam1 = sTmp
                            sprSTD.Col = sprSTD.Col + 1
                                sTmp = sDiv(1)
                                If LenB(sTmp) > 0 Then Call basFunction.Set_SprType_Text(sprSTD, "CENTER", "LEFT", LenB(sTmp), sTmp)
                                sTam2 = sTmp
                            sprSTD.Col = sprSTD.Col + 1
                                sTmp = sDiv(2)
                                If LenB(sTmp) > 0 Then Call basFunction.Set_SprType_Text(sprSTD, "CENTER", "LEFT", LenB(sTmp), sTmp)
                                sTam3 = sTmp
                            sprSTD.Col = sprSTD.Col + 1
                            
                        Case 4
                            sprSTD.Col = sprSTD.Col + 1
                                sTmp = sDiv(0)
                                If LenB(sTmp) > 0 Then Call basFunction.Set_SprType_Text(sprSTD, "CENTER", "LEFT", LenB(sTmp), sTmp)
                                sTam1 = sTmp
                            sprSTD.Col = sprSTD.Col + 1
                                sTmp = sDiv(1)
                                If LenB(sTmp) > 0 Then Call basFunction.Set_SprType_Text(sprSTD, "CENTER", "LEFT", LenB(sTmp), sTmp)
                                sTam2 = sTmp
                            sprSTD.Col = sprSTD.Col + 1
                                sTmp = sDiv(2)
                                If LenB(sTmp) > 0 Then Call basFunction.Set_SprType_Text(sprSTD, "CENTER", "LEFT", LenB(sTmp), sTmp)
                                sTam3 = sTmp
                            sprSTD.Col = sprSTD.Col + 1
                                sTmp = sDiv(3)
                                If LenB(sTmp) > 0 Then Call basFunction.Set_SprType_Text(sprSTD, "CENTER", "LEFT", LenB(sTmp), sTmp)
                                sTam4 = sTmp
                    End Select
                    
                    
                    For ni = 1 To 4 Step 1
                        
                        If ni = 1 Then sGbn = sTam1
                        If ni = 2 Then sGbn = sTam2
                        If ni = 3 Then sGbn = sTam3
                        If ni = 4 Then sGbn = sTam4
                        
                        Select Case sGbn
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
                        
                        sprSTD.Col = sprSTD.Col + 1
                        If LenB(sTmp) > 0 Then Call basFunction.Set_SprType_Text(sprSTD, "CENTER", "LEFT", LenB(sTmp), sTmp)
                    Next ni
                    
                    
                    sprSTD.SetCellBorder sprSTD.Col, sprSTD.Row, sprSTD.Col, sprSTD.Row, 2, basModule.SectionColor1, CellBorderStyleSolid
                    
                    
                    sprSTD.Col = sprSTD.Col + 1
                        sTmp = " ": If IsNull(.Fields("GWA_BAN1")) = False Then sTmp = Trim(.Fields("GWA_BAN1"))
                            Call basFunction.Set_SprType_Text(sprSTD, "CENTER", "LEFT", LenB(sTmp), sTmp)
                    sprSTD.Col = sprSTD.Col + 1
                        sTmp = " ": If IsNull(.Fields("GWA_BAN2")) = False Then sTmp = Trim(.Fields("GWA_BAN2"))
                            Call basFunction.Set_SprType_Text(sprSTD, "CENTER", "LEFT", LenB(sTmp), sTmp)
                    sprSTD.Col = sprSTD.Col + 1
                        sTmp = " ": If IsNull(.Fields("GWA_BAN3")) = False Then sTmp = Trim(.Fields("GWA_BAN3"))
                            Call basFunction.Set_SprType_Text(sprSTD, "CENTER", "LEFT", LenB(sTmp), sTmp)
                    sprSTD.Col = sprSTD.Col + 1
                        sTmp = " ": If IsNull(.Fields("GWA_BAN4")) = False Then sTmp = Trim(.Fields("GWA_BAN4"))
                            Call basFunction.Set_SprType_Text(sprSTD, "CENTER", "LEFT", LenB(sTmp), sTmp)
                    
                    sprSTD.SetCellBorder sprSTD.Col, sprSTD.Row, sprSTD.Col, sprSTD.Row, 2, basModule.SectionColor1, CellBorderStyleSolid
                    
                    sprSTD.Col = sprSTD.Col + 1
                        sTmp = " ": If IsNull(.Fields("GWA_BANNM1")) = False Then sTmp = Trim(.Fields("GWA_BANNM1"))
                            Call basFunction.Set_SprType_Text(sprSTD, "CENTER", "LEFT", LenB(sTmp), sTmp)
                    sprSTD.Col = sprSTD.Col + 1
                        sTmp = " ": If IsNull(.Fields("GWA_BANNM2")) = False Then sTmp = Trim(.Fields("GWA_BANNM2"))
                            Call basFunction.Set_SprType_Text(sprSTD, "CENTER", "LEFT", LenB(sTmp), sTmp)
                    sprSTD.Col = sprSTD.Col + 1
                        sTmp = " ": If IsNull(.Fields("GWA_BANNM3")) = False Then sTmp = Trim(.Fields("GWA_BANNM3"))
                            Call basFunction.Set_SprType_Text(sprSTD, "CENTER", "LEFT", LenB(sTmp), sTmp)
                    sprSTD.Col = sprSTD.Col + 1
                        sTmp = " ": If IsNull(.Fields("GWA_BANNM4")) = False Then sTmp = Trim(.Fields("GWA_BANNM4"))
                            Call basFunction.Set_SprType_Text(sprSTD, "CENTER", "LEFT", LenB(sTmp), sTmp)
                    
                    sprSTD.SetCellBorder sprSTD.Col, sprSTD.Row, sprSTD.Col, sprSTD.Row, 2, basModule.SectionColor1, CellBorderStyleSolid
                    
                    sprSTD.Col = sprSTD.Col + 1
                        sTmp = " ": If IsNull(.Fields("CL_CLOSE")) = False Then sTmp = Trim(.Fields("CL_CLOSE"))
                            Call basFunction.Set_SprType_Text(sprSTD, "CENTER", "LEFT", LenB(sTmp), sTmp)
                    
                End If
                
                .MoveNext       '<< �����׸�
                
            Next nRec
            
            sprSTD.Row = 1:       sprSTD.Row2 = sprSTD.MaxRows
            sprSTD.Col = 1:       sprSTD.Col2 = sprSTD.MaxCols
            sprSTD.BlockMode = True
                sprSTD.BackColor = basModule.WhiteColor
                sprSTD.BackColorStyle = BackColorStyleUnderGrid
            sprSTD.BlockMode = False
            
            sprSTD.Row = 1:       sprSTD.Row2 = sprSTD.MaxRows
            sprSTD.Col = 11:      sprSTD.Col2 = 14
            sprSTD.BlockMode = True
                sprSTD.BackColor = &HFFFFC0
                sprSTD.BackColorStyle = BackColorStyleUnderGrid
            sprSTD.BlockMode = False

        '>> spread lock
            sprSTD.Row = 1:       sprSTD.Row2 = sprSTD.MaxRows
            sprSTD.Col = 1:       sprSTD.Col2 = sprSTD.MaxCols
            sprSTD.BlockMode = True
                sprSTD.Lock = True
                sprSTD.Protect = True
            sprSTD.BlockMode = False
            
        End If
    End With
    
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    MsgBox "��ȸ�� �Ϸ��Ͽ����ϴ�.", vbInformation + vbOKOnly, "�ݺ� ������û���� ��ȸ"
    
    Exit Sub
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    On Error GoTo 0
    MsgBox "�ݺ� ������û���� ��ȸ�� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "�ݺ� ������û���� ��ȸ"
End Sub









'// ������
Private Sub sprGwamok_Click(ByVal Col As Long, ByVal Row As Long)
    
    Dim sSubjCD     As String
    Dim sLsn_FM     As String
    Dim sLsnCD      As String
    Dim sKaeyol     As String
    Dim sLsnType    As String
    
    Dim sSubjNM     As String
    Dim nSubjColor  As Long
    
    Dim sTmp        As String
    
    '< ���� ��ϵ� ���� ��ȸ >
    Dim sStr        As String
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    Dim ni          As Long
    
    Dim sWork       As String
    
    If Row < 1 Then Exit Sub
    If Col < 1 Then Exit Sub
    
    On Error GoTo ErrStmt
    
    
    With sprGwamok      '<< COLUMN���� ������.
        If .MaxCols = 0 Then Exit Sub
        
        Select Case txtKaeyol
            Case "01", "03"
                If optTamgu(0).Value = True Then sSubjCD = "X":     sSubjNM = "":                   nSubjColor = optTamgu(0).BackColor
                If optTamgu(1).Value = True Then sSubjCD = "01":    sSubjNM = optTamgu(1).Caption:  nSubjColor = optTamgu(1).BackColor
                If optTamgu(2).Value = True Then sSubjCD = "02":    sSubjNM = optTamgu(2).Caption:  nSubjColor = optTamgu(2).BackColor
                If optTamgu(3).Value = True Then sSubjCD = "03":    sSubjNM = optTamgu(3).Caption:  nSubjColor = optTamgu(3).BackColor
                If optTamgu(4).Value = True Then sSubjCD = "04":    sSubjNM = optTamgu(4).Caption:  nSubjColor = optTamgu(4).BackColor
                If optTamgu(5).Value = True Then sSubjCD = "05":    sSubjNM = optTamgu(5).Caption:  nSubjColor = optTamgu(5).BackColor
                If optTamgu(6).Value = True Then sSubjCD = "06":    sSubjNM = optTamgu(6).Caption:  nSubjColor = optTamgu(6).BackColor
                If optTamgu(7).Value = True Then sSubjCD = "07":    sSubjNM = optTamgu(7).Caption:  nSubjColor = optTamgu(7).BackColor
                If optTamgu(8).Value = True Then sSubjCD = "08":    sSubjNM = optTamgu(8).Caption:  nSubjColor = optTamgu(8).BackColor
                If optTamgu(9).Value = True Then sSubjCD = "09":    sSubjNM = optTamgu(9).Caption:  nSubjColor = optTamgu(9).BackColor
                If optTamgu(10).Value = True Then sSubjCD = "10":   sSubjNM = optTamgu(10).Caption: nSubjColor = optTamgu(10).BackColor
                If optTamgu(11).Value = True Then sSubjCD = "11":   sSubjNM = optTamgu(11).Caption: nSubjColor = optTamgu(11).BackColor
                
            Case "02"
                If optTamgu(0).Value = True Then sSubjCD = "X":     sSubjNM = "":                   nSubjColor = optTamgu(0).BackColor
                If optTamgu(1).Value = True Then sSubjCD = "01":    sSubjNM = optTamgu(1).Caption:  nSubjColor = optTamgu(1).BackColor
                If optTamgu(2).Value = True Then sSubjCD = "02":    sSubjNM = optTamgu(2).Caption:  nSubjColor = optTamgu(2).BackColor
                If optTamgu(3).Value = True Then sSubjCD = "03":    sSubjNM = optTamgu(3).Caption:  nSubjColor = optTamgu(3).BackColor
                If optTamgu(4).Value = True Then sSubjCD = "04":    sSubjNM = optTamgu(4).Caption:  nSubjColor = optTamgu(4).BackColor
                If optTamgu(5).Value = True Then sSubjCD = "05":    sSubjNM = optTamgu(5).Caption:  nSubjColor = optTamgu(5).BackColor
                If optTamgu(6).Value = True Then sSubjCD = "06":    sSubjNM = optTamgu(6).Caption:  nSubjColor = optTamgu(6).BackColor
                If optTamgu(7).Value = True Then sSubjCD = "07":    sSubjNM = optTamgu(7).Caption:  nSubjColor = optTamgu(7).BackColor
                If optTamgu(8).Value = True Then sSubjCD = "08":    sSubjNM = optTamgu(8).Caption:  nSubjColor = optTamgu(8).BackColor
                
        End Select
    
        '## �̵����� �ð�ǥ ���
        sLsn_FM = Format(Row, "00")         ' ����/�̵� 01,02,03,04
        .Row = SpreadHeader
        .Col = Col
            sLsnCD = Trim(.Text)    ' ���ڵ�
        
        sKaeyol = Trim(txtKaeyol.Text)
        sLsnType = Trim(txtLsnType.Text)
        'sSubjCD        '< ���� �����Ǿ� ������
        
        
        
    '## ���� ��ϵ� ���� ��ȸ -------------------------------------------------------------------------
            Set DBCmd = New ADODB.Command
            Set DBRec = New ADODB.Recordset
            Set DBParam = New ADODB.Parameter
            
            DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
            
            sStr = ""
            sStr = sStr & "  SELECT ACID, LSN_FM, LSNCD, KAEYOL, LSNTYPE, SUBJCD"
            sStr = sStr & "    FROM SDTRX20TB"
            sStr = sStr & "   WHERE ACID   = '" & Trim(basModule.SchCD) & "'"
            sStr = sStr & "     AND LSN_FM = '" & sLsn_FM & "'"
            sStr = sStr & "     AND LSNCD  = '" & sLsnCD & "'"
            
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
            
'            '>> ACID
'            sTmp = Trim(basModule.SchCD)
'            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'                Set DBParam = DBCmd.CreateParameter("LSNCD", adVarChar, adParamInput, nLength, Trim(sTmp)): DBCmd.Parameters.Append DBParam
'            '>> ����/�̵�����
'            sTmp = sLsn_FM
'            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'                Set DBParam = DBCmd.CreateParameter("LSNCD", adVarChar, adParamInput, nLength, Trim(sTmp)): DBCmd.Parameters.Append DBParam
'            '>> �� �ڵ�
'            sTmp = sLsnCD
'            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'                Set DBParam = DBCmd.CreateParameter("LSNNM", adVarChar, adParamInput, nLength, Trim(sTmp)): DBCmd.Parameters.Append DBParam
            
            DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic���·� ���ԵǸ� record count�� �� �� ����.
            Do While DBRec.State And adStateExecuting
                DoEvents
            Loop
            
            With DBRec
                If .RecordCount = 0 Then
                    Select Case sSubjCD
                        Case "X"
                            sWork = "NOT"
                        Case Else
                            sWork = "INSERT"
                    End Select
                Else
                    Select Case sSubjCD
                        Case "X"
                            sWork = "DELETE"
                        Case Else
                            sWork = "UPDATE"
                    End Select
                    
                End If
            End With
    '--------------------------------------------------------------------------------------------------
        
        Select Case sWork
            Case "NOT"
                
            Case "INSERT"
                If inSert_Movement_TimeTable(Trim(basModule.SchCD), sLsn_FM, sLsnCD, sKaeyol, sLsnType, sSubjCD) = False Then
                    MsgBox "��Ͻ� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "�̵����� �ð�ǥ ���"
                Else
                    .Row = Row
                    .Col = Col
                        sTmp = sSubjNM & Space(30) & sSubjCD
                        Call basFunction.Set_SprType_Text(sprGwamok, "center", "left", basFunction.LenKor(sTmp), sTmp)
                        .BackColor = nSubjColor
                End If
            Case "UPDATE"
                If Update_Movement_TimeTable(Trim(basModule.SchCD), sLsn_FM, sLsnCD, sKaeyol, sLsnType, sSubjCD) = False Then
                    MsgBox "���ſ����� �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "�̵����� �ð�ǥ ���"
                Else
                    .Row = Row
                    .Col = Col
                        sTmp = sSubjNM & Space(30) & sSubjCD
                        Call basFunction.Set_SprType_Text(sprGwamok, "center", "left", basFunction.LenKor(sTmp), sTmp)
                        .BackColor = nSubjColor
                End If
                
            Case "DELETE"
                
                Select Case sSubjCD
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
                If MsgBox("��������" & sTmp & " �����Ͻðڽ��ϱ�?", vbQuestion + vbYesNo, "�̵����� �ð�ǥ ���") = vbNo Then
                    Exit Sub
                End If
            
                If Delete_Movement_TimeTable(Trim(basModule.SchCD), sLsn_FM, sLsnCD, sKaeyol, sLsnType, sSubjCD) = False Then
                    MsgBox "����ó���� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "�̵����� �ð�ǥ ���"
                Else
                    .Row = Row
                    .Col = Col
                        .Text = sSubjNM
                        .BackColor = nSubjColor
                        
                End If
        End Select
        
    End With
    
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    Exit Sub
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    MsgBox "ó���� ������ �߻��Ͽ����ϴ�." & vbCrLf & Trim(CStr(Err.Number)) & ":" & Err.Description, vbCritical + vbOKOnly, "�̵����� �ð�ǥ ���"
    On Error GoTo 0
    
End Sub



'>> �̵����� �ð�ǥ ����
Private Function Delete_Movement_TimeTable(ByVal aSchCD As String, _
                                           ByVal aLsn_FM As String, _
                                           ByVal aLsnCD As String, _
                                           ByVal aKaeyol As String, _
                                           ByVal aLsnType As String, _
                                           ByVal aSubjCD As String) As Boolean
    
    Dim sStr        As String
    Dim DBCmd       As ADODB.Command
    Dim DBParam     As ADODB.Parameter
    
    Dim ni          As Long
    
    Dim sTmp        As String
    Dim nLength     As Long
    Dim nExe        As Long
    
    Dim bRet        As Boolean
    
    bRet = False
    
    On Error GoTo ErrStmt
    
    basDataBase.DBConn.BeginTrans
    
    Set DBCmd = New ADODB.Command
    Set DBParam = New ADODB.Parameter
    
    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
                
    '>> ���� �Ķ���Ͱ� ���� ������ �޸𸮿��� ������.
    For ni = 0 To DBCmd.Parameters.Count - 1 Step 1
        DBCmd.Parameters.Delete (0)
    Next ni
    
    nExe = 0
    
    sStr = ""
    sStr = sStr & "  DELETE "
    sStr = sStr & "    FROM SDTRX20TB"
    sStr = sStr & "   WHERE ACID   = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "     AND LSN_FM = '" & aLsn_FM & "'"
    sStr = sStr & "     AND LSNCD  = '" & aLsnCD & "'"
            
    '>> ���� �Ķ���Ͱ� ���� ������ �޸𸮿��� ������.
    For ni = 0 To DBCmd.Parameters.Count - 1 Step 1
        DBCmd.Parameters.Delete (0)
    Next ni
            
'    '>> ACID
'    sTmp = Trim(basModule.SchCD)
'    nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'        Set DBParam = DBCmd.CreateParameter("LSNCD", adVarChar, adParamInput, nLength, Trim(sTmp)): DBCmd.Parameters.Append DBParam
'    '>> ����/�̵�����
'    sTmp = aLsn_FM
'    nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'        Set DBParam = DBCmd.CreateParameter("LSNCD", adVarChar, adParamInput, nLength, Trim(sTmp)): DBCmd.Parameters.Append DBParam
'    '>> �� �ڵ�
'    sTmp = aLsnCD
'    nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'        Set DBParam = DBCmd.CreateParameter("LSNNM", adVarChar, adParamInput, nLength, Trim(sTmp)): DBCmd.Parameters.Append DBParam
            
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
    Else
        basDataBase.DBConn.RollbackTrans
    End If
        
    Set DBCmd = Nothing
    Set DBParam = Nothing
    
    Delete_Movement_TimeTable = bRet
    Exit Function
    
ErrStmt:
    basDataBase.DBConn.RollbackTrans
    On Error GoTo 0
    
    Set DBCmd = Nothing
    Set DBParam = Nothing
    
    Delete_Movement_TimeTable = bRet
    
End Function


'## �̵������ð�ǥ ���
Private Function inSert_Movement_TimeTable(ByVal aSchCD As String, _
                                           ByVal aLsn_FM As String, _
                                           ByVal aLsnCD As String, _
                                           ByVal aKaeyol As String, _
                                           ByVal aLsnType As String, _
                                           ByVal aSubjCD As String) As Boolean
    
    Dim sStr        As String
    Dim DBCmd       As ADODB.Command
    Dim DBParam     As ADODB.Parameter
    
    Dim ni          As Long
    
    Dim sTmp        As String
    Dim nLength     As Long
    Dim nExe        As Long
    
    Dim bRet        As Boolean
    
    bRet = False
    
    On Error GoTo ErrStmt
    
    basDataBase.DBConn.BeginTrans
    
    Set DBCmd = New ADODB.Command
    Set DBParam = New ADODB.Parameter
    
    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
    
    nExe = 0
    
    sStr = ""
    sStr = sStr & "  INSERT INTO SDTRX20TB (ACID, LSN_FM, LSNCD, KAEYOL, LSNTYPE, SUBJCD)"
    sStr = sStr & "  VALUES ( "
    sStr = sStr & "          '" & aSchCD & "',"
    sStr = sStr & "          '" & aLsn_FM & "',"
    sStr = sStr & "          '" & aLsnCD & "',"
    sStr = sStr & "          '" & aKaeyol & "',"
    sStr = sStr & "          '" & aLsnType & "',"
    sStr = sStr & "          '" & aSubjCD & "'"
    sStr = sStr & "  ) "
    
    '>> ���� �Ķ���Ͱ� ���� ������ �޸𸮿��� ������.
    For ni = 0 To DBCmd.Parameters.Count - 1 Step 1
        DBCmd.Parameters.Delete (0)
    Next ni
            
'    '>> ACID
'    sTmp = Trim(basModule.SchCD)
'    nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'        Set DBParam = DBCmd.CreateParameter("LSNCD", adVarChar, adParamInput, nLength, Trim(sTmp)): DBCmd.Parameters.Append DBParam
'    '>> ����/�̵�����
'    sTmp = aLsn_FM
'    nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'        Set DBParam = DBCmd.CreateParameter("LSNCD", adVarChar, adParamInput, nLength, Trim(sTmp)): DBCmd.Parameters.Append DBParam
'    '>> �� �ڵ�
'    sTmp = aLsnCD
'    nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'        Set DBParam = DBCmd.CreateParameter("LSNNM", adVarChar, adParamInput, nLength, Trim(sTmp)): DBCmd.Parameters.Append DBParam
'    '>> �迭
'    sTmp = AKAEYOL
'    nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'        Set DBParam = DBCmd.CreateParameter("LSNNM", adVarChar, adParamInput, nLength, Trim(sTmp)): DBCmd.Parameters.Append DBParam
'    '>> �� ����
'    sTmp = ALSNTYPE
'    nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'        Set DBParam = DBCmd.CreateParameter("LSNNM", adVarChar, adParamInput, nLength, Trim(sTmp)): DBCmd.Parameters.Append DBParam
'    '>> ����
'    sTmp = ASUBJCD
'    nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'        Set DBParam = DBCmd.CreateParameter("LSNNM", adVarChar, adParamInput, nLength, Trim(sTmp)): DBCmd.Parameters.Append DBParam
            
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
    Else
        basDataBase.DBConn.RollbackTrans
    End If
        
    Set DBCmd = Nothing
    Set DBParam = Nothing
    
    inSert_Movement_TimeTable = bRet
    Exit Function
    
ErrStmt:
    basDataBase.DBConn.RollbackTrans
    On Error GoTo 0
    
    Set DBCmd = Nothing
    Set DBParam = Nothing
    
    inSert_Movement_TimeTable = bRet
End Function


'## �̵������ð�ǥ ���
Private Function Update_Movement_TimeTable(ByVal aSchCD As String, _
                                           ByVal aLsn_FM As String, _
                                           ByVal aLsnCD As String, _
                                           ByVal aKaeyol As String, _
                                           ByVal aLsnType As String, _
                                           ByVal aSubjCD As String) As Boolean

    Dim sStr        As String
    Dim DBCmd       As ADODB.Command
    Dim DBParam     As ADODB.Parameter
    
    Dim ni          As Long
    
    Dim sTmp        As String
    Dim nLength     As Long
    Dim nExe        As Long
    
    Dim bRet        As Boolean
    
    bRet = False
    
    On Error GoTo ErrStmt
    
    basDataBase.DBConn.BeginTrans
    
    Set DBCmd = New ADODB.Command
    Set DBParam = New ADODB.Parameter
    
    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
    
    nExe = 0
    
    sStr = ""
    sStr = sStr & "  UPDATE SDTRX20TB "
    sStr = sStr & "     SET KAEYOL  = '" & aKaeyol & "', "
    sStr = sStr & "         LSNTYPE = '" & aLsnType & "', "
    sStr = sStr & "         SUBJCD  = '" & aSubjCD & "' "
    
    sStr = sStr & "   WHERE ACID    = '" & aSchCD & "'"
    sStr = sStr & "     AND LSN_FM  = '" & aLsn_FM & "'"
    sStr = sStr & "     AND LSNCD   = '" & aLsnCD & "'"
    
    '>> ���� �Ķ���Ͱ� ���� ������ �޸𸮿��� ������.
    For ni = 0 To DBCmd.Parameters.Count - 1 Step 1
        DBCmd.Parameters.Delete (0)
    Next ni
            
'    '>> �迭
'    sTmp = AKAEYOL
'    nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'        Set DBParam = DBCmd.CreateParameter("LSNNM", adVarChar, adParamInput, nLength, Trim(sTmp)): DBCmd.Parameters.Append DBParam
'    '>> �� ����
'    sTmp = ALSNTYPE
'    nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'        Set DBParam = DBCmd.CreateParameter("LSNNM", adVarChar, adParamInput, nLength, Trim(sTmp)): DBCmd.Parameters.Append DBParam
'    '>> ����
'    sTmp = ASUBJCD
'    nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'        Set DBParam = DBCmd.CreateParameter("LSNNM", adVarChar, adParamInput, nLength, Trim(sTmp)): DBCmd.Parameters.Append DBParam

'    '>> ACID
'    sTmp = Trim(basModule.SchCD)
'    nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'        Set DBParam = DBCmd.CreateParameter("LSNCD", adVarChar, adParamInput, nLength, Trim(sTmp)): DBCmd.Parameters.Append DBParam
'    '>> ����/�̵�����
'    sTmp = aLsn_FM
'    nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'        Set DBParam = DBCmd.CreateParameter("LSNCD", adVarChar, adParamInput, nLength, Trim(sTmp)): DBCmd.Parameters.Append DBParam
'    '>> �� �ڵ�
'    sTmp = aLsnCD
'    nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'        Set DBParam = DBCmd.CreateParameter("LSNNM", adVarChar, adParamInput, nLength, Trim(sTmp)): DBCmd.Parameters.Append DBParam
            
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
    Else
        basDataBase.DBConn.RollbackTrans
    End If
        
    Set DBCmd = Nothing
    Set DBParam = Nothing
    
    Update_Movement_TimeTable = bRet
    Exit Function
    
ErrStmt:
    basDataBase.DBConn.RollbackTrans
    On Error GoTo 0
    
    Set DBCmd = Nothing
    Set DBParam = Nothing
    
    Update_Movement_TimeTable = bRet
End Function





















'## �̵��� ���
Private Sub cmdinSertLsn_Click()
    Dim nCol        As Long
    
    Dim sLsnCD      As String
    Dim sTmp        As String
    
    With sprGwamok
    
        If .MaxCols = 0 Then
            MsgBox "�ݺ� ���� ��û���� ��ȸ�� �̵����� �߰��Ͻʽÿ�.", vbExclamation + vbOKOnly, "�̵��� ���"
            Exit Sub
        End If
        
        sLsnCD = ""
        For nCol = 1 To .MaxCols Step 1
            .Col = nCol
            .Row = SpreadHeader
                sTmp = Trim(.Text)
                
            If sTmp > sLsnCD Then
                sLsnCD = sTmp
            End If
        Next nCol
        
        If sLsnCD > "90000" Then
            sLsnCD = Format(CLng(sLsnCD) + 1, "00000")
        Else
            sLsnCD = "90001"
        End If
        
        .MaxCols = .MaxCols + 1
        .Col = .MaxCols
        .Row = SpreadHeader
            .Text = sLsnCD
        .Row = SpreadHeader + 1
            .Text = "�̵�" & Trim(CStr(CLng(sLsnCD) - 90000))
            
    End With
End Sub






'## �̵��� ����
Private Sub cmdDeleteLsn_Click()
    Dim sStr        As String
    Dim DBCmd       As ADODB.Command
    Dim DBParam     As ADODB.Parameter
    
    Dim ni          As Long
    
    Dim sTmp        As String
    Dim nLength     As Long
    Dim nExe        As Long
    
    Dim nCol        As Long
    Dim bChk        As Boolean
    Dim sDelLsnCD   As String
    Dim sDelLsnNM   As String
    
    
    '## ������ �̵��� ��ȸ --------------------------
    bChk = False
    With sprGwamok
        For nCol = .MaxCols To 1 Step -1
            .Row = SpreadHeader
            .Col = nCol
            If Trim(.Text) > "90000" Then
                sDelLsnCD = Trim(.Text)
                    bChk = True
                
                .Row = SpreadHeader + 1
                sDelLsnNM = Trim(.Text)
                Exit For
            End If
        Next nCol
    End With
    
    If MsgBox("�� " & sDelLsnNM & " ���� ��ü������ �����Ͻðڽ��ϱ�?", vbQuestion + vbYesNo, "�̵������� ����") = vbNo Then
        MsgBox "����Ͽ����ϴ�.", vbInformation + vbOKOnly, "�̵������� ����"
        Exit Sub
    End If
    '------------------------------------------------
    
    
    On Error GoTo ErrStmt
    
    basDataBase.DBConn.BeginTrans
    
    Set DBCmd = New ADODB.Command
    Set DBParam = New ADODB.Parameter
    
    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
                
    '>> ���� �Ķ���Ͱ� ���� ������ �޸𸮿��� ������.
    For ni = 0 To DBCmd.Parameters.Count - 1 Step 1
        DBCmd.Parameters.Delete (0)
    Next ni
    
    nExe = 0
    
    sStr = ""
    sStr = sStr & "  DELETE "
    sStr = sStr & "    FROM SDTRX20TB"
    sStr = sStr & "   WHERE ACID   = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "     AND LSNCD  = '" & sDelLsnCD & "'"
            
    '>> ���� �Ķ���Ͱ� ���� ������ �޸𸮿��� ������.
    For ni = 0 To DBCmd.Parameters.Count - 1 Step 1
        DBCmd.Parameters.Delete (0)
    Next ni
            
'    '>> ACID
'    sTmp = Trim(basModule.SchCD)
'    nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'        Set DBParam = DBCmd.CreateParameter("LSNCD", adVarChar, adParamInput, nLength, Trim(sTmp)): DBCmd.Parameters.Append DBParam
'    '>> ����/�̵�����
'    sTmp = aLsn_FM
'    nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'        Set DBParam = DBCmd.CreateParameter("LSNCD", adVarChar, adParamInput, nLength, Trim(sTmp)): DBCmd.Parameters.Append DBParam
'    '>> �� �ڵ�
'    sTmp = aLsnCD
'    nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'        Set DBParam = DBCmd.CreateParameter("LSNNM", adVarChar, adParamInput, nLength, Trim(sTmp)): DBCmd.Parameters.Append DBParam
            
    DBCmd.CommandText = sStr
    DBCmd.CommandType = adCmdText
    DBCmd.CommandTimeout = 30
    
    DBCmd.Execute nExe, , -1
            
    Do While basDataBase.DBConn.State And adStateExecuting
        DoEvents
    Loop
    
    If nExe >= 1 Then
        basDataBase.DBConn.CommitTrans
        MsgBox "�̵��������� �����Ͽ����ϴ�.", vbInformation + vbOKOnly, "�̵������� ����"
        
        With sprGwamok
            .MaxCols = .MaxCols - 1
            
        End With
    ElseIf nExe = 0 Then
        basDataBase.DBConn.CommitTrans
        MsgBox "������ ������ �����ϴ�." & vbCrLf & _
               "�̵��� �߰��� ������ ����Ͻ��� ������" & vbCrLf & _
               "������ ������ �����ϴ�.", vbExclamation + vbOKOnly, "�̵������� ����"
               
        With sprGwamok
            .MaxCols = .MaxCols - 1
            
        End With
    Else
        basDataBase.DBConn.RollbackTrans
        MsgBox "�̵������� ������ ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "�̵������� ����"
    End If
        
    Set DBCmd = Nothing
    Set DBParam = Nothing
    
    Exit Sub
    
ErrStmt:
    basDataBase.DBConn.RollbackTrans
    On Error GoTo 0
    
    Set DBCmd = Nothing
    Set DBParam = Nothing
End Sub





'## ��ϳ��� ��ȸ
Private Sub cmdFindLsn_Click()
    On Error GoTo ErrStmt
    
    Call Find_Detail_Lsn_Header         '> �� ������ȸ
    Call Find_Detail_Gwamok_Data        '> ���� ������
    
    MsgBox "��ȸ�Ͽ����ϴ�.", vbInformation + vbOKOnly, "�̵����� �ð�ǥ ��ȸ"
    
    Exit Sub
ErrStmt:
    MsgBox "��ȸ�� ������ �߻��Ͽ����ϴ�." & vbCrLf & _
           Trim(CStr(Err.Number)) & ":" & Trim(Err.Description), vbCritical + vbOKOnly, "�̵����� �ð�ǥ ��ȸ"
    On Error GoTo 0
End Sub

'## ���� �󼼳��� ��ȸ
Private Sub Find_Detail_Gwamok_Data()
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    
    Dim nLength     As Long
    Dim sStr        As String
    Dim ni          As Integer
    Dim nRec        As Long
    
    Dim sTmp        As String
    Dim nTmp        As Double
    
    Dim nRow        As Long
    Dim nCol        As Long
    
    Dim sLsn_FM     As String
    Dim sLsnCD      As String
    Dim sSubjCD     As String
    Dim sSubj       As String
    Dim nSubjColor  As Long
    
    Dim nLsn_FM     As Long
    Dim nLsnCD      As Long
    
    On Error GoTo ErrStmt
    
    If Trim(txtKaeyol.Text) = "" Then
        MsgBox "�̵��ð�ǥ ��ȸ�� �迭�� �����ϼ���.", vbExclamation + vbOKOnly, "�̵��ð�ǥ ��ȸ"
        Exit Sub
    End If
    
    If Trim(txtLsnType) = "" Then
        MsgBox "�����¸� �����ϼ���.", vbExclamation + vbOKOnly, "�̵��ð�ǥ ��ȸ"
        Exit Sub
    ElseIf Trim(txtLsnType) = "ALL" Then
        MsgBox "�����¸� �����ϼ���.", vbExclamation + vbOKOnly, "�̵��ð�ǥ ��ȸ"
    End If
    
    sStr = ""
    sStr = sStr & "  SELECT ACID, LSN_FM, LSNCD, SUBJCD, KAEYOL "
    sStr = sStr & "    From SDTRX20TB"
    sStr = sStr & "   WHERE ACID    = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "     AND KAEYOL  = '" & Trim(txtKaeyol.Text) & "'"
    sStr = sStr & "     AND LSNTYPE = '" & Trim(txtLsnType.Text) & "'"
    sStr = sStr & "   ORDER BY LSNCD, LSN_FM"
    
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
    
'    '>> �п�
'        sTmp = Trim(basModule.SchCD)
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
'    '>> �迭
'        sTmp = Trim(txtKaeyol.Text)
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
'    '>> �� ����
'        sTmp = Trim(txtLsnType.Text)
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
                
                If IsNull(.Fields("LSN_FM")) = False And _
                   IsNull(.Fields("LSNCD")) = False Then
                   
                    sLsn_FM = Trim(.Fields("LSN_FM"))
                    sLsnCD = Trim(.Fields("LSNCD"))
                    sSubjCD = Trim(.Fields("SUBJCD"))
                    
                    Select Case sLsn_FM
                        Case "01"
                            nLsn_FM = 1
                        Case "02"
                            nLsn_FM = 2
                        Case "03"
                            nLsn_FM = 3
                        Case "04"
                            nLsn_FM = 4
                    End Select
                    
                    For nCol = 1 To sprGwamok.MaxCols Step 1
                        sprGwamok.Row = SpreadHeader
                        sprGwamok.Col = nCol
                            sTmp = sprGwamok.Text
                        If StrComp(sLsnCD, sTmp, vbTextCompare) = 0 Then
                            nLsnCD = sprGwamok.Col
                            Exit For
                        End If
                    Next nCol
                    
                    
                    optTamgu(0).Value = True
                    Select Case Trim(.Fields("KAEYOL"))
                        Case "01", "02"
                            Select Case Trim(.Fields("SUBJCD"))
                                Case "01"
                                    sSubj = "����"
                                    nSubjColor = optTamgu(1).BackColor
                                Case "02"
                                    sSubj = "����"
                                    nSubjColor = optTamgu(2).BackColor
                                Case "03"
                                    sSubj = "����"
                                    nSubjColor = optTamgu(3).BackColor
                                Case "04"
                                    sSubj = "�ѱ�"
                                    nSubjColor = optTamgu(4).BackColor
                                Case "05"
                                    sSubj = "�����"
                                    nSubjColor = optTamgu(5).BackColor
                                Case "06"
                                    sSubj = "����"
                                    nSubjColor = optTamgu(6).BackColor
                                Case "07"
                                    sSubj = "����"
                                    nSubjColor = optTamgu(7).BackColor
                                Case "08"
                                    sSubj = "��ġ"
                                    nSubjColor = optTamgu(8).BackColor
                                Case "09"
                                    sSubj = "�繮"
                                    nSubjColor = optTamgu(9).BackColor
                                Case "10"
                                    sSubj = "����"
                                    nSubjColor = optTamgu(10).BackColor
                                Case "11"
                                    sSubj = "����"
                                    nSubjColor = optTamgu(11).BackColor
                            End Select
                        Case "03"
                            Select Case Trim(.Fields("SUBJCD"))
                                Case "01"
                                    sSubj = "��1"
                                    nSubjColor = optTamgu(1).BackColor
                                Case "02"
                                    sSubj = "ȭ1"
                                    nSubjColor = optTamgu(2).BackColor
                                Case "03"
                                    sSubj = "��1"
                                    nSubjColor = optTamgu(3).BackColor
                                Case "04"
                                    sSubj = "��1"
                                    nSubjColor = optTamgu(4).BackColor
                                Case "05"
                                    sSubj = "��2"
                                    nSubjColor = optTamgu(5).BackColor
                                Case "06"
                                    sSubj = "ȭ2"
                                    nSubjColor = optTamgu(6).BackColor
                                Case "07"
                                    sSubj = "��2"
                                    nSubjColor = optTamgu(7).BackColor
                                Case "08"
                                    sSubj = "��2"
                                    nSubjColor = optTamgu(8).BackColor
                            End Select
                    End Select
                    
                    sSubj = sSubj & Space(30) & Trim(.Fields("SUBJCD"))
                    'nSubjColor
                    
                    sprGwamok.Row = nLsn_FM
                    sprGwamok.Col = nLsnCD
                        Call basFunction.Set_SprType_Text(sprGwamok, "center", "left", basFunction.LenKor(sSubj), sSubj)
                        sprGwamok.BackColor = nSubjColor
                        
                End If
                
                .MoveNext       '<< �����׸�
                
            Next nRec
         
            
         
        End If
    End With
    
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    Exit Sub
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    On Error GoTo 0
    MsgBox "�̵�����ǥ ��ȸ�� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "�̵�����ǥ ��ȸ"
    
End Sub

'## �̵��ð�ǥ ��� ó��
Private Sub Find_Detail_Lsn_Header()
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    
    Dim nLength     As Long
    Dim sStr        As String
    Dim ni          As Integer
    Dim nRec        As Long
    
    Dim sTmp        As String
    Dim nTmp        As Double
    
    On Error GoTo ErrStmt
    
    If Trim(txtKaeyol.Text) = "" Then
        MsgBox "�̵��ð�ǥ ��ȸ�� �迭�� �����ϼ���.", vbExclamation + vbOKOnly, "�̵��ð�ǥ ��ȸ"
        Exit Sub
    End If
    
    If Trim(txtLsnType) = "" Then
        MsgBox "�����¸� �����ϼ���.", vbExclamation + vbOKOnly, "�̵��ð�ǥ ��ȸ"
        Exit Sub
    ElseIf Trim(txtLsnType) = "ALL" Then
        MsgBox "�����¸� �����ϼ���.", vbExclamation + vbOKOnly, "�̵��ð�ǥ ��ȸ"
    End If
    
    
    sprGwamok.MaxCols = 0
    
    sStr = ""
    sStr = sStr & "  SELECT LSNCD, GET_LSNNM(LSNCD) AS LSNNM"
    sStr = sStr & "    FROM (SELECT LSNCD"
    sStr = sStr & "            FROM SDTRX20TB"
    sStr = sStr & "           WHERE ACID    = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "             AND KAEYOL  = '" & Trim(txtKaeyol.Text) & "'"
    sStr = sStr & "             AND LSNTYPE = '" & Trim(txtLsnType.Text) & "'"
    sStr = sStr & "           GROUP BY LSNCD"
    sStr = sStr & "          )"
    sStr = sStr & "   ORDER BY LSNCD"
    
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
    
'    '>> �п�
'        sTmp = Trim(basModule.SchCD)
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
'    '>> �迭
'        sTmp = Trim(txtKaeyol.Text)
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
'    '>> �� ����
'        sTmp = Trim(txtLsnType.Text)
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
       
        
    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic���·� ���ԵǸ� record count�� �� �� ����.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop
    
    With DBRec
        If .RecordCount > 0 Then
            .MoveFirst
            
            sprGwamok.ColHeaderRows = 2
            sprGwamok.MaxCols = .RecordCount
            
            For nRec = 1 To .RecordCount Step 1
                sprGwamok.Col = nRec
                
                sprGwamok.Row = SpreadHeader:       sprGwamok.RowHeight(sprGwamok.Row) = nRowHeight
                    sTmp = " ": If IsNull(.Fields("LSNCD")) = False Then sTmp = Trim(.Fields("LSNCD"))
                    sprGwamok.Text = sTmp
                    sprGwamok.RowHidden = True
                    
                sprGwamok.Row = SpreadHeader + 1:   sprGwamok.RowHeight(sprGwamok.Row) = nRowHeight + 2
                    sTmp = " ": If IsNull(.Fields("LSNNM")) = False Then sTmp = Trim(.Fields("LSNNM"))
                    sprGwamok.Text = sTmp
                   
                .MoveNext       '<< �����׸�
            Next nRec
        End If
    End With
    
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    Exit Sub
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    On Error GoTo 0
    MsgBox "�̵��ð�ǥ ��ȸ�� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "�̵��ð�ǥ ��ȸ"
    
End Sub


Private Sub sprSTD_Click(ByVal Col As Long, ByVal Row As Long)
    Dim sDiv()      As String
    
    Dim sT1         As String
    Dim sT2         As String
    
    If Row < 1 Then Exit Sub
    If Col < 1 Then Exit Sub
    
    
    With sprSTD
        If Trim(.Tag) = "" Then .Tag = "1"
        
        .Row = CLng(.Tag):  .Row2 = .Row
        .Col = 1:           .Col2 = .MaxCols
        .BlockMode = True
            .BackColor = basModule.WhiteColor
            .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
        
        .Row = 1:       .Row2 = .MaxRows
        .Col = 11:      .Col2 = 14
        .BlockMode = True
            .BackColor = &HFFFFC0
            .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
        
        .Row = Row:     .Row2 = .Row
        .Col = 1:       .Col2 = .MaxCols
        .BlockMode = True
            .BackColor = basModule.BackColor2
            .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
        
        .Tag = Trim(CStr(Row))
        
        If cboLsnin.ListCount > 1 Then
            
            sDiv = Split(cboLsnin.Text, "[T]", -1, vbTextCompare)
            sT1 = Trim(sDiv(0))         '< �ݸ�
            sT2 = Trim(sDiv(1))         '< ���ڵ�
            
            .Row = Row
            Select Case sT2
            
                Case "IN"
                    
                Case "OUT"
                    .Col = Col:     .Text = ""
                    .Col = Col - 4: .Text = ""
                Case Else
                    .Col = Col
                        Call basFunction.Set_SprType_Text(sprSTD, "center", "left", basFunction.LenKor(sT1), sT1)
                    .Col = Col - 4
                        Call basFunction.Set_SprType_Text(sprSTD, "center", "left", basFunction.LenKor(sT2), sT2)
                    .Col = 4
                        .Value = 1
            
            End Select
        End If
    End With
    
End Sub












'######################################################################################################################################################
'## �л�üũ
'######################################################################################################################################################

'## �������ɿ��� Ȯ��
Private Sub cmdStdGwamokChk_Click()

    '## �� ������ �۾� ��������� copy
    Call Process_LsnCopySpread

    If sprCopyLsn.MaxRows > 0 And sprSTD.MaxRows > 0 Then
        Call Gwamok_Matching                                '<< 1���� �ݿ��� ��� ���� �� �ִ� ����� ����

        Call TMR021.Show_TMR_WorkSheet_Data(sprCopyLsn, Trim(txtKaeyol.Text))

        Load TMR021
        TMR021.Show

    End If

    MsgBox "�Ϸ��Ͽ����ϴ�.", vbInformation + vbOKOnly, "�������ɿ��� Ȯ��"
    
End Sub



Private Sub cmdStdGwamokChk_Show_Click()
    If Trim(txtKaeyol.Text) = "" Then
        MsgBox "�ݺ� ���� ��û���� ��ȸ�� �Ͻʽÿ�.", vbExclamation + vbOKOnly, "�������� ó������ ����"
        Exit Sub
    End If
    
    If sprCopyLsn.MaxRows > 0 Then
        Call TMR021.Show_TMR_WorkSheet_Data(sprCopyLsn, Trim(txtKaeyol.Text))

        Load TMR021
        TMR021.Show
    End If
End Sub

Private Sub cmdOrdGwamok_View_Click()
    If Trim(txtKaeyol.Text) = "" Then
        MsgBox "�ݺ� ���� ��û���� ��ȸ�� �Ͻʽÿ�.", vbExclamation + vbOKOnly, "�л���û���� ��ģ���� ����"
        Exit Sub
    End If

    'Call TMR022.Show_OrdGwamok_STD(Trim(txtKaeyol.Text))

    Load TMR022
    TMR022.Show
End Sub




'-----------------------------------------------
'�ݺ� ��û�ο� ��ȸ
'-----------------------------------------------
Private Sub Process_LsnCopySpread()

    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    
    Dim nLength     As Long
    Dim sStr        As String
    Dim ni          As Integer
    Dim nRec        As Long
    
    Dim sTmp        As String
    Dim nTmp        As Double
    
    Dim nCol        As Integer
    Dim siTem       As String
    
    Dim nRet        As Long
    
    Dim nRow_LSN    As Long
    Dim nCol_Lsn    As Long
    Dim sSql_Lsn    As String
    Dim sGwamok     As String
    
    On Error GoTo ErrStmt
    
    With sprLsn
        sSql_Lsn = ""
        For nRow_LSN = 1 To .MaxRows Step 1
            .Row = nRow_LSN
            .Col = 3
            If .Value = 1 Then
                If sSql_Lsn > " " Then sSql_Lsn = sSql_Lsn & ","
                If sSql_Lsn = "" Then sSql_Lsn = sSql_Lsn & "("
                
                .Col = 1
                    sSql_Lsn = sSql_Lsn & "'" & Trim(.Text) & "'"
            End If
        Next nRow_LSN
    End With
    
    If sSql_Lsn = "" Then
        MsgBox "�۾���� ���� �����ϴ�.", vbExclamation + vbOKOnly, "�ݺ� ��û�ο� ��ȸ"
        Exit Sub
    Else
        sSql_Lsn = sSql_Lsn & ")"
    End If
    
    
    sprCopyLsn.MaxRows = 0
    If sprLsn.MaxCols = 0 Then Exit Sub
        sprCopyLsn.MaxCols = sprLsn.MaxCols - 2         '< �۾�����/ ó���ο��� ����
        sprCopyLsn.Col = 1

    For nCol_Lsn = 1 To sprLsn.MaxCols Step 1
        sprLsn.Row = SpreadHeader
        sprLsn.Col = nCol_Lsn:      sTmp = Trim(sprLsn.Text)

        sprCopyLsn.Row = SpreadHeader

            Select Case nCol_Lsn
                Case 1, 2, 5

                    sprCopyLsn.Text = sTmp:     sprCopyLsn.ColWidth(sprCopyLsn.Col) = 7
                    sprCopyLsn.Col = sprCopyLsn.Col + 1
                    
                Case 6 To sprLsn.MaxCols
                    Select Case sTmp
                        Case "����"
                            sGwamok = "01"
                        Case "����"
                            sGwamok = "02"
                        Case "����"
                            sGwamok = "03"
                        Case "�ѱ�"
                            sGwamok = "04"
                        Case "�����"
                            sGwamok = "05"
                        Case "����"
                            sGwamok = "06"
                        Case "����"
                            sGwamok = "07"
                        Case "��ġ"
                            sGwamok = "08"
                        Case "�繮"
                            sGwamok = "09"
                        Case "����"
                            sGwamok = "10"
                        Case "����"
                            sGwamok = "11"
                        Case "��1"
                            sGwamok = "51"
                        Case "ȭ1"
                            sGwamok = "52"
                        Case "��1"
                            sGwamok = "53"
                        Case "��1"
                            sGwamok = "54"
                        Case "��2"
                            sGwamok = "55"
                        Case "ȭ2"
                            sGwamok = "56"
                        Case "��2"
                            sGwamok = "57"
                        Case "��2"
                            sGwamok = "58"
                    End Select
                    sprCopyLsn.Text = sGwamok:  sprCopyLsn.ColWidth(sprCopyLsn.Col) = 5
                    
                    sprCopyLsn.Col = sprCopyLsn.Col + 1
                    
            End Select
    Next nCol_Lsn
    
    
    
    
    nRet = 0
    
    sStr = ""
    sStr = sStr & "  SELECT LSNCD, LSNNM, "
    sStr = sStr & "         S_LSN,"
    sStr = sStr & "         SEL1 ,"
    sStr = sStr & "         SEL2 ,"
    sStr = sStr & "         SEL3 ,"
    sStr = sStr & "         SEL4 ,"
    sStr = sStr & "         SEL5 ,"
    sStr = sStr & "         SEL6 ,"
    sStr = sStr & "         SEL7 ,"
    sStr = sStr & "         SEL8 ,"
    sStr = sStr & "         SEL9 ,"
    sStr = sStr & "         SEL10,"
    sStr = sStr & "         SEL11,"
    sStr = sStr & "         KAEYOL"
    
    sStr = sStr & "    FROM (SELECT LSNCD,"
    sStr = sStr & "                 GET_LSNNM(LSNCD) AS LSNNM,"
    
    sStr = sStr & "                 0 AS S_LSN,"
    sStr = sStr & "                 SUM(SEL1 ) AS SEL1 ,"
    sStr = sStr & "                 SUM(SEL2 ) AS SEL2 ,"
    sStr = sStr & "                 SUM(SEL3 ) AS SEL3 ,"
    sStr = sStr & "                 SUM(SEL4 ) AS SEL4 ,"
    sStr = sStr & "                 SUM(SEL5 ) AS SEL5 ,"
    sStr = sStr & "                 SUM(SEL6 ) AS SEL6 ,"
    sStr = sStr & "                 SUM(SEL7 ) AS SEL7 ,"
    sStr = sStr & "                 SUM(SEL8 ) AS SEL8 ,"
    sStr = sStr & "                 SUM(SEL9 ) AS SEL9 ,"
    sStr = sStr & "                 SUM(SEL10) AS SEL10,"
    sStr = sStr & "                 SUM(SEL11) AS SEL11,"
    sStr = sStr & "                 MAX(GAEYUL_CD) AS KAEYOL"
    
    sStr = sStr & "           FROM (SELECT LSNCD, "
    sStr = sStr & "                        GAEYUL_CD,"
    
    sStr = sStr & "                        SEL1 ,"
    sStr = sStr & "                        SEL2 ,"
    sStr = sStr & "                        SEL3 ,"
    sStr = sStr & "                        SEL4 ,"
    sStr = sStr & "                        SEL5 ,"
    sStr = sStr & "                        SEL6 ,"
    sStr = sStr & "                        SEL7 ,"
    sStr = sStr & "                        SEL8 ,"
    sStr = sStr & "                        SEL9 ,"
    sStr = sStr & "                        SEL10,"
    sStr = sStr & "                        SEL11,"
    sStr = sStr & "                        CL_CLOSE "
    
    sStr = sStr & "                  FROM (SELECT SEL_CLASS AS LSNCD,"
    sStr = sStr & "                               CASE WHEN SEL1 > ' ' THEN"
    sStr = sStr & "                                  '01'"
    sStr = sStr & "                               ELSE CASE WHEN SEL3 > ' ' THEN"
    sStr = sStr & "                                  '02'"
    sStr = sStr & "                               END END GAEYUL_CD,"
    
    sStr = sStr & "                        /* ��Ž, ��Ž �и� */"
    sStr = sStr & "                               CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'01|') > 0 THEN          /* ��Ž-���� */"
    sStr = sStr & "                                  1"
    sStr = sStr & "                               ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'51|') > 0 THEN     /* ��Ž-����1 */"
    sStr = sStr & "                                  1"
    sStr = sStr & "                               ELSE"
    sStr = sStr & "                                  0"
    sStr = sStr & "                               END END SEL1,"
    sStr = sStr & "                               CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'02|') > 0 THEN          /* ��Ž-���� */"
    sStr = sStr & "                                  1"
    sStr = sStr & "                               ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'52|') > 0 THEN     /* ��Ž-ȭ��1 */"
    sStr = sStr & "                                  1"
    sStr = sStr & "                               ELSE"
    sStr = sStr & "                                  0"
    sStr = sStr & "                               END END SEL2,"
    sStr = sStr & "                               CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'03|') > 0 THEN          /* ��Ž-���� */"
    sStr = sStr & "                                  1"
    sStr = sStr & "                               ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'53|') > 0 THEN     /* ��Ž-����1 */"
    sStr = sStr & "                                  1"
    sStr = sStr & "                               ELSE"
    sStr = sStr & "                                  0"
    sStr = sStr & "                               END END SEL3,"
    sStr = sStr & "                               CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'04|') > 0 THEN          /* ��Ž-�ѱ������� */"
    sStr = sStr & "                                  1"
    sStr = sStr & "                               ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'54|') > 0 THEN     /* ��Ž-��������1 */"
    sStr = sStr & "                                  1"
    sStr = sStr & "                               ELSE"
    sStr = sStr & "                                  0"
    sStr = sStr & "                               END END SEL4,"
    sStr = sStr & "                               CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'05|') > 0 THEN          /* ��Ž-����� */"
    sStr = sStr & "                                  1"
    sStr = sStr & "                               ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'55|') > 0 THEN     /* ��Ž-����2 */"
    sStr = sStr & "                                  1"
    sStr = sStr & "                               ELSE"
    sStr = sStr & "                                  0"
    sStr = sStr & "                               END END SEL5,"
    sStr = sStr & "                               CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'06|') > 0 THEN          /* ��Ž-�������� */"
    sStr = sStr & "                                  1"
    sStr = sStr & "                               ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'56|') > 0 THEN     /* ��Ž-ȭ��2 */"
    sStr = sStr & "                                  1"
    sStr = sStr & "                               ELSE"
    sStr = sStr & "                                  0"
    sStr = sStr & "                               END END SEL6,"
    sStr = sStr & "                               CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'07|') > 0 THEN          /* ��Ž-�ѱ����� */"
    sStr = sStr & "                                   1"
    sStr = sStr & "                               ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'57|') > 0 THEN     /* ��Ž-����2 */"
    sStr = sStr & "                                   1"
    sStr = sStr & "                               ELSE"
    sStr = sStr & "                                   0"
    sStr = sStr & "                               END END SEL7,"
    sStr = sStr & "                               CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'08|') > 0 THEN          /* ��Ž-��ġ */"
    sStr = sStr & "                                   1"
    sStr = sStr & "                               ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'58|') > 0 THEN     /* ��Ž-��������2 */"
    sStr = sStr & "                                   1"
    sStr = sStr & "                               ELSE"
    sStr = sStr & "                                   0"
    sStr = sStr & "                               END END SEL8,"
    sStr = sStr & "                               CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'09|') > 0 THEN          /* ��Ž-��ȸ��ȭ */"
    sStr = sStr & "                                   1"
    sStr = sStr & "                               ELSE"
    sStr = sStr & "                                   0"
    sStr = sStr & "                               END SEL9,"
    sStr = sStr & "                               CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'10|') > 0 THEN          /* ��Ž-������ȸ */"
    sStr = sStr & "                                   1"
    sStr = sStr & "                               ELSE"
    sStr = sStr & "                                   0"
    sStr = sStr & "                               END SEL10,"
    sStr = sStr & "                               CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'11|') > 0 THEN          /* ��Ž-�������� */"
    sStr = sStr & "                                   1"
    sStr = sStr & "                               ELSE"
    sStr = sStr & "                                   0"
    sStr = sStr & "                               END SEL11, "
    sStr = sStr & "                               CL_CLOSE "
    
    sStr = sStr & "                          FROM CLTTL01TB"
    sStr = sStr & "                         WHERE ACID   = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "                           AND CL_CLOSE IS NULL "
    
    sStr = sStr & "                           AND SEL_CLASS IN " & sSql_Lsn         '< Ư���� ���븸 ��ȸ
    
    sStr = sStr & "                        )"
    sStr = sStr & "                    WHERE GAEYUL_CD = '" & Trim(txtKaeyol.Text) & "'"
    sStr = sStr & "                   )"
    sStr = sStr & "              GROUP BY LSNCD"
    sStr = sStr & "              HAVING LSNCD"
    sStr = sStr & "                  IN (SELECT LSNCD"
    sStr = sStr & "                        FROM SDLSN01TB"
    sStr = sStr & "                       WHERE ACID    = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "                         AND KAEYOL  = '" & Trim(txtKaeyol.Text) & "'"
    If Trim(Right(cboLsnType.Text, 30)) <> "ALL" Then
        sStr = sStr & "                     AND LSNTYPE = '" & Trim(txtLsnType.Text) & "'"
    End If
    sStr = sStr & "                     )"
    sStr = sStr & "           )"
    sStr = sStr & "      ORDER BY LSNNM "
    
    
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
    
'    '>> �п�
'        sTmp = Trim(basModule.SchCD)
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
'    '>> �迭
'        sTmp = Trim(Right(cboKaeyol.Text, 30))
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
'    '>> �п�
'        sTmp = Trim(basModule.SchCD)
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
'    '>> �迭
'        sTmp = Trim(Right(cboKaeyol.Text, 30))
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
'    '>> �� ����
'        If Trim(Right(cboLsnType.Text, 30)) <> "ALL" Then
'            sTmp = Trim(Right(cboKaeyol.Text, 30))
'            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'                Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
'        End If
        
        
    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic���·� ���ԵǸ� record count�� �� �� ����.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop
    
    With DBRec
        If .RecordCount > 0 Then
            .MoveFirst
            
            
            For nRec = 1 To .RecordCount Step 1
            
                nRet = nRet + 1
                
                sprCopyLsn.MaxRows = sprCopyLsn.MaxRows + 1
                sprCopyLsn.Row = sprCopyLsn.MaxRows
                
                sprCopyLsn.Col = 1
                    sTmp = " ": If IsNull(.Fields("LSNCD")) = False Then sTmp = Trim(.Fields("LSNCD"))
                        Call basFunction.Set_SprType_Text(sprCopyLsn, "CENTER", "LEFT", LenB(sTmp), sTmp)
                sprCopyLsn.Col = sprCopyLsn.Col + 1
                    sTmp = " ": If IsNull(.Fields("LSNNM")) = False Then sTmp = Trim(.Fields("LSNNM"))
                        Call basFunction.Set_SprType_Text(sprCopyLsn, "CENTER", "LEFT", LenB(sTmp), sTmp)
                    
                '## ���ο�
                sprCopyLsn.Col = sprCopyLsn.Col + 1:    nTmp = 0
                    If IsNull(.Fields("S_LSN")) = False Then
                        nTmp = CDbl(.Fields("S_LSN"))
                    End If
                    Call basFunction.Set_SprType_Numeric(sprCopyLsn, 0, 0, 99999, ",", nTmp)
                    
                    
                '<< �ι��ڿ� ���� : 8 ���� >>
                For nCol = 1 To 8 Step 1
                    sprCopyLsn.Col = sprCopyLsn.Col + 1:    nTmp = 0
                    siTem = "SEL" & Trim(CStr(nCol))
                    
                    If IsNull(.Fields(siTem)) = False Then nTmp = CDbl(.Fields(siTem))
                    Call basFunction.Set_SprType_Numeric(sprCopyLsn, 0, 0, 99999, "", nTmp)
                Next nCol
                
                
                Select Case Trim(.Fields("KAEYOL"))
                    Case "01", "03"
                        '��Ž�� 9~11
                        For nCol = 9 To 11 Step 1
                            sprCopyLsn.Col = sprCopyLsn.Col + 1:    nTmp = 0
                            siTem = "SEL" & Trim(CStr(nCol))
                            
                            If IsNull(.Fields(siTem)) = False Then nTmp = CDbl(.Fields(siTem))
                            Call basFunction.Set_SprType_Numeric(sprCopyLsn, 0, 0, 99999, "", nTmp)
                            
                        Next nCol
                        
                    Case "02"
                        '��Ž�� SKIP
                End Select
                
                
                .MoveNext       '<< �����׸�
                
            Next nRec
            
            sprCopyLsn.Row = 1:       sprCopyLsn.Row2 = sprCopyLsn.MaxRows
            sprCopyLsn.Col = 1:       sprCopyLsn.Col2 = sprCopyLsn.MaxCols
            sprCopyLsn.BlockMode = True
                sprCopyLsn.BackColor = basModule.WhiteColor
                sprCopyLsn.BackColorStyle = BackColorStyleUnderGrid
            sprCopyLsn.BlockMode = False

        '>> spread lock
            sprCopyLsn.Row = 1:       sprCopyLsn.Row2 = sprCopyLsn.MaxRows
            sprCopyLsn.Col = 1:       sprCopyLsn.Col2 = sprCopyLsn.MaxCols
            sprCopyLsn.BlockMode = True
                sprCopyLsn.Lock = True
                sprCopyLsn.Protect = True
            sprCopyLsn.BlockMode = False
            
        End If
    End With
    
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    Exit Sub
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    On Error GoTo 0
    MsgBox "�۾������ ��ȸ�� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "�۾������ ��ȸ"
    
End Sub








'## �ѹݿ��� ��� hit �ϴ� ���
Private Sub Gwamok_Matching()
    
    Dim nRow_STD                As Long
    
    
    Dim sSugang_SubjCD          As String
    Dim sSugang_LsnCD           As String
    Dim sSugang_LsnNM           As String
    
    Dim nCol_Gwamok             As Long
    Dim nRow_Gwamok             As Long
    
    Dim sGwamok_LsnCD           As String
    Dim sGwamok_LsnNM           As String
    
    Dim sGwamok_SubjCD          As String
    
    Dim nCol_STD                As Long
    Dim sTmpBlank               As String       '< blank
    
    Dim nLsn_Hit_Sum            As Long         '< �� ���� ���� �� : 4���� �Ϸ�
    
    Dim nRow_CopyLsn            As Long
    Dim nCol_CopyLsn            As Long
    
    Dim nMinusGwamokinWon       As Long
    
    Dim bMinusinWon_CopyLsn     As Boolean
    
    Dim ni                      As Integer
    Dim sSelLsn()               As String
    
    Dim nTotal_ORD_GwamokSu     As Integer
    
    
    '< ��ü ��ϵ� ���� �ʱ�ȭ >
    For nRow_STD = 1 To sprSTD.MaxRows Step 1
        sprSTD.Row = nRow_STD
        sprSTD.Col = sprSTD.MaxCols
        If Trim(sprSTD.Text) = "" Then
            For nCol_STD = 15 To 22 Step 1
                sprSTD.Row = nRow_STD
                
                sprSTD.Col = nCol_STD
                    sTmpBlank = " "
                    Call basFunction.Set_SprType_Text(sprSTD, "center", "left", 1, sTmpBlank)
            Next nCol_STD
            
            sprSTD.Col = 4
                sprSTD.Value = 0
        End If
    Next nRow_STD
    
    
    For nRow_STD = 1 To sprSTD.MaxRows Step 1
        
        nLsn_Hit_Sum = 0            '< �� ���� : �� �л��� ���� 4�� ���;� ��� �´� ����.
        
        sprSTD.Row = nRow_STD
        sprSTD.Col = sprSTD.MaxCols
        
        If Trim(sprSTD.Text) = "" Then          '< �������� ���� �л��� ���ؼ��� ó��
            
            sSugang_SubjCD = ""
            
            nTotal_ORD_GwamokSu = 0             '< �����
            
            sprSTD.Row = nRow_STD
            sprSTD.Col = 7:       sSugang_SubjCD = sSugang_SubjCD & Trim(sprSTD.Text) & "|"
                If Trim(sprSTD.Text) > " " Then nTotal_ORD_GwamokSu = nTotal_ORD_GwamokSu + 1
                
            sprSTD.Col = 8:       sSugang_SubjCD = sSugang_SubjCD & Trim(sprSTD.Text) & "|"
                If Trim(sprSTD.Text) > " " Then nTotal_ORD_GwamokSu = nTotal_ORD_GwamokSu + 1
                
            sprSTD.Col = 9:       sSugang_SubjCD = sSugang_SubjCD & Trim(sprSTD.Text) & "|"
                If Trim(sprSTD.Text) > " " Then nTotal_ORD_GwamokSu = nTotal_ORD_GwamokSu + 1
                
            sprSTD.Col = 10:      sSugang_SubjCD = sSugang_SubjCD & Trim(sprSTD.Text) & "|"
                If Trim(sprSTD.Text) > " " Then nTotal_ORD_GwamokSu = nTotal_ORD_GwamokSu + 1
            
            ' sSugang_SubjCD ���� �л��� ��û��
            sprSTD.Col = 5:       sSugang_LsnCD = Trim(sprSTD.Text)
            
            
            '1. �̵����� �ð�ǥ���� ���� ã��.
            '   �ݳ��� ��� �´����� ���� ���� üũ�Ѵ�.
            For nCol_Gwamok = 1 To sprGwamok.MaxCols Step 1
                
                sprGwamok.Col = nCol_Gwamok
                
                sprGwamok.Row = SpreadHeader:       sGwamok_LsnCD = Trim(sprGwamok.Text)
                sprGwamok.Row = SpreadHeader + 1:   sGwamok_LsnNM = Trim(sprGwamok.Text)
                    
                    
                If StrComp(sSugang_LsnCD, sGwamok_LsnCD, vbTextCompare) = "0" Then
                    
                    For nRow_Gwamok = 1 To sprGwamok.MaxRows Step 1
                        
                        sprGwamok.Row = nRow_Gwamok
                        sprGwamok.Col = nCol_Gwamok
                        
                            sGwamok_SubjCD = Trim(Right(sprGwamok.Text, 30))
                            
                        If InStr(1, sSugang_SubjCD, sGwamok_SubjCD, vbTextCompare) > 0 Then
                            
                            For nCol_STD = 7 To 10 Step 1
                                
                                sprSTD.Col = nCol_STD
                                If StrComp(Trim(sprSTD.Text), sGwamok_SubjCD, vbTextCompare) = 0 Then
                                    
                                    sprSTD.Col = nCol_STD + 8
                                        Call basFunction.Set_SprType_Text(sprSTD, "center", "left", basFunction.LenKor(sGwamok_LsnCD), sGwamok_LsnCD)
                                    sprSTD.Col = nCol_STD + 12
                                        Call basFunction.Set_SprType_Text(sprSTD, "center", "left", basFunction.LenKor(sGwamok_LsnNM), sGwamok_LsnNM)
                                    
                                    
                                    nLsn_Hit_Sum = nLsn_Hit_Sum + 1         '<< hit����
                                    
                                End If
                            Next nCol_STD
                            
                            
                        End If
                        
                    Next nRow_Gwamok
                    
                End If
            Next nCol_Gwamok
            
            
            
            '## �ο��� ���� -------------------------------------------------------------------------------------------------------
            
            bMinusinWon_CopyLsn = False
            ReDim sSelLsn(4) As String
            
            If nLsn_Hit_Sum = nTotal_ORD_GwamokSu Then          '<< ��û������� ��
            
                For nCol_STD = 1 To 4 Step 1
                    
                    sprSTD.Col = 15 + nCol_STD - 1:      sGwamok_LsnCD = Trim(sprSTD.Text)              '< ��û���� ��������.
                    sprSTD.Col = 7 + nCol_STD - 1:       sGwamok_SubjCD = Trim(sprSTD.Text)
                    
                    For nRow_CopyLsn = 1 To sprCopyLsn.MaxRows Step 1
                        sprCopyLsn.Row = nRow_CopyLsn
                        sprCopyLsn.Col = 1
                        
                        
                        If StrComp(sGwamok_LsnCD, Trim(sprCopyLsn.Text), vbTextCompare) = 0 Then        '< ���� �� ã�� (���� spread����)
                            
                            '## �ݺ��ο� ó�� ����
                            For ni = 1 To 4 Step 1
                                If StrComp(sSelLsn(ni), Trim(sprCopyLsn.Text), vbTextCompare) = 0 Then
                                    Exit For
                                ElseIf StrComp(sSelLsn(ni), "", vbTextCompare) = 0 Then
                                    sSelLsn(ni) = Trim(sprCopyLsn.Text)                                 '< �� ���
                                    Exit For
                                End If
                            Next ni
                            
                            
                            For nCol_CopyLsn = 4 To sprCopyLsn.MaxCols Step 1
                                sprCopyLsn.Row = SpreadHeader
                                sprCopyLsn.Col = nCol_CopyLsn
                                
                                If StrComp(sGwamok_SubjCD, Trim(sprCopyLsn.Text), vbTextCompare) = 0 Then
                                
                                    nMinusGwamokinWon = 0           '< �����ο�
    
                                    sprCopyLsn.Row = nRow_CopyLsn
                                    sprCopyLsn.Col = nCol_CopyLsn:      nMinusGwamokinWon = sprCopyLsn.Value - 1
                                    
                                    If nMinusGwamokinWon >= 0 Then                   '<<  (�ο� > 0) �̸� ������.
                                        sprCopyLsn.Row = nRow_CopyLsn
                                        
                                        sprCopyLsn.Col = nCol_CopyLsn:      sprCopyLsn.Value = sprCopyLsn.Value - 1
                                        
                                        bMinusinWon_CopyLsn = True                  '< bMinusinWon_CopyLsn = true �� �Ǹ� OK
                                        GoTo Next_Statement
                                    Else
                                        bMinusinWon_CopyLsn = False
                                    End If
                                    
                                End If
                                
                            Next nCol_CopyLsn
                        End If
                    Next nRow_CopyLsn
                Next nCol_STD
                
                
                '## �� ó��
Next_Statement:
                
                For nRow_CopyLsn = 1 To sprCopyLsn.MaxRows Step 1
                    sprCopyLsn.Row = nRow_CopyLsn
                    sprCopyLsn.Col = 1
                    
                    For ni = 1 To 4 Step 1
                        If StrComp(Trim(sprCopyLsn.Text), sSelLsn(ni), vbTextCompare) = 0 Then
                            sprCopyLsn.Row = nRow_CopyLsn
                            sprCopyLsn.Col = 3
                                sprCopyLsn.Value = sprCopyLsn.Value + 1
                                
                            Exit For
                        End If
                    Next ni
                Next nRow_CopyLsn
            
                                        
                
                
            End If
            
            
            
            If bMinusinWon_CopyLsn = True Then
                sprSTD.Col = 4
                    sprSTD.Value = 1
            Else
                For nCol_Gwamok = 15 To 22 Step 1               '<< 4������ 1�����̶� ��߳��� �������
                    sprSTD.Col = nCol_Gwamok:       sprSTD.Text = ""
                Next nCol_Gwamok
                sprSTD.Col = 4:     sprSTD.Value = 0            '<< ó������ ���� ��Ÿ��
                
                
                
                '#########################################################################################
                '## 1���� ������ ��Ī���� �����Ƿ�, �����ؼ� �������ɿ��θ� �Ǵ��ؾ� �Ѵ�.
                '#########################################################################################
                    Call Compound_Matching(nRow_STD)
                    
                '#########################################################################################
                
                
            End If
            '---------------------------------------------------------------------------------------------------------------------
            
            
        End If
    Next nRow_STD
    
End Sub

Private Sub Compound_Matching(ByVal aRow_STD As Long)
    
    Dim nRow_STD                As Long
    Dim nCol_STD                As Long
    
    Dim nRow_Gwamok             As Long
    Dim nCol_Gwamok             As Long
    
    Dim nAdd_Row                As Long
    Dim nTmp_Row                As Long
    
    Dim sGwamok1                As String
    Dim sGwamok2                As String
    Dim sGwamok3                As String
    Dim sGwamok4                As String
    
    Dim sTmpGwamok1             As String
    Dim sTmpGwamok2             As String
    Dim sTmpGwamok3             As String
    Dim sTmpGwamok4             As String
    
    Dim sTmpLsnCD1              As String
    Dim sTmpLsnCD2              As String
    Dim sTmpLsnCD3              As String
    Dim sTmpLsnCD4              As String
    
    Dim sTmpLsnNM1              As String
    Dim sTmpLsnNM2              As String
    Dim sTmpLsnNM3              As String
    Dim sTmpLsnNM4              As String
    
    Dim nRow_CopyLsn            As Long
    Dim nCol_CopyLsn            As Long
    
    Dim bHit                    As Boolean
    
    Dim sGwamok_LsnCD           As String
    Dim sGwamok_SubjCD          As String

    Dim nMinusGwamokinWon       As Long
    Dim bMinusinWon_CopyLsn     As Boolean
    Dim sSelLsn()               As String
    Dim ni                      As Integer
    Dim nTmp                    As Long
    
    Dim nTotal_ORD_GwamokSu     As Integer
    Dim nAcc_ORD_GwamokSu       As Integer
    
    sprSTD.Row = aRow_STD
    
    For nRow_Gwamok = 0 To 3 Step 1
        
        nTotal_ORD_GwamokSu = 0     '< �л� ��û�����
        nAcc_ORD_GwamokSu = 0
       
        sprSTD.Col = 7:         sGwamok1 = Trim(sprSTD.Text):   If sGwamok1 > " " Then nTotal_ORD_GwamokSu = nTotal_ORD_GwamokSu + 1
        sprSTD.Col = 8:         sGwamok2 = Trim(sprSTD.Text):   If sGwamok2 > " " Then nTotal_ORD_GwamokSu = nTotal_ORD_GwamokSu + 1
        sprSTD.Col = 9:         sGwamok3 = Trim(sprSTD.Text):   If sGwamok3 > " " Then nTotal_ORD_GwamokSu = nTotal_ORD_GwamokSu + 1
        sprSTD.Col = 10:        sGwamok4 = Trim(sprSTD.Text):   If sGwamok4 > " " Then nTotal_ORD_GwamokSu = nTotal_ORD_GwamokSu + 1
        
        sTmpLsnCD1 = ""
        sTmpLsnCD2 = ""
        sTmpLsnCD3 = ""
        sTmpLsnCD4 = ""
        
        bHit = False
    
        nTmp_Row = nRow_Gwamok Mod 4:       nAdd_Row = nTmp_Row + 1
            sTmpLsnCD1 = "":        sTmpLsnNM1 = ""
            For nCol_Gwamok = 1 To sprGwamok.MaxCols Step 1
                sprGwamok.Row = nAdd_Row
                sprGwamok.Col = nCol_Gwamok
                    sTmpGwamok1 = Trim(Right(sprGwamok.Text, 30))
                    
                If StrComp(sGwamok1, sTmpGwamok1, vbTextCompare) = 0 Then
                    sprGwamok.Row = SpreadHeader
                    sprGwamok.Col = nCol_Gwamok
                        sTmpLsnCD1 = Trim(sprGwamok.Text)                       '< ������ ������ ���� �ִ´�.
                    
                    sprGwamok.Row = SpreadHeader + 1
                    sprGwamok.Col = nCol_Gwamok
                        sTmpLsnNM1 = Trim(sprGwamok.Text)
                    
                    For nRow_CopyLsn = 1 To sprCopyLsn.MaxRows Step 1
                    
'                        sprCopyLsn.Row = nRow_CopyLsn
'                        sprCopyLsn.Col = 1
'                        If StrComp(sTmpLsnCD1, Trim(sprCopyLsn.Text), vbTextCompare) = 0 Or sTmpLsnCD1 > "90000" Then       '< ���ο� ����� SPREAD���� ��(LSNCD) ã��
                        
                            For nCol_CopyLsn = 4 To sprCopyLsn.MaxCols Step 1
                                sprCopyLsn.Row = SpreadHeader
                                sprCopyLsn.Col = nCol_CopyLsn
                                
                                If StrComp(sGwamok1, Trim(sprCopyLsn.Text), vbTextCompare) = 0 Then     '< ���ο� ����� SPREAD���� ����(GWAMOK) ã��
                                    sprCopyLsn.Row = nRow_CopyLsn
                                    sprCopyLsn.Col = nCol_CopyLsn
                                        nTmp = sprCopyLsn.Value - 1
                                    
                                    If nTmp >= 0 Then
                                        '# 1
                                        sprSTD.Row = aRow_STD:  sprSTD.Col = 15:    Call basFunction.Set_SprType_Text(sprSTD, "center", "left", basFunction.LenKor(sTmpLsnCD1), sTmpLsnCD1)
                                        sprSTD.Row = aRow_STD:  sprSTD.Col = 19:    Call basFunction.Set_SprType_Text(sprSTD, "center", "left", basFunction.LenKor(sTmpLsnNM1), sTmpLsnNM1)
                                        
                                        nAcc_ORD_GwamokSu = nAcc_ORD_GwamokSu + 1
                                        GoTo Next_Gwamok1           '< �ο������� ��� �����ؾ� OK
                                    Else
                                        sTmpLsnCD1 = ""
                                        sTmpLsnNM1 = ""
                                    End If
                                End If
                            Next nCol_CopyLsn
                            
'                        End If
                        
                    Next nRow_CopyLsn
                    
                    sTmpLsnCD1 = ""
                    sTmpLsnNM1 = ""
                    
                End If
            Next nCol_Gwamok
            
Next_Gwamok1:
        nTmp_Row = nAdd_Row Mod 4:          nAdd_Row = nTmp_Row + 1
            sTmpLsnCD2 = "":        sTmpLsnNM2 = ""
            For nCol_Gwamok = 1 To sprGwamok.MaxCols Step 1
                sprGwamok.Row = nAdd_Row
                sprGwamok.Col = nCol_Gwamok
                    sTmpGwamok2 = Trim(Right(sprGwamok.Text, 30))
                    
                If StrComp(sGwamok2, sTmpGwamok2, vbTextCompare) = 0 Then
                    sprGwamok.Row = SpreadHeader
                    sprGwamok.Col = nCol_Gwamok
                        sTmpLsnCD2 = Trim(sprGwamok.Text)                       '< ������ ������ ���� �ִ´�.
                    
                    sprGwamok.Row = SpreadHeader + 1
                    sprGwamok.Col = nCol_Gwamok
                        sTmpLsnNM2 = Trim(sprGwamok.Text)
                    
                    For nRow_CopyLsn = 1 To sprCopyLsn.MaxRows Step 1
                    
'                        sprCopyLsn.Row = nRow_CopyLsn
'                        sprCopyLsn.Col = 1
'                        If StrComp(sTmpLsnCD2, Trim(sprCopyLsn.Text), vbTextCompare) = 0 Or sTmpLsnCD2 > "90000" Then         '< ���ο� ����� SPREAD���� ��(LSNCD) ã��
                        
                            For nCol_CopyLsn = 4 To sprCopyLsn.MaxCols Step 1
                                sprCopyLsn.Row = SpreadHeader
                                sprCopyLsn.Col = nCol_CopyLsn
                                
                                If StrComp(sGwamok2, Trim(sprCopyLsn.Text), vbTextCompare) = 0 Then     '< ���ο� ����� SPREAD���� ����(GWAMOK) ã��
                                    sprCopyLsn.Row = nRow_CopyLsn
                                    sprCopyLsn.Col = nCol_CopyLsn
                                        nTmp = sprCopyLsn.Value - 1
                                    
                                    If nTmp >= 0 Then
                                        '# 2
                                        sprSTD.Row = aRow_STD:  sprSTD.Col = 16:    Call basFunction.Set_SprType_Text(sprSTD, "center", "left", basFunction.LenKor(sTmpLsnCD2), sTmpLsnCD2)
                                        sprSTD.Row = aRow_STD:  sprSTD.Col = 20:    Call basFunction.Set_SprType_Text(sprSTD, "center", "left", basFunction.LenKor(sTmpLsnNM2), sTmpLsnNM2)
                                        
                                        nAcc_ORD_GwamokSu = nAcc_ORD_GwamokSu + 1
                                        GoTo Next_Gwamok2           '< �ο������� ��� �����ؾ� OK
                                    Else
                                        sTmpLsnCD2 = ""
                                        sTmpLsnNM2 = ""
                                    End If
                                End If
                            Next nCol_CopyLsn
                            
'                        End If
                        
                    Next nRow_CopyLsn
                    
                    sTmpLsnCD2 = ""
                    sTmpLsnNM2 = ""
                    
                End If
                
            Next nCol_Gwamok
            
Next_Gwamok2:
        nTmp_Row = nAdd_Row Mod 4:          nAdd_Row = nTmp_Row + 1
            sTmpLsnCD3 = "":        sTmpLsnNM3 = ""
            For nCol_Gwamok = 1 To sprGwamok.MaxCols Step 1
                sprGwamok.Row = nAdd_Row
                sprGwamok.Col = nCol_Gwamok
                    sTmpGwamok3 = Trim(Right(sprGwamok.Text, 30))
                    
                If StrComp(sGwamok3, sTmpGwamok3, vbTextCompare) = 0 Then
                    sprGwamok.Row = SpreadHeader
                    sprGwamok.Col = nCol_Gwamok
                        sTmpLsnCD3 = Trim(sprGwamok.Text)                       '< ������ ������ ���� �ִ´�.
                    
                    sprGwamok.Row = SpreadHeader + 1
                    sprGwamok.Col = nCol_Gwamok
                        sTmpLsnNM3 = Trim(sprGwamok.Text)
                    
                    For nRow_CopyLsn = 1 To sprCopyLsn.MaxRows Step 1
                    
'                        sprCopyLsn.Row = nRow_CopyLsn
'                        sprCopyLsn.Col = 1
'                        If StrComp(sTmpLsnCD3, Trim(sprCopyLsn.Text), vbTextCompare) = 0 Or sTmpLsnCD3 > "90000" Then         '< ���ο� ����� SPREAD���� ��(LSNCD) ã��
                        
                            For nCol_CopyLsn = 4 To sprCopyLsn.MaxCols Step 1
                                sprCopyLsn.Row = SpreadHeader
                                sprCopyLsn.Col = nCol_CopyLsn
                                
                                If StrComp(sGwamok3, Trim(sprCopyLsn.Text), vbTextCompare) = 0 Then     '< ���ο� ����� SPREAD���� ����(GWAMOK) ã��
                                    sprCopyLsn.Row = nRow_CopyLsn
                                    sprCopyLsn.Col = nCol_CopyLsn
                                        nTmp = sprCopyLsn.Value - 1
                                    
                                    If nTmp >= 0 Then
                                        '# 3
                                        sprSTD.Row = aRow_STD:  sprSTD.Col = 17:    Call basFunction.Set_SprType_Text(sprSTD, "center", "left", basFunction.LenKor(sTmpLsnCD3), sTmpLsnCD3)
                                        sprSTD.Row = aRow_STD:  sprSTD.Col = 21:    Call basFunction.Set_SprType_Text(sprSTD, "center", "left", basFunction.LenKor(sTmpLsnNM3), sTmpLsnNM3)
                                        
                                        nAcc_ORD_GwamokSu = nAcc_ORD_GwamokSu + 1
                                        GoTo Next_Gwamok3           '< �ο������� ��� �����ؾ� OK
                                    Else
                                        sTmpLsnCD3 = ""
                                        sTmpLsnNM3 = ""
                                    End If
                                End If
                            Next nCol_CopyLsn
                            
'                        End If
                        
                    Next nRow_CopyLsn
                    
                    sTmpLsnCD3 = ""
                    sTmpLsnNM3 = ""
                    
                End If
                
            Next nCol_Gwamok
        
Next_Gwamok3:
        nTmp_Row = nAdd_Row Mod 4:          nAdd_Row = nTmp_Row + 1
            sTmpLsnCD4 = "":        sTmpLsnNM4 = ""
            For nCol_Gwamok = 1 To sprGwamok.MaxCols Step 1
                sprGwamok.Row = nAdd_Row
                sprGwamok.Col = nCol_Gwamok
                    sTmpGwamok4 = Trim(Right(sprGwamok.Text, 30))
                    
                If StrComp(sGwamok4, sTmpGwamok4, vbTextCompare) = 0 Then
                    sprGwamok.Row = SpreadHeader
                    sprGwamok.Col = nCol_Gwamok
                        sTmpLsnCD4 = Trim(sprGwamok.Text)                       '< ������ ������ ���� �ִ´�.
                    
                    sprGwamok.Row = SpreadHeader + 1
                    sprGwamok.Col = nCol_Gwamok
                        sTmpLsnNM4 = Trim(sprGwamok.Text)
                    
                    For nRow_CopyLsn = 1 To sprCopyLsn.MaxRows Step 1
                    
'                        sprCopyLsn.Row = nRow_CopyLsn
'                        sprCopyLsn.Col = 1
'                        If StrComp(sTmpLsnCD4, Trim(sprCopyLsn.Text), vbTextCompare) = 0 Or sTmpLsnCD4 > "90000" Then         '< ���ο� ����� SPREAD���� ��(LSNCD) ã��
                        
                            For nCol_CopyLsn = 4 To sprCopyLsn.MaxCols Step 1
                                sprCopyLsn.Row = SpreadHeader
                                sprCopyLsn.Col = nCol_CopyLsn
                                
                                If StrComp(sGwamok4, Trim(sprCopyLsn.Text), vbTextCompare) = 0 Then     '< ���ο� ����� SPREAD���� ����(GWAMOK) ã��
                                    sprCopyLsn.Row = nRow_CopyLsn
                                    sprCopyLsn.Col = nCol_CopyLsn
                                        nTmp = sprCopyLsn.Value - 1
                                    
                                    If nTmp >= 0 Then
                                        '# 4
                                        sprSTD.Row = aRow_STD:  sprSTD.Col = 18:    Call basFunction.Set_SprType_Text(sprSTD, "center", "left", basFunction.LenKor(sTmpLsnCD4), sTmpLsnCD4)
                                        sprSTD.Row = aRow_STD:  sprSTD.Col = 22:    Call basFunction.Set_SprType_Text(sprSTD, "center", "left", basFunction.LenKor(sTmpLsnNM4), sTmpLsnNM4)
                    
                                        nAcc_ORD_GwamokSu = nAcc_ORD_GwamokSu + 1
                                        GoTo Next_Gwamok4           '< �ο������� ��� �����ؾ� OK
                                    Else
                                        sTmpLsnCD4 = ""
                                        sTmpLsnNM4 = ""
                                    End If
                                End If
                            Next nCol_CopyLsn
                            
'                        End If
                        
                    Next nRow_CopyLsn
                    
                    sTmpLsnCD4 = ""
                    sTmpLsnNM4 = ""
                    
                End If
                
            Next nCol_Gwamok
        
Next_Gwamok4:

        If nAcc_ORD_GwamokSu = nTotal_ORD_GwamokSu Then
            
            'no action : ��� ����
            
            bHit = True
            Exit For
            
        Else
            '/* �� ��θ� �������� �����Ƿ� �н� */
            sprSTD.Row = aRow_STD:  sprSTD.Col = 15:    Call basFunction.Set_SprType_Text(sprSTD, "center", "left", 1, " ")
            sprSTD.Row = aRow_STD:  sprSTD.Col = 16:    Call basFunction.Set_SprType_Text(sprSTD, "center", "left", 1, " ")
            sprSTD.Row = aRow_STD:  sprSTD.Col = 17:    Call basFunction.Set_SprType_Text(sprSTD, "center", "left", 1, " ")
            sprSTD.Row = aRow_STD:  sprSTD.Col = 18:    Call basFunction.Set_SprType_Text(sprSTD, "center", "left", 1, " ")
            
            sprSTD.Row = aRow_STD:  sprSTD.Col = 19:    Call basFunction.Set_SprType_Text(sprSTD, "center", "left", 1, " ")
            sprSTD.Row = aRow_STD:  sprSTD.Col = 20:    Call basFunction.Set_SprType_Text(sprSTD, "center", "left", 1, " ")
            sprSTD.Row = aRow_STD:  sprSTD.Col = 21:    Call basFunction.Set_SprType_Text(sprSTD, "center", "left", 1, " ")
            sprSTD.Row = aRow_STD:  sprSTD.Col = 22:    Call basFunction.Set_SprType_Text(sprSTD, "center", "left", 1, " ")
            
            sTmpLsnCD1 = ""
            sTmpLsnCD2 = ""
            sTmpLsnCD3 = ""
            sTmpLsnCD4 = ""
            
        End If
        
    Next nRow_Gwamok
    
    
    
    '## �ο��� ���� -------------------------------------------------------------------------------------------------------\
    
    bMinusinWon_CopyLsn = False
    ReDim sSelLsn(4) As String
    
    If bHit = True Then
        For nCol_STD = 1 To 4 Step 1

            sprSTD.Col = 15 + nCol_STD - 1:      sGwamok_LsnCD = Trim(sprSTD.Text)              '< ��û���� ��������.
            sprSTD.Col = 7 + nCol_STD - 1:       sGwamok_SubjCD = Trim(sprSTD.Text)

            For nRow_CopyLsn = 1 To sprCopyLsn.MaxRows Step 1
                sprCopyLsn.Row = nRow_CopyLsn
                sprCopyLsn.Col = 1


                'If StrComp(sGwamok_LsnCD, Trim(sprCopyLsn.Text), vbTextCompare) = 0 Then        '< ���� �� ã�� (���� spread����)

                    '## �ݺ��ο� ó�� ����
                    For ni = 1 To 4 Step 1
                        If StrComp(sSelLsn(ni), Trim(sprCopyLsn.Text), vbTextCompare) = 0 Then
                            Exit For
                        ElseIf StrComp(sSelLsn(ni), "", vbTextCompare) = 0 Then
                            sSelLsn(ni) = Trim(sprCopyLsn.Text)                                 '< �� ���
                            Exit For
                        End If
                    Next ni


                    For nCol_CopyLsn = 4 To sprCopyLsn.MaxCols Step 1
                        sprCopyLsn.Row = SpreadHeader
                        sprCopyLsn.Col = nCol_CopyLsn

                        If StrComp(sGwamok_SubjCD, Trim(sprCopyLsn.Text), vbTextCompare) = 0 Then

                            nMinusGwamokinWon = 0           '< �����ο�

                            sprCopyLsn.Row = nRow_CopyLsn
                            sprCopyLsn.Col = nCol_CopyLsn:      nMinusGwamokinWon = sprCopyLsn.Value - 1

                            If nMinusGwamokinWon > 0 Then                   '<<  (�ο� > 0) �̸� ������.
                                sprCopyLsn.Row = nRow_CopyLsn

                                sprCopyLsn.Col = nCol_CopyLsn:      sprCopyLsn.Value = sprCopyLsn.Value - 1

                                bMinusinWon_CopyLsn = True                  '< bMinusinWon_CopyLsn = true �� �Ǹ� OK
                                Exit For
                            Else
                                bMinusinWon_CopyLsn = False
                            End If

                        End If

                    Next nCol_CopyLsn
                    
                'End If
                
            Next nRow_CopyLsn
        Next nCol_STD

        '## �� ó��
        For nRow_CopyLsn = 1 To sprCopyLsn.MaxRows Step 1
            sprCopyLsn.Row = nRow_CopyLsn
            sprCopyLsn.Col = 1

            For ni = 1 To 4 Step 1
                If StrComp(Trim(sprCopyLsn.Text), sSelLsn(ni), vbTextCompare) = 0 Then
                    sprCopyLsn.Row = nRow_CopyLsn
                    sprCopyLsn.Col = 3
                        sprCopyLsn.Value = sprCopyLsn.Value + 1

                    Exit For
                End If
            Next ni
        Next nRow_CopyLsn

        If bMinusinWon_CopyLsn = True Then
            sprSTD.Col = 4:     sprSTD.Value = 1
        Else
            For nCol_Gwamok = 15 To 22 Step 1               '<< 4������ 1�����̶� ��߳��� �������
                sprSTD.Col = nCol_Gwamok:       sprSTD.Text = ""
            Next nCol_Gwamok
            sprSTD.Col = 4:     sprSTD.Value = 0            '<< ó������ ���� ��Ÿ��
        End If
        
    End If
            
    
End Sub





'## ����ó������ ����ϱ�
'## update �� �ֽ��ϴ�.
Private Sub cmdStdGwamokSave_Click()
    Dim DBCmd       As ADODB.Command        '<< �л� �� ���� ����ϱ�
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    
    Dim sStr        As String
    Dim nLength     As Long
    Dim nExe        As Long
    
    Dim nAccExe     As Long
    Dim nTotExe     As Long
    
    Dim nRow        As Long
    
    Dim sSchNO      As String
    Dim sGwamok1    As String
    Dim sGwamok2    As String
    Dim sGwamok3    As String
    Dim sGwamok4    As String
    
    On Error GoTo ErrStmt
    
    basDataBase.DBConn.BeginTrans

    Set DBCmd = New ADODB.Command
    Set DBParam = New ADODB.Parameter

    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection

    nTotExe = 0
    nAccExe = 0

    For nRow = 1 To sprSTD.MaxRows Step 1
        
        sprSTD.Row = nRow
        sprSTD.Col = 4
        
        If sprSTD.Value = 1 Then            '< ������ ������
            
            nAccExe = nAccExe + 1
            
            
            '< ���� >
            sprSTD.Row = nRow
            
            sprSTD.Col = 1:         sSchNO = Trim(sprSTD.Text)
            sprSTD.Col = 15:        sGwamok1 = Trim(sprSTD.Text)
            sprSTD.Col = 16:        sGwamok2 = Trim(sprSTD.Text)
            sprSTD.Col = 17:        sGwamok3 = Trim(sprSTD.Text)
            sprSTD.Col = 18:        sGwamok4 = Trim(sprSTD.Text)
            
            
            
            sStr = ""
            sStr = sStr & "  UPDATE CLTTL01TB "
            sStr = sStr & "     SET CL_CLOSE = '" & Mid(Trim(fpCL_Close.UnFmtText), 3, 4) & "', "
            sStr = sStr & "         GWA_BAN1 = '" & sGwamok1 & "', "
            sStr = sStr & "         GWA_BAN2 = '" & sGwamok2 & "', "
            sStr = sStr & "         GWA_BAN3 = '" & sGwamok3 & "', "
            sStr = sStr & "         GWA_BAN4 = '" & sGwamok4 & "'  "
            sStr = sStr & "   WHERE SCHNO  = '" & sSchNO & "'"
            
            
'    '>> test
'        sTmp = ""
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
    
            If nExe = 1 Then
                nTotExe = nTotExe + 1
            End If
            
        End If
        
    Next nRow
    
    If nAccExe = nTotExe Then
        basDataBase.DBConn.CommitTrans
        
        Call cmdFind_STD_Subj_Click
        chkAll.Value = 1
        Call chkAll_Click
        Call cmdFind_LSN_in_STD_Click
        
        MsgBox "����ó�������� ����Ͽ����ϴ�.", vbInformation + vbOKOnly, "����ó������ ���"
        
    Else
        basDataBase.DBConn.RollbackTrans
        MsgBox "����ó�������� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "����ó������ ���"
    End If
    
    
    Set DBRec = Nothing
    Set DBCmd = Nothing
    
    Exit Sub
ErrStmt:
    basDataBase.DBConn.RollbackTrans
    MsgBox "����ó������ ��Ͻ� ������ �߻��Ͽ����ϴ�." & vbCrLf & _
           Trim(CStr(Err.Number)) & " : " & Trim(Err.Description), vbCritical + vbOKOnly, "����ó������ ���"
    
    Set DBRec = Nothing
    Set DBCmd = Nothing
    
    On Error GoTo 0
    
End Sub







Private Sub sprCopyLsn_Click(ByVal Col As Long, ByVal Row As Long)
    If Row < 1 Then Exit Sub
    If Col < 1 Then Exit Sub
    
    
    With sprCopyLsn
        If Trim(.Tag) = "" Then .Tag = "1"
        
        .Row = CLng(.Tag):  .Row2 = .Row
        .Col = 1:           .Col2 = .MaxCols
        .BlockMode = True
            .BackColor = basModule.WhiteColor
            .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
        
        .Row = Row:     .Row2 = .Row
        .Col = 1:       .Col2 = .MaxCols
        .BlockMode = True
            .BackColor = basModule.BackColor2
            .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
        
        .Tag = Trim(CStr(Row))
        
    End With
End Sub


Private Sub sprBaseLsn_Click(ByVal Col As Long, ByVal Row As Long)
    If Row < 1 Then Exit Sub
    If Col < 1 Then Exit Sub
    
    
    With sprBaseLsn
        If Trim(.Tag) = "" Then .Tag = "1"
        
        .Row = CLng(.Tag):  .Row2 = .Row
        .Col = 1:           .Col2 = .MaxCols
        .BlockMode = True
            .BackColor = basModule.WhiteColor
            .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
        
        .Row = Row:     .Row2 = .Row
        .Col = 1:       .Col2 = .MaxCols
        .BlockMode = True
            .BackColor = basModule.BackColor1
            .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
        
        .Tag = Trim(CStr(Row))
        
    End With
End Sub














