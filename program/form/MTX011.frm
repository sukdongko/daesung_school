VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form MTX011 
   Caption         =   "�ð�ǥ ����� >> ������ �ð�ǥ ��� cp"
   ClientHeight    =   10155
   ClientLeft      =   810
   ClientTop       =   2325
   ClientWidth     =   15600
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10155
   ScaleWidth      =   15600
   Begin VB.Frame Frame9 
      BackColor       =   &H00808080&
      BorderStyle     =   0  '����
      Caption         =   "Frame9"
      Height          =   375
      Left            =   3450
      TabIndex        =   55
      Top             =   0
      Width           =   5655
      Begin VB.Frame Frame10 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  '����
         Caption         =   "Frame10"
         Height          =   315
         Left            =   30
         TabIndex        =   56
         Top             =   30
         Width           =   5595
         Begin VB.ComboBox cboKaeyol_All 
            Height          =   300
            Left            =   1980
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   1
            Top             =   0
            Width           =   1005
         End
         Begin VB.Label Label4 
            BackStyle       =   0  '����
            Caption         =   "�迭 ��������"
            Height          =   210
            Left            =   420
            TabIndex        =   57
            Top             =   45
            Width           =   1185
         End
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H0082C8E8&
      BorderStyle     =   0  '����
      Caption         =   "Frame7"
      Height          =   10125
      Left            =   0
      TabIndex        =   37
      Top             =   0
      Width           =   5415
      Begin VB.Frame Frame8 
         BackColor       =   &H00D2EAF5&
         BorderStyle     =   0  '����
         Caption         =   "Frame8"
         Height          =   10065
         Left            =   30
         TabIndex        =   38
         Top             =   30
         Width           =   5355
         Begin VB.Frame FLine 
            BackColor       =   &H00808080&
            BorderStyle     =   0  '����
            ForeColor       =   &H00FF8080&
            Height          =   9660
            Index           =   19
            Left            =   4620
            TabIndex        =   54
            Top             =   360
            Width           =   30
         End
         Begin VB.CommandButton cmdFindAll 
            Caption         =   "��ü ������ �ð� ��ȸ"
            Height          =   345
            Left            =   450
            TabIndex        =   0
            Top             =   0
            Width           =   2295
         End
         Begin VB.Frame FLine 
            BackColor       =   &H00808080&
            BorderStyle     =   0  '����
            ForeColor       =   &H000000C0&
            Height          =   30
            Index           =   0
            Left            =   60
            TabIndex        =   52
            Top             =   1920
            Width           =   5235
         End
         Begin VB.Frame FLine 
            BackColor       =   &H00808080&
            BorderStyle     =   0  '����
            ForeColor       =   &H000000C0&
            Height          =   30
            Index           =   1
            Left            =   60
            TabIndex        =   51
            Top             =   2820
            Width           =   5235
         End
         Begin VB.Frame FLine 
            BackColor       =   &H00808080&
            BorderStyle     =   0  '����
            ForeColor       =   &H000000C0&
            Height          =   30
            Index           =   3
            Left            =   60
            TabIndex        =   50
            Top             =   3720
            Width           =   5235
         End
         Begin VB.Frame FLine 
            BackColor       =   &H00808080&
            BorderStyle     =   0  '����
            ForeColor       =   &H000000C0&
            Height          =   30
            Index           =   4
            Left            =   60
            TabIndex        =   49
            Top             =   4620
            Width           =   5235
         End
         Begin VB.Frame FLine 
            BackColor       =   &H00808080&
            BorderStyle     =   0  '����
            ForeColor       =   &H000000C0&
            Height          =   30
            Index           =   6
            Left            =   60
            TabIndex        =   48
            Top             =   5520
            Width           =   5235
         End
         Begin VB.Frame FLine 
            BackColor       =   &H00808080&
            BorderStyle     =   0  '����
            ForeColor       =   &H000000C0&
            Height          =   30
            Index           =   7
            Left            =   60
            TabIndex        =   47
            Top             =   6420
            Width           =   5235
         End
         Begin VB.Frame FLine 
            BackColor       =   &H00808080&
            BorderStyle     =   0  '����
            ForeColor       =   &H000000C0&
            Height          =   30
            Index           =   8
            Left            =   60
            TabIndex        =   46
            Top             =   7320
            Width           =   5235
         End
         Begin VB.Frame FLine 
            BackColor       =   &H00808080&
            BorderStyle     =   0  '����
            ForeColor       =   &H000000C0&
            Height          =   30
            Index           =   9
            Left            =   60
            TabIndex        =   45
            Top             =   8220
            Width           =   5235
         End
         Begin VB.Frame FLine 
            BackColor       =   &H00808080&
            BorderStyle     =   0  '����
            ForeColor       =   &H000000C0&
            Height          =   30
            Index           =   10
            Left            =   60
            TabIndex        =   44
            Top             =   9120
            Width           =   5235
         End
         Begin VB.Frame FLine 
            BackColor       =   &H00808080&
            BorderStyle     =   0  '����
            ForeColor       =   &H00FF8080&
            Height          =   9660
            Index           =   11
            Left            =   1320
            TabIndex        =   43
            Top             =   360
            Width           =   30
         End
         Begin VB.Frame FLine 
            BackColor       =   &H00808080&
            BorderStyle     =   0  '����
            ForeColor       =   &H00FF8080&
            Height          =   9660
            Index           =   12
            Left            =   1980
            TabIndex        =   42
            Top             =   360
            Width           =   30
         End
         Begin VB.Frame FLine 
            BackColor       =   &H00808080&
            BorderStyle     =   0  '����
            ForeColor       =   &H00FF8080&
            Height          =   9660
            Index           =   16
            Left            =   2640
            TabIndex        =   41
            Top             =   360
            Width           =   30
         End
         Begin VB.Frame FLine 
            BackColor       =   &H00808080&
            BorderStyle     =   0  '����
            ForeColor       =   &H00FF8080&
            Height          =   9660
            Index           =   17
            Left            =   3300
            TabIndex        =   40
            Top             =   360
            Width           =   30
         End
         Begin VB.Frame FLine 
            BackColor       =   &H00808080&
            BorderStyle     =   0  '����
            ForeColor       =   &H00FF8080&
            Height          =   9660
            Index           =   18
            Left            =   3960
            TabIndex        =   39
            Top             =   360
            Width           =   30
         End
         Begin FPSpread.vaSpread sprTrx_T 
            Height          =   9675
            Left            =   30
            TabIndex        =   2
            Top             =   360
            Width           =   5295
            _Version        =   393216
            _ExtentX        =   9340
            _ExtentY        =   17066
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
            MaxCols         =   7
            MaxRows         =   40
            ScrollBars      =   0
            SpreadDesigner  =   "MTX011.frx":0000
         End
         Begin VB.Label Label3 
            BackStyle       =   0  '����
            Caption         =   ">> ��ü ������ �ð�ǥ"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   120
            TabIndex        =   53
            Top             =   120
            Width           =   3015
         End
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C6AD84&
      BorderStyle     =   0  '����
      Caption         =   "Frame5"
      Height          =   10125
      Left            =   5430
      TabIndex        =   34
      Top             =   0
      Width           =   3675
      Begin VB.Frame Frame6 
         BackColor       =   &H00F7EFE7&
         BorderStyle     =   0  '����
         Caption         =   "Frame6"
         Height          =   10065
         Left            =   30
         TabIndex        =   35
         Top             =   30
         Width           =   3615
         Begin VB.CommandButton cmdKeyiN_Time 
            Caption         =   "Ű���� �Է�"
            Height          =   495
            Left            =   1710
            TabIndex        =   4
            Top             =   720
            Width           =   1665
         End
         Begin VB.CommandButton cmdKeyNew_Time 
            Caption         =   "�ű�"
            Height          =   495
            Left            =   180
            TabIndex        =   3
            Top             =   720
            Width           =   1365
         End
         Begin FPSpread.vaSpread sprKeyiN 
            Height          =   7455
            Left            =   30
            TabIndex        =   5
            Top             =   1440
            Width           =   3525
            _Version        =   393216
            _ExtentX        =   6218
            _ExtentY        =   13150
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
            MaxCols         =   4
            ProcessTab      =   -1  'True
            SpreadDesigner  =   "MTX011.frx":0A03
         End
         Begin VB.Label Label7 
            BackStyle       =   0  '����
            Caption         =   $"MTX011.frx":22C0
            Height          =   750
            Left            =   90
            TabIndex        =   58
            Top             =   9060
            Width           =   3345
         End
         Begin VB.Label Label5 
            BackStyle       =   0  '����
            Caption         =   ">> Ű���� ���"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   375
            Left            =   90
            TabIndex        =   36
            Top             =   420
            Width           =   3345
         End
      End
   End
   Begin VB.Frame fraPB 
      BackColor       =   &H0082C8E8&
      BorderStyle     =   0  '����
      Caption         =   "Frame4"
      Height          =   5205
      Left            =   8130
      TabIndex        =   24
      Top             =   11490
      Width           =   4875
      Begin VB.Frame Frame2 
         BackColor       =   &H00D2EAF5&
         BorderStyle     =   0  '����
         Caption         =   "Frame2"
         Height          =   5145
         Left            =   30
         TabIndex        =   25
         Top             =   30
         Width           =   4815
         Begin VB.TextBox txtTrxNM 
            Height          =   270
            IMEMode         =   10  '�ѱ� 
            Left            =   1320
            MaxLength       =   10
            TabIndex        =   20
            Text            =   "txtTrxNM"
            Top             =   1095
            Width           =   1455
         End
         Begin VB.ComboBox cboKaeyol_PB 
            Height          =   300
            Left            =   1320
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   21
            Top             =   1470
            Width           =   1035
         End
         Begin VB.CommandButton cmdNewPB 
            Caption         =   "�ű�"
            Height          =   435
            Left            =   210
            TabIndex        =   17
            Top             =   480
            Width           =   1125
         End
         Begin VB.CommandButton cmdSavePB 
            Caption         =   "���"
            Height          =   405
            Left            =   1710
            TabIndex        =   18
            Top             =   480
            Width           =   1155
         End
         Begin VB.CommandButton cmdDelPB 
            Caption         =   "����"
            Height          =   405
            Left            =   3180
            TabIndex        =   19
            Top             =   480
            Width           =   1125
         End
         Begin FPSpread.vaSpread sprPB 
            Height          =   3015
            Left            =   180
            TabIndex        =   23
            Top             =   1950
            Width           =   4485
            _Version        =   393216
            _ExtentX        =   7911
            _ExtentY        =   5318
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
            MaxCols         =   5
            SpreadDesigner  =   "MTX011.frx":2328
         End
         Begin MSComDlg.CommonDialog dlgCommon 
            Left            =   -30
            Top             =   690
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Label Label2 
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
            Left            =   3960
            TabIndex        =   28
            Top             =   90
            Width           =   1035
         End
         Begin VB.Label Label41 
            Alignment       =   1  '������ ����
            BackStyle       =   0  '����
            Caption         =   "���� �̸�"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   210
            TabIndex        =   27
            Top             =   1125
            Width           =   975
         End
         Begin VB.Label Label1 
            Alignment       =   1  '������ ����
            BackStyle       =   0  '����
            Caption         =   "�迭"
            Height          =   210
            Left            =   180
            TabIndex        =   26
            Top             =   1515
            Width           =   975
         End
         Begin VB.Label lblTrxColor 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  '���� ����
            Caption         =   $"MTX011.frx":3C13
            Height          =   705
            Left            =   3000
            TabIndex        =   22
            Top             =   1080
            Width           =   765
         End
      End
   End
   Begin VB.Frame fraData 
      BackColor       =   &H00C6AD84&
      BorderStyle     =   0  '����
      Caption         =   "Frame4"
      Height          =   10125
      Left            =   9120
      TabIndex        =   29
      Top             =   0
      Width           =   6405
      Begin VB.Frame Frame3 
         BackColor       =   &H00F7EFE7&
         BorderStyle     =   0  '����
         Caption         =   "Frame3"
         Height          =   10065
         Left            =   30
         TabIndex        =   30
         Top             =   30
         Width           =   6345
         Begin VB.TextBox txtTrx_CL_S 
            Enabled         =   0   'False
            Height          =   300
            Left            =   5130
            TabIndex        =   16
            Text            =   "txtTrx_CL_S"
            Top             =   5460
            Width           =   1125
         End
         Begin VB.TextBox txtTrxNM_S 
            Enabled         =   0   'False
            Height          =   300
            Left            =   5130
            TabIndex        =   15
            Text            =   "txtTrxNM_S"
            Top             =   5160
            Width           =   1125
         End
         Begin VB.TextBox txtKaeyol_S 
            Enabled         =   0   'False
            Height          =   300
            Left            =   5130
            TabIndex        =   14
            Text            =   "txtKaeyol_S"
            Top             =   4860
            Width           =   1125
         End
         Begin VB.TextBox txtTrxCD_S 
            Enabled         =   0   'False
            Height          =   300
            Left            =   5130
            TabIndex        =   13
            Text            =   "txtTrxCD_S"
            Top             =   4560
            Width           =   1125
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H0082C8E8&
            BorderStyle     =   0  '����
            Caption         =   "Frame4"
            Height          =   465
            Left            =   30
            TabIndex        =   32
            Top             =   5280
            Width           =   4845
            Begin VB.Frame Frame1 
               BackColor       =   &H00F7EFE7&
               BorderStyle     =   0  '����
               Height          =   405
               Left            =   30
               TabIndex        =   33
               Top             =   30
               Width           =   4785
               Begin VB.OptionButton optDelChk 
                  BackColor       =   &H00F7EFE7&
                  Caption         =   "����"
                  Height          =   405
                  Left            =   2490
                  TabIndex        =   11
                  Top             =   0
                  Width           =   1365
               End
               Begin VB.OptionButton optSaveChk 
                  BackColor       =   &H0082C8E8&
                  Caption         =   "���"
                  Height          =   405
                  Left            =   690
                  TabIndex        =   10
                  Top             =   0
                  Width           =   1365
               End
            End
         End
         Begin VB.CommandButton cmdTrx01 
            Caption         =   "������ �׸���ȸ"
            Height          =   405
            Left            =   1500
            TabIndex        =   7
            Top             =   75
            Width           =   1635
         End
         Begin VB.ComboBox cboKaeyol 
            Height          =   300
            Left            =   450
            Style           =   2  '��Ӵٿ� ���
            TabIndex        =   6
            Top             =   120
            Width           =   885
         End
         Begin VB.CommandButton cmdPB_iNsert 
            Caption         =   "������ �׸��߰�"
            Height          =   405
            Left            =   3270
            TabIndex        =   8
            Top             =   60
            Width           =   1665
         End
         Begin FPSpread.vaSpread sprTrxType 
            Height          =   4125
            Left            =   30
            TabIndex        =   12
            Top             =   5790
            Width           =   6285
            _Version        =   393216
            _ExtentX        =   11086
            _ExtentY        =   7276
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
            MaxCols         =   7
            MaxRows         =   10
            ScrollBars      =   0
            SpreadDesigner  =   "MTX011.frx":3C29
         End
         Begin FPSpread.vaSpread sprTRX01 
            Height          =   4725
            Left            =   60
            TabIndex        =   9
            Top             =   510
            Width           =   4875
            _Version        =   393216
            _ExtentX        =   8599
            _ExtentY        =   8334
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
            MaxCols         =   4
            ScrollBars      =   2
            SpreadDesigner  =   "MTX011.frx":4154
         End
         Begin VB.Label Label6 
            BackStyle       =   0  '����
            Caption         =   "�迭"
            Height          =   210
            Left            =   60
            TabIndex        =   31
            Top             =   165
            Width           =   975
         End
      End
   End
End
Attribute VB_Name = "MTX011"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################################################
'   �� �� ��  �� : �뼺�п� ���л���, �ݹ��� & �ð�ǥ ���α׷�
'   ����ý��۸� :
'   ��   ��   �� : MTX011
'   �� ��  �� �� : �ð�ǥ ����� >> ������ �ð�ǥ ��� CP
'
'   ��   ��   �� : 2007/12/26
'   ��   ��   �� : ���ϱ�
' --------------------------------------------------------------------------------------------------------------
' --------------------------------------------------------------------------------------------------------------
'                 ��     ��     ��     ��
' --------------------------------------------------------------------------------------------------------------
'   1. ������ :
'   2. ��  �� :
'################################################################################################################


Option Explicit




Private Sub cboKaeyol_All_Click()
    If Me.Tag = "LOAD" Then Exit Sub
    
    cboKaeyol.ListIndex = cboKaeyol_All.ListIndex
    cboKaeyol_PB.ListIndex = cboKaeyol.ListIndex
    
    cmdNewPB.Tag = "SELECT"
        Call cmdFindAll_Click
        Call cmdTrx01_Click
        Call cmdNewPB_Click
    cmdNewPB.Tag = ""
End Sub




Private Sub cmdPB_iNsert_Click()
    fraPB.Visible = True
    txtTrxNM.SetFocus
    
End Sub

Private Sub Form_Load()
        
    Me.Move 0, 0, 15700, 10600
    
    Me.Tag = "LOAD"
        With sprTRX01
            .ShadowColor = basModule.ShadowColor1
            .ShadowDark = basModule.ShadowDark1
            .ShadowText = basModule.ShadowText1
            .GridColor = basModule.GridColor1
            .GrayAreaBackColor = basModule.GrayAreaBackColor1
            
            .Row = 1:   .Row2 = .MaxRows
            .Col = 1:   .Col2 = .MaxCols
            .BlockMode = True
                .Lock = True
                .Protect = True
            .BlockMode = False
            
            .MaxRows = 0
        End With
        
        With sprPB
            .ShadowColor = basModule.ShadowColor2
            .ShadowDark = basModule.ShadowDark2
            .ShadowText = basModule.ShadowText2
            .GridColor = basModule.GridColor2
            .GrayAreaBackColor = basModule.GrayAreaBackColor2
            
            .Row = 1:   .Row2 = .MaxRows
            .Col = 1:   .Col2 = .MaxCols
            .BlockMode = True
                .Lock = True
                .Protect = True
            .BlockMode = False
            
            .MaxRows = 0
        End With
        
        With sprTrxType
            .ShadowColor = basModule.ShadowColor2
            .ShadowDark = basModule.ShadowDark2
            .ShadowText = basModule.ShadowText2
            .GridColor = basModule.GridColor2
            .GrayAreaBackColor = basModule.GrayAreaBackColor2
            
            .Row = 1:   .Row2 = .MaxRows
            .Col = 1:   .Col2 = .MaxCols
            .BlockMode = True
                .Lock = True
                .Protect = True
            .BlockMode = False
        End With
        
        With sprTrx_T
            .ShadowColor = basModule.ShadowColor1
            .ShadowDark = basModule.ShadowDark1
            .ShadowText = basModule.ShadowText1
            .GridColor = &HFFFFFF
            .GrayAreaBackColor = basModule.GrayAreaBackColor1
            
            .Row = 1:   .Row2 = .MaxRows
            .Col = 1:   .Col2 = .MaxCols
            .BlockMode = True
                .Lock = True
                .Protect = True
            .BlockMode = False
        End With
        
        With sprKeyiN
            .ShadowColor = basModule.ShadowColor2
            .ShadowDark = basModule.ShadowDark2
            .ShadowText = basModule.ShadowText2
            .GridColor = basModule.GridColor2
            .GrayAreaBackColor = basModule.GrayAreaBackColor2
            
            .Row = 1:   .Row2 = .MaxRows
            .Col = 1:   .Col2 = .MaxCols
            .BlockMode = True
                .Lock = True
                .Protect = True
            .BlockMode = False
            .MaxRows = 0
        End With
        
        With cboKaeyol
            .Clear
            .AddItem "�ι�" & Space(30) & "01"
            .AddItem "�ڿ�" & Space(30) & "02"
            
            .ListIndex = 0
        End With
        
        With cboKaeyol_PB
            .Clear
            .AddItem "�ι�" & Space(30) & "01"
            .AddItem "�ڿ�" & Space(30) & "02"
            
            .ListIndex = 0
        End With
        
        With cboKaeyol_All
            .Clear
            .AddItem "�ι�" & Space(30) & "01"
            .AddItem "�ڿ�" & Space(30) & "02"
            
            .ListIndex = 0
        End With

        fraPB.Top = fraData.Top + 590
        fraPB.Left = fraData.Left + 210
        fraPB.ZOrder 0
        fraPB.Visible = False
        
        Call initData
            
    Me.Tag = ""
    
End Sub


Private Sub cboKaeyol_Click()
    If Me.Tag = "LOAD" Then Exit Sub
    
    cboKaeyol_PB.ListIndex = cboKaeyol.ListIndex
    
    cmdNewPB.Tag = "SELECT"
        Call cmdFindAll_Click
        Call cmdTrx01_Click
        Call cmdNewPB_Click
    cmdNewPB.Tag = ""
End Sub

Private Sub cboKaeyol_PB_Click()
    If Me.Tag = "LOAD" Then Exit Sub
    
    cboKaeyol.ListIndex = cboKaeyol_PB.ListIndex
    
    cmdNewPB.Tag = "SELECT"
        Call cmdFindAll_Click
        Call cmdTrx01_Click
        Call cmdNewPB_Click
    cmdNewPB.Tag = ""
End Sub

Private Sub initData()
    
    cmdNewPB.Tag = "SELECT"
        Call cmdFindAll_Click
        Call cmdTrx01_Click
        Call cmdNewPB_Click
    cmdNewPB.Tag = ""
    
    txtTrxNM.Text = ""
    lblTrxColor.BackColor = basModule.WhiteColor
    
    optSaveChk.Value = True
    optDelChk.Value = False
    
    txtTrxCD_S.Text = ""
    txtKaeyol_S.Text = ""
    txtTrxNM_S.Text = ""
    txtTrx_CL_S.Text = ""
    txtTrx_CL_S.BackColor = &HFFFFFF
    
End Sub


'������ �׸���ȸ
Private Sub cmdTrx01_Click()
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    
    Dim sStr        As String
    Dim nRec        As Long
    Dim ni          As Integer
    
    Dim nColor      As Long
    Dim sTmp        As String
    
    Dim sComp       As String
    
    
    On Error GoTo ErrStmt
    
    sStr = ""
    sStr = sStr & "  SELECT ACID, TRXCD, TRXNM, TRX_CL"
    sStr = sStr & "    FROM (SELECT ACID, TRXCD, TRXNM, TRX_CL"
    sStr = sStr & "            From SDTRX01TB"
    sStr = sStr & "           WHERE ACID   = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "             AND KAEYOL = '" & Trim(Right(cboKaeyol.Text, 30)) & "'"
    sStr = sStr & "             AND TRXCD LIKE 'PB%'"
    sStr = sStr & "          Union All"
    sStr = sStr & "          SELECT ACID, TRXCD, TRXNM, TRX_CL"
    sStr = sStr & "            From SDTRX01TB"
    sStr = sStr & "           WHERE ACID   = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "             AND KAEYOL = '" & Trim(Right(cboKaeyol.Text, 30)) & "'"
    sStr = sStr & "             AND TRXCD LIKE 'A%'"
    sStr = sStr & "          Union All"
    sStr = sStr & "          SELECT ACID, TRXCD, TRXNM, TRX_CL"
    sStr = sStr & "            From SDTRX01TB"
    sStr = sStr & "           WHERE ACID   = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "             AND KAEYOL = '" & Trim(Right(cboKaeyol.Text, 30)) & "'"
    sStr = sStr & "             AND TRXCD LIKE 'B%'"
    sStr = sStr & "          Union All"
    sStr = sStr & "          SELECT ACID, TRXCD, TRXNM, TRX_CL"
    sStr = sStr & "            From SDTRX01TB"
    sStr = sStr & "           WHERE ACID   = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "             AND KAEYOL = '" & Trim(Right(cboKaeyol.Text, 30)) & "'"
    sStr = sStr & "             AND TRXCD LIKE 'C%'"
    sStr = sStr & "          )"
    
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

    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic���·� ���ԵǸ� record count�� �� �� ����.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop
    
    With DBRec
        
        sprTRX01.MaxRows = 0
        
        If .RecordCount > 0 Then
            .MoveFirst
            
            For nRec = 1 To .RecordCount Step 1
                sprTRX01.MaxRows = sprTRX01.MaxRows + 1
                sprTRX01.Row = sprTRX01.MaxRows:        sprTRX01.RowHeight(sprTRX01.Row) = 14
                
                sprTRX01.Col = 1:       sTmp = "":      If IsNull(.Fields("TRXCD")) = False Then sTmp = Trim(.Fields("TRXCD"))
                    Call basFunction.Set_SprType_Text(sprTRX01, "center", "left", basFunction.LenKor(sTmp), sTmp)
                    
                If StrComp(Left(sTmp, 1), sComp, vbTextCompare) <> 0 Then
                    Call sprTRX01.SetCellBorder(1, sprTRX01.Row, sprTRX01.MaxCols, sprTRX01.Row, 4, basModule.SectionColor1, CellBorderStyleSolid)
                    sComp = Left(sTmp, 1)
                End If
                
                sprTRX01.Col = 2:       sTmp = "":      If IsNull(.Fields("TRXNM")) = False Then sTmp = Trim(.Fields("TRXNM"))
                    Call basFunction.Set_SprType_Text(sprTRX01, "center", "left", basFunction.LenKor(sTmp), sTmp)
                sprTRX01.Col = 3
                    nColor = 0
                    If IsNumeric(.Fields("TRX_CL")) = True Then nColor = CLng(.Fields("TRX_CL"))
                    sprTRX01.Row2 = sprTRX01.Row
                    sprTRX01.Col2 = sprTRX01.Col
                    sprTRX01.BlockMode = True
                        sprTRX01.BackColor = nColor
                        sprTRX01.BackColorStyle = BackColorStyleUnderGrid
                    sprTRX01.BlockMode = False
                sprTRX01.Col = 4
                    Call basFunction.Set_SprType_ChkBox(sprTRX01)
                    
                .MoveNext
            Next nRec
        End If
    End With
    
    
    With sprTRX01
        .Row = 1:       .Row2 = .MaxRows
        .Col = 1:       .Col2 = .MaxCols
        .BlockMode = True
            .Lock = True
            .Protect = True
        .BlockMode = False
    End With

    If cmdNewPB.Tag = "" Then
        MsgBox "������ �׸� ��ȸ�Ͽ����ϴ�.", vbInformation + vbOKOnly, "������ �׸���ȸ"
    End If

    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    Exit Sub
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    On Error GoTo 0
    MsgBox "������ �׸� ��ȸ�� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "������ �׸���ȸ"
End Sub






Private Sub Label2_Click()
    fraPB.Visible = False
End Sub



Private Sub sprTRX01_Click(ByVal Col As Long, ByVal Row As Long)
    If Row < 1 Then Exit Sub
    
    With sprTRX01
        .Enabled = False
        
        If Trim(.Tag) = "" Then .Tag = "1"
        
        .Row = CLng(.Tag):  .Row2 = .Row
        .Col = 1:           .Col2 = 2
        .BlockMode = True
            .BackColor = basModule.WhiteColor
            .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
        
        .Row = CLng(.Tag):  .Row2 = .Row
        .Col = 4:           .Col2 = 4
        .BlockMode = True
            .BackColor = basModule.WhiteColor
            .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
        
        .Col = .MaxCols:    .Value = 0
        
        .Row = Row:         .Row2 = .Row
        .Col = 1:           .Col2 = 2
        .BlockMode = True
            .BackColor = basModule.SelectColor1
            .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
        
        .Row = Row:         .Row2 = .Row
        .Col = 4:           .Col2 = 4
        .BlockMode = True
            .BackColor = basModule.SelectColor1
            .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
        
        .Tag = Trim(CStr(Row))
        
        .Col = .MaxCols:    .Value = 1
        
        .Col = 1:
            Call Find_TrxPart_Detail(Left(Trim(.Text), 1))      '< ������ȸ
            txtTrxCD_S.Text = Trim(.Text)                           '< ������ �ڵ�
        .Col = 2
            txtTrxNM_S.Text = Trim(.Text)
        .Col = 3
            txtTrx_CL_S.BackColor = .BackColor
        
        txtKaeyol_S.Text = Trim(Right(cboKaeyol.Text, 30))      '< �迭�ڵ�
        
        .Enabled = True
    End With
    
    
End Sub


Private Sub lblTrxColor_Click()

    On Error GoTo ErrStmt
    
    With dlgCommon
        .CancelError = True
        .ShowColor
        
        lblTrxColor.BackColor = .color
         
    End With
    
    Exit Sub
ErrStmt:
    
End Sub


Private Sub cmdNewPB_Click()
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    
    Dim sStr        As String
    Dim nRec        As Long
    Dim ni          As Integer
    
    Dim nColor      As Long
    Dim sTmp        As String
    
    txtTrxNM.Text = ""
    lblTrxColor.BackColor = basModule.WhiteColor
    
    
' ������ ��ȸ
    
    On Error GoTo ErrStmt
    
    sStr = ""
    sStr = sStr & "  SELECT TRXCD, KAEYOL, TRXNM, TRX_CL"
    sStr = sStr & "    From SDTRX01TB"
    sStr = sStr & "   WHERE ACID   = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "     AND KAEYOL = '" & Trim(Right(cboKaeyol_PB.Text, 30)) & "'"
    sStr = sStr & "     AND TRXCD LIKE 'PB%'"
    
    
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

    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic���·� ���ԵǸ� record count�� �� �� ����.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop
    
    With DBRec
        
        sprPB.MaxRows = 0
        txtTrxNM.Tag = ""       '< �űԵ��
        
        
        If .RecordCount > 0 Then
            .MoveFirst
            
            For nRec = 1 To .RecordCount Step 1
                sprPB.MaxRows = sprPB.MaxRows + 1
                sprPB.Row = sprPB.MaxRows:        sprPB.RowHeight(sprPB.Row) = 14
                
                sprPB.Col = 1:                  sTmp = "":      If IsNull(.Fields("TRXCD")) = False Then sTmp = Trim(.Fields("TRXCD"))
                    Call basFunction.Set_SprType_Text(sprPB, "center", "left", basFunction.LenKor(sTmp), sTmp)
                sprPB.Col = sprPB.Col + 1:      sTmp = "":      If IsNull(.Fields("KAEYOL")) = False Then sTmp = Trim(.Fields("KAEYOL"))
                    Call basFunction.Set_SprType_Text(sprPB, "center", "left", basFunction.LenKor(sTmp), sTmp)
                sprPB.Col = sprPB.Col + 1:      sTmp = "":      If IsNull(.Fields("TRXNM")) = False Then sTmp = Trim(.Fields("TRXNM"))
                    Call basFunction.Set_SprType_Text(sprPB, "center", "left", basFunction.LenKor(sTmp), sTmp)
                sprPB.Col = sprPB.Col + 1
                    nColor = 0
                    If IsNumeric(.Fields("TRX_CL")) = True Then nColor = CLng(.Fields("TRX_CL"))
                    sprPB.Row2 = sprPB.Row
                    sprPB.Col2 = sprPB.Col
                    sprPB.BlockMode = True
                        sprPB.BackColor = nColor
                        sprPB.BackColorStyle = BackColorStyleUnderGrid
                    sprPB.BlockMode = False
                sprPB.Col = sprPB.Col + 1
                    Call basFunction.Set_SprType_ChkBox(sprPB)
                    
                .MoveNext
            Next nRec
        End If
    End With
    
    
    With sprPB
        .Row = 1:       .Row2 = .MaxRows
        .Col = 1:       .Col2 = .MaxCols
        .BlockMode = True
            .Lock = True
            .Protect = True
        .BlockMode = False
    End With

    If cmdNewPB.Tag = "" Then
        MsgBox "���볻�� ��ȸ�Ͽ����ϴ�.", vbInformation + vbOKOnly, "������ �׸���ȸ"
    End If

'    If MTX011.Tag <> "LOAD" Then
'        txtTrxNM.SetFocus
'    End If
    
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    Exit Sub
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    On Error GoTo 0
    MsgBox "���볻�� ��ȸ�� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "������ �׸���ȸ"
    
End Sub

'������ �ð�ǥ ���볻�� ���
Private Sub cmdSavePB_Click()
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter

    Dim sStr        As String
    Dim sTmp        As String
    Dim sTrxCD      As String
    
    Dim nExe        As Long
    Dim ni          As Long
    
    
    On Error GoTo ErrStmt
    
    
    If Trim(txtTrxNM.Text) = "" Then
        MsgBox "���뱸�� ������ �����ϴ�.", vbExclamation + vbOKOnly, "������ �׸񳻿�"
        Exit Sub
    End If
    
    Set DBCmd = New ADODB.Command
    Set DBRec = New ADODB.Recordset
    Set DBParam = New ADODB.Parameter
    
    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
    
    basDataBase.DBConn.BeginTrans
    

'## ����
    If Trim(txtTrxNM.Tag) <> "" Then
        sTrxCD = Trim(txtTrxNM.Tag)
        
    Else
'## �ű�
        sStr = ""
        sStr = sStr & "  SELECT MAX(TRXCD) AS TRXCD"
        sStr = sStr & "    FROM (SELECT 'PB'||TRIM(TO_CHAR(TO_NUMBER(MAX(SUBSTR(TRXCD,3,2))) + 1,'00')) AS TRXCD"
        sStr = sStr & "            From SDTRX01TB"
        sStr = sStr & "           WHERE ACID   = '" & Trim(basModule.SchCD) & "'"
        sStr = sStr & "             AND TRXCD  LIKE 'PB%'"
        sStr = sStr & "             AND KAEYOL = '" & Trim(Right(cboKaeyol_PB.Text, 30)) & "'"
        sStr = sStr & "          Union All"
        sStr = sStr & "          SELECT 'PB01' AS TRXCD"
        sStr = sStr & "            From DUAL"
        sStr = sStr & "         )"
        
        DBCmd.CommandText = sStr
        DBCmd.CommandType = adCmdText
        DBCmd.CommandTimeout = 30
        


                    
    '    ' ACID
    '        sTmp = Trim(basModule.SchCD)
    '        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
    '            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
            
        DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic���·� ���ԵǸ� record count�� �� �� ����.
        Do While DBRec.State And adStateExecuting
            DoEvents
        Loop
                
        With DBRec
            sTrxCD = "PB01"
            
            If .RecordCount > 0 Then
                .MoveFirst
                If IsNull(.Fields("TRXCD")) = False Then
                    sTrxCD = Trim(.Fields("TRXCD"))
                Else
                    sTrxCD = "PB01"
                End If
            End If
        End With
    End If
    
    
    On Error GoTo 0
    On Error Resume Next
    
'## ���
    sStr = ""
    sStr = sStr & "  INSERT INTO SDTRX01TB (ACID, TRXCD, KAEYOL, TRXNM, TRX_CL) "
    sStr = sStr & "         VALUES ("
    sStr = sStr & "                 '" & Trim(basModule.SchCD) & "',"
    sStr = sStr & "                 '" & sTrxCD & "',"
    sStr = sStr & "                 '" & Trim(Right(cboKaeyol_PB.Text, 30)) & "',"
    sStr = sStr & "                 '" & Trim(txtTrxNM.Text) & "',"
    sStr = sStr & "                 " & lblTrxColor.BackColor & ""
    sStr = sStr & "         )"
    



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
        basDataBase.DBConn.CommitTrans
        cmdNewPB.Tag = "SAVE"
            Call cmdFindAll_Click
            Call cmdNewPB_Click
            Call cmdTrx01_Click
        cmdNewPB.Tag = ""
        MsgBox "��� �Ͽ����ϴ�.", vbInformation + vbOKOnly, "������ �׸񳻿�"
        
    Else
' UPDATE
        
        On Error GoTo 0
        On Error GoTo ErrStmt
        
        sStr = ""
        sStr = sStr & "  UPDATE SDTRX01TB "
        sStr = sStr & "     SET TRXNM  = '" & Trim(txtTrxNM.Text) & "',"
        sStr = sStr & "         TRX_CL = " & lblTrxColor.BackColor & ""
        sStr = sStr & "   WHERE ACID   = '" & Trim(basModule.SchCD) & "'"
        sStr = sStr & "     AND TRXCD  = '" & sTrxCD & "'"
        sStr = sStr & "     AND KAEYOL = '" & Trim(Right(cboKaeyol_PB.Text, 30)) & "'"
        
        


    
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
            basDataBase.DBConn.CommitTrans
            
            cmdNewPB.Tag = "SAVE"
                Call cmdFindAll_Click
                Call cmdNewPB_Click
                Call cmdTrx01_Click
            cmdNewPB.Tag = ""
            
            MsgBox "���� �Ͽ����ϴ�.", vbInformation + vbOKOnly, "������ �׸񳻿�"
        Else
            basDataBase.DBConn.RollbackTrans
            MsgBox "ó���� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "������ �׸񳻿�"
        End If
    End If
        
    Set DBCmd = Nothing
    Set DBRec = Nothing
    Set DBParam = Nothing
    
    Exit Sub

ErrStmt:
    basDataBase.DBConn.RollbackTrans
    MsgBox "���볻�� ��Ͻ� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "������ �׸񳻿�"
    
    Set DBCmd = Nothing
    Set DBRec = Nothing
    Set DBParam = Nothing

End Sub


Private Sub sprPB_Click(ByVal Col As Long, ByVal Row As Long)
    If Row < 1 Then Exit Sub
    
    With sprPB
        If Trim(.Tag) = "" Then .Tag = "1"
        
        .Row = CLng(.Tag):  .Row2 = .Row
        .Col = 1:           .Col2 = 3
        .BlockMode = True
            .BackColor = basModule.WhiteColor
            .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
        
        .Row = CLng(.Tag):  .Row2 = .Row
        .Col = 5:           .Col2 = 5
        .BlockMode = True
            .BackColor = basModule.WhiteColor
            .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
        
        .Col = .MaxCols:    .Value = 0
        
        .Row = Row:         .Row2 = .Row
        .Col = 1:           .Col2 = 3
        .BlockMode = True
            .BackColor = basModule.SelectColor2
            .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
        
        .Row = Row:         .Row2 = .Row
        .Col = 5:           .Col2 = 5
        .BlockMode = True
            .BackColor = basModule.SelectColor2
            .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
        
        .Col = .MaxCols:    .Value = 1
        
        .Col = 1:   txtTrxNM.Tag = Trim(.Text)
        .Col = 2
            Select Case Trim(.Text)
                Case "01"
                    cboKaeyol_PB.ListIndex = 0
                Case "02"
                    cboKaeyol_PB.ListIndex = 1
            End Select
        .Col = 3:   txtTrxNM.Text = Trim(.Text)
        .Col = 4:   lblTrxColor.BackColor = CLng(.BackColor)
        
        .Tag = Trim(CStr(Row))
        
    End With
    
End Sub


'������ �ð�ǥ ���볻�� ����
Private Sub cmdDelPB_Click()
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter

    Dim sStr        As String
    Dim sTmp        As String
    Dim sTrxCD      As String
    
    Dim nExe        As Long
    Dim ni          As Long
    
    
    On Error GoTo ErrStmt
    
    
    If Trim(txtTrxNM.Tag) = "" Then
        MsgBox "���뱸�� ������ �����ϴ�." & vbCrLf & _
               "��ȸ �� �����Ͻʽÿ�.", vbExclamation + vbOKOnly, "������ �׸񳻿�"
        Exit Sub
    End If
    
    basDataBase.DBConn.BeginTrans
    
    Set DBCmd = New ADODB.Command
    Set DBRec = New ADODB.Recordset
    Set DBParam = New ADODB.Parameter
    
    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
    
    sStr = ""
    sStr = sStr & "  DELETE "
    sStr = sStr & "    FROM SDTRX01TB "
    sStr = sStr & "   WHERE ACID   = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "     AND TRXCD  = '" & Trim(txtTrxNM.Tag) & "'"
    sStr = sStr & "     AND KAEYOL = '" & Trim(Right(cboKaeyol_PB.Text, 30)) & "'"
        



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
        basDataBase.DBConn.CommitTrans
        
        cmdNewPB.Tag = "DEL"
            Call cmdFindAll_Click
            Call cmdNewPB_Click
            Call cmdTrx01_Click
        cmdNewPB.Tag = ""
        
        MsgBox "���� �Ͽ����ϴ�.", vbInformation + vbOKOnly, "������ �׸񳻿�"
    Else
    
        basDataBase.DBConn.RollbackTrans
        MsgBox "ó���� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "������ �׸񳻿�"
    End If
        
    Set DBCmd = Nothing
    Set DBParam = Nothing
    
    Exit Sub

ErrStmt:
    basDataBase.DBConn.RollbackTrans
    MsgBox "���볻�� ������ ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "������ �׸񳻿�"
    
    Set DBCmd = Nothing
    Set DBParam = Nothing

End Sub











































Private Sub Find_TrxPart_Detail(ByVal aTrxTypes As String)
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter

    Dim nLength     As Long
    Dim sStr        As String
    Dim ni          As Integer
    Dim nRec        As Long

    Dim sTmp        As String
    Dim nTmp        As Double

    Dim sTrxTypes   As String
    Dim sTrxCD      As String
    Dim sTrx        As String
    Dim nLesson     As Integer
    Dim nWeeks      As Integer
    Dim nColor      As Long

    Dim nRow        As Long
    Dim nCol        As Long
    Dim sKaeyol     As String

    On Error GoTo ErrStmt

'    With sprTRX01
'        For nRow = 1 To .MaxRows Step 1
'            .Row = nRow
'            .Col = .MaxCols
'            If .Value = 1 Then
'                .Col = 1
'                sTrxTypes = Left(Trim(.Text), 1)            '< �����ڵ�
'            End If
'        Next nRow
'    End With


    sTrxTypes = aTrxTypes       '< �����ڵ�

    sKaeyol = Trim(Right(cboKaeyol.Text, 30))

    '<< �ʱ�ȭ
    With sprTrxType
        For nRow = 1 To .MaxRows Step 1
            For nCol = 1 To .MaxCols Step 1
                .Row = nRow
                .Col = nCol
                    .Text = ""
            Next nCol
        Next nRow

        .Row = 1:       .Row2 = .MaxRows
        .Col = 1:       .Col2 = .MaxCols
        .BlockMode = True
            .BackColor = basModule.WhiteColor
            .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
    End With


    sStr = ""
    sStr = sStr & "  SELECT A.TRXCD, A.TRXNM, B.LESSON, B.WEEKS, A.TRX_CL "
    sStr = sStr & "    FROM SDTRX01TB A, SDTRX11TB B"
    sStr = sStr & "   Where A.ACID   = B.ACID"
    sStr = sStr & "     AND A.TRXCD  = B.TRXCD"
    sStr = sStr & "     AND A.KAEYOL = B.KAEYOL"
    sStr = sStr & "     AND A.ACID   = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "     AND A.KAEYOL = '" & sKaeyol & "'"
    sStr = sStr & "     AND A.TRXCD  LIKE '" & sTrxTypes & "%'"                 '< �����ڵ� ��ȸ
    sStr = sStr & "  UNION ALL"
    sStr = sStr & "  SELECT A.TRXCD, A.TRXNM, B.LESSON, B.WEEKS, A.TRX_CL"
    sStr = sStr & "    FROM SDTRX01TB A, SDTRX11TB B"
    sStr = sStr & "   Where A.ACID   = B.ACID"
    sStr = sStr & "     AND A.TRXCD  = B.TRXCD"
    sStr = sStr & "     AND A.KAEYOL = B.KAEYOL"
    sStr = sStr & "     AND A.ACID   = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "     AND A.KAEYOL = '" & sKaeyol & "'"
    sStr = sStr & "     AND A.TRXCD LIKE 'PB%'"

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
'    ' LSNTYPE
'        sTmp = Left(Trim(Right(cboLsnType, 30)), 1) & "%"
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("LSNTYPE", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam

    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic���·� ���ԵǸ� record count�� �� �� ����.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop

    With DBRec

        If .RecordCount > 0 Then
            .MoveFirst

            For nRec = 1 To .RecordCount Step 1
                nColor = 0
                sTrx = ""
'                    If IsNull(.Fields("TRXCD")) = False Then
'                         sTrx = Trim(.Fields("TRXCD"))
'                         sTrxCD = sTrx
'                    End If

'                                If IsNull(.Fields("TRXNM")) = False Then sTrx = Trim(.Fields("TRXNM")) & Space(30) & sTrx
                                If IsNull(.Fields("TRXNM")) = False Then sTrx = Trim(.Fields("TRXNM"))

                nLesson = 0:    If IsNull(.Fields("LESSON")) = False Then nLesson = CInt(.Fields("LESSON"))
                nWeeks = 0:     If IsNull(.Fields("WEEKS")) = False Then nWeeks = CInt(.Fields("WEEKS"))
                nColor = 0:     If IsNull(.Fields("TRX_CL")) = False Then nColor = CLng(.Fields("TRX_CL"))

                Select Case nWeeks
                    Case 2
                        sprTrxType.Col = 1
                    Case 3
                        sprTrxType.Col = 2
                    Case 4
                        sprTrxType.Col = 3
                    Case 5
                        sprTrxType.Col = 4
                    Case 6
                        sprTrxType.Col = 5
                    Case 7
                        sprTrxType.Col = 6
                    Case 1
                        sprTrxType.Col = 7
                End Select
                sprTrxType.Row = nLesson
                sTmp = sprTrxType.Text
                    If InStr(1, sTmp, sTrx, vbTextCompare) = 0 Then
                        If basFunction.LenKor(sTmp) > 0 Then
                            sTrx = sTmp & vbCrLf & sTrx
                        End If
                        Call basFunction.Set_SprType_Text(sprTrxType, "TOP", "LEFT", basFunction.LenKor(sTrx), Trim(sTrx))
                        sprTrxType.TypeEditMultiLine = True
                    End If

                sprTrxType.Row2 = sprTrxType.Row
                sprTrxType.Col2 = sprTrxType.Col
                sprTrxType.BlockMode = True
                    sprTrxType.BackColor = nColor
                    sprTrxType.BackColorStyle = BackColorStyleUnderGrid
                sprTrxType.BlockMode = False


                .MoveNext
            Next nRec
        End If
    End With

    'MsgBox "������ �ð�ǥ ��ȸ�Ͽ����ϴ�.", vbInformation + vbOKOnly, "�ð�ǥ ��ȸ�ϱ�"

    Set DBCmd = Nothing
    Set DBRec = Nothing

    Exit Sub
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing

    On Error GoTo 0
    MsgBox "������ �ð� ��ȸ�� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "�ð�ǥ ��ȸ�ϱ�"
End Sub




'>> ��ü ������ �ð�ǥ ��ȸ
Private Sub cmdFindAll_Click()
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter

    Dim nLength     As Long
    Dim sStr        As String
    Dim ni          As Integer
    Dim nRec        As Long

    Dim sTmp        As String
    Dim nTmp        As Double

    Dim sTrxTypes   As String
    Dim sTrxCD      As String
    Dim sTrx        As String
    Dim nLesson     As Integer
    Dim nWeeks      As Integer
    Dim nColor      As Long

    Dim nRow        As Long
    Dim nCol        As Long
    Dim sKaeyol     As String

    On Error GoTo ErrStmt
    
    sKaeyol = Trim(Right(cboKaeyol.Text, 30))

    '<< �ʱ�ȭ
    With sprTrx_T
        For nRow = 1 To .MaxRows Step 1
            For nCol = 1 To .MaxCols Step 1
                .Row = nRow
                .Col = nCol
                    .Text = ""
            Next nCol
        Next nRow

        .Row = 1:       .Row2 = .MaxRows
        .Col = 1:       .Col2 = .MaxCols
        .BlockMode = True
            .BackColor = basModule.WhiteColor
            .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
    End With


    sStr = ""
    sStr = sStr & "  SELECT A.TRXCD, A.TRXNM, B.LESSON, B.WEEKS, A.TRX_CL "
    sStr = sStr & "    FROM SDTRX01TB A, SDTRX11TB B"
    sStr = sStr & "   Where A.ACID   = B.ACID"
    sStr = sStr & "     AND A.TRXCD  = B.TRXCD"
    sStr = sStr & "     AND A.KAEYOL = B.KAEYOL"
    sStr = sStr & "     AND A.ACID   = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "     AND A.KAEYOL = '" & sKaeyol & "'"
    sStr = sStr & "   ORDER BY A.TRXCD "
    
    
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
'    ' LSNTYPE
'        sTmp = Left(Trim(Right(cboLsnType, 30)), 1) & "%"
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("LSNTYPE", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam

    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic���·� ���ԵǸ� record count�� �� �� ����.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop

    With DBRec

        If .RecordCount > 0 Then
            .MoveFirst

            For nRec = 1 To .RecordCount Step 1
                nColor = 0
                
                sTrxCD = "":    If IsNull(.Fields("TRXCD")) = False Then sTrxCD = Trim(.Fields("TRXCD"))
                sTrx = "":      If IsNull(.Fields("TRXNM")) = False Then sTrx = Trim(.Fields("TRXNM"))
                nLesson = 0:    If IsNull(.Fields("LESSON")) = False Then nLesson = CInt(.Fields("LESSON"))
                nWeeks = 0:     If IsNull(.Fields("WEEKS")) = False Then nWeeks = CInt(.Fields("WEEKS"))
                nColor = 0:     If IsNull(.Fields("TRX_CL")) = False Then nColor = CLng(.Fields("TRX_CL"))

                Select Case nWeeks
                    Case 2
                        sprTrx_T.Col = 1
                    Case 3
                        sprTrx_T.Col = 2
                    Case 4
                        sprTrx_T.Col = 3
                    Case 5
                        sprTrx_T.Col = 4
                    Case 6
                        sprTrx_T.Col = 5
                    Case 7
                        sprTrx_T.Col = 6
                    Case 1
                        sprTrx_T.Col = 7
                End Select
                
                Select Case Left(Trim(sTrxCD), 1)
                    Case "A"
                        sprTrx_T.Row = nLesson + (3 * (nLesson - 1))
                    Case "B"
                        sprTrx_T.Row = nLesson + (3 * (nLesson - 1)) + 1
                    Case "C"
                        sprTrx_T.Row = nLesson + (3 * (nLesson - 1)) + 2
                    Case Else
                        sprTrx_T.Row = nLesson + (3 * (nLesson - 1)) + 3
                End Select
                
                sTmp = sprTrx_T.Text
                    If InStr(1, sTmp, sTrx, vbTextCompare) = 0 Then
                        If basFunction.LenKor(sTmp) > 0 Then
                            sTrx = sTmp & vbCrLf & sTrx
                        End If
                        Call basFunction.Set_SprType_Text(sprTrx_T, "TOP", "LEFT", basFunction.LenKor(sTrx), Trim(sTrx))
                        sprTrx_T.TypeEditMultiLine = True
                    End If

                sprTrx_T.Row2 = sprTrx_T.Row
                sprTrx_T.Col2 = sprTrx_T.Col
                sprTrx_T.BlockMode = True
                    sprTrx_T.BackColor = nColor
                    sprTrx_T.BackColorStyle = BackColorStyleUnderGrid
                sprTrx_T.BlockMode = False


                .MoveNext
            Next nRec
        End If
    End With

    'MsgBox "������ �ð�ǥ ��ȸ�Ͽ����ϴ�.", vbInformation + vbOKOnly, "�ð�ǥ ��ȸ�ϱ�"

    Set DBCmd = Nothing
    Set DBRec = Nothing

    Exit Sub
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing

    On Error GoTo 0
    MsgBox "������ �ð� ��ȸ�� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "�ð�ǥ ��ȸ�ϱ�"
End Sub





'## ������ �ð�ǥ ��� �� ����
Private Sub sprTrxType_DblClick(ByVal Col As Long, ByVal Row As Long)
    Dim nWeekDay        As Long
    Dim sTrxCD          As String
    
    If Trim(txtTrxCD_S.Text) = "" Then
        MsgBox "����� ������ ��� �������忡�� �����ϼ���.", vbExclamation + vbOKOnly, "������ �ð�ǥ ���"
        Exit Sub
    End If
    
    If Trim(txtKaeyol_S.Text) = "" Then
        MsgBox "����� ������ ��� �������忡�� �����ϼ���.", vbExclamation + vbOKOnly, "������ �ð�ǥ ���"
        Exit Sub
    End If
    
    If Row < 1 Or Col < 1 Then
        MsgBox "����� ���ϰ� ���ø� �����ϼ���.", vbExclamation + vbOKOnly, "������ �ð�ǥ ���"
        Exit Sub
    End If
    
    On Error GoTo ErrStmt
    
    With sprTrxType
        
        '>> ���ϼ���
        Select Case Col
            Case 1
                nWeekDay = 2
            Case 2
                nWeekDay = 3
            Case 3
                nWeekDay = 4
            Case 4
                nWeekDay = 5
            Case 5
                nWeekDay = 6
            Case 6
                nWeekDay = 7
            Case 7
                nWeekDay = 1
        End Select
            
        
    '## ������ �ð�ǥ ���
        If optSaveChk.Value = True Then
            Select Case Find_Early_Save_Data(Left(Trim(txtTrxCD_S.Text), 1), Row, nWeekDay)
            '>> �ð�ǥ ���
                Case "IN"
                    If Save_Setting_Data(Trim(txtTrxCD_S.Text), Trim(txtKaeyol_S.Text), Row, nWeekDay) = True Then
                        cmdNewPB.Tag = "SAVE"
                            Call cmdFindAll_Click
                            Call Find_TrxPart_Detail(Left(Trim(txtTrxCD_S.Text), 1))        '< ������ȸ
                        cmdNewPB.Tag = ""
                        
                    End If
                Case "NOT"
                    ' no action
            End Select
                
    '## ������ �ð�ǥ ����
        ElseIf optDelChk.Value = True Then
            
            sTrxCD = Find_Delete_TRX_Data(Left(Trim(txtTrxCD_S.Text), 1), Row, nWeekDay)        '< ������ ������ �ð�ǥ �ڵ�
            
            If sTrxCD > " " Then
                If Delete_Setting_Data(sTrxCD, Trim(txtKaeyol_S.Text), Row, nWeekDay) = True Then
                    cmdNewPB.Tag = "DELETE"
                        Call cmdFindAll_Click
                        Call Find_TrxPart_Detail(Left(Trim(txtTrxCD_S.Text), 1))        '< ������ȸ
                    cmdNewPB.Tag = ""
                
                End If
            End If
        End If
    End With
    
    Exit Sub
ErrStmt:
    On Error GoTo 0
    
    MsgBox "ó���� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "�ð����� ��� �� ����"
    
End Sub





'===============================================================================================================================================================
'## ��ϵ� ���� ��ȸ
Private Function Find_Delete_TRX_Data(ByVal aGbn As String, ByVal aLesson As Integer, ByVal aWeek As Integer) As String
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    
    Dim sStr        As String
    Dim sRet        As String
    
    Dim ni          As Long
    
    On Error Resume Next
    
    sStr = ""
    sStr = sStr & "  SELECT B.ACID, B.TRXCD, A.TRXNM, A.KAEYOL"
    sStr = sStr & "    FROM SDTRX01TB A, SDTRX11TB B"
    sStr = sStr & "   WHERE A.ACID   = B.ACID "
    sStr = sStr & "     AND A.TRXCD  = B.TRXCD"
    sStr = sStr & "     AND A.KAEYOL = B.KAEYOL"
    sStr = sStr & "     AND A.ACID   = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "     AND A.KAEYOL = '" & Trim(Right(cboKaeyol.Text, 30)) & "'"
    sStr = sStr & "     AND B.LESSON = " & Trim(CStr(aLesson))
    sStr = sStr & "     AND B.WEEKS  = " & Trim(CStr(aWeek))
    sStr = sStr & "     AND B.TRXCD  LIKE '" & Trim(aGbn) & "%'"
    
    
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
'    '>> lesson
'        nTmp = aLesson
'            Set DBParam = DBCmd.CreateParameter("LESSON", adDouble, adParamInput, , nTmp):   DBCmd.Parameters.Append DBParam
'    '>> week
'        nTmp = aWeek
'            Set DBParam = DBCmd.CreateParameter("WEEKS", adDouble, adParamInput, , nTmp):   DBCmd.Parameters.Append DBParam

    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic���·� ���ԵǸ� record count�� �� �� ����.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop
    
    sRet = ""
    With DBRec
        
        Select Case .RecordCount
            Case 0
                sRet = ""
                
            Case Is = 1
                .MoveFirst
                    
                sRet = ""
                If IsNull(.Fields("TRXCD")) = False Then sRet = Trim(.Fields("TRXCD"))
                
        End Select
    End With
    
    Find_Delete_TRX_Data = sRet

End Function

'>> ���� ��ϵ� ������ �ִ��� Ȯ����.
Private Function Find_Early_Save_Data(ByVal aGbn As String, ByVal aLesson As Integer, ByVal aWeek As Integer) As String
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    
    Dim ni          As Integer
    Dim nRec        As Long
    
    Dim sStr        As String
    Dim sRet        As String
    Dim sTmp        As String
    Dim sWeekday    As String
    
    Dim sChks       As String
    
    On Error Resume Next
    
    sStr = ""
    sStr = sStr & "  SELECT B.ACID, B.TRXCD, A.TRXNM, A.KAEYOL"
    sStr = sStr & "    FROM SDTRX01TB A, SDTRX11TB B"
    sStr = sStr & "   WHERE A.ACID   = B.ACID "
    sStr = sStr & "     AND A.TRXCD  = B.TRXCD"
    sStr = sStr & "     AND A.KAEYOL = B.KAEYOL"
    sStr = sStr & "     AND A.ACID   = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "     AND A.KAEYOL = '" & Trim(Right(cboKaeyol.Text, 30)) & "'"
    sStr = sStr & "     AND B.LESSON = " & Trim(CStr(aLesson))
    sStr = sStr & "     AND B.WEEKS  = " & Trim(CStr(aWeek))
            
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
'    '>> lesson
'        nTmp = aLesson
'            Set DBParam = DBCmd.CreateParameter("LESSON", adDouble, adParamInput, , nTmp):   DBCmd.Parameters.Append DBParam
'    '>> week
'        nTmp = aWeek
'            Set DBParam = DBCmd.CreateParameter("WEEKS", adDouble, adParamInput, , nTmp):   DBCmd.Parameters.Append DBParam

    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic���·� ���ԵǸ� record count�� �� �� ����.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop
    
    sRet = "NOT"
    sChks = ""
    
    With DBRec
        
        Select Case .RecordCount
            Case 0
                sRet = "IN"
                
            Case Is > 0
                .MoveFirst
                    
                sTmp = ""
                For nRec = 1 To .RecordCount Step 1
                    If IsNull(.Fields("TRXNM")) = False Then
                        If nRec > 1 Then sTmp = sTmp & vbCrLf
                        sTmp = sTmp & Trim(.Fields("TRXNM"))
                        
                        If StrComp(aGbn, Left(Trim(.Fields("TRXNM")), 1), vbTextCompare) = 0 Then
                            sChks = "CANCEL"
                            Exit For
                        End If
                    End If
                    
                    .MoveNext
                Next nRec
                
                If sChks = "CANCEL" Then
                    Select Case aWeek
                        Case 2
                            sWeekday = "��"
                        Case 3
                            sWeekday = "ȭ"
                        Case 4
                            sWeekday = "��"
                        Case 5
                            sWeekday = "��"
                        Case 6
                            sWeekday = "��"
                        Case 7
                            sWeekday = "��"
                        Case 1
                            sWeekday = "��"
                    End Select
                    
                    sTmp = sWeekday & "���� - " & Trim(CStr(aLesson)) & "����" & vbCrLf
                    MsgBox sTmp & _
                           "'" & aGbn & "'" & " ���� ������ ���� �ð�ǥ�� �����Ƿ� ����� �� �����ϴ�.", vbExclamation + vbOKOnly, "���� ��ϳ��� ��ȸ"
                           
                    sRet = "NOT"
                Else
                    Select Case aWeek
                        Case 2
                            sWeekday = "��"
                        Case 3
                            sWeekday = "ȭ"
                        Case 4
                            sWeekday = "��"
                        Case 5
                            sWeekday = "��"
                        Case 6
                            sWeekday = "��"
                        Case 7
                            sWeekday = "��"
                        Case 1
                            sWeekday = "��"
                    End Select
                    
                    sTmp = sWeekday & "���� - " & Trim(CStr(aLesson)) & "����" & vbCrLf & vbCrLf & sTmp
                    If MsgBox(sTmp & vbCrLf & "������ �ֽ��ϴ�. �����Ͻðڽ��ϱ�?", vbQuestion + vbYesNo, "���� ��ϳ��� ��ȸ") = vbYes Then
                        sRet = "IN"
                    Else
                        sRet = "NOT"
                    End If
                End If
        End Select
    End With
    
    Find_Early_Save_Data = sRet

End Function


'>> �����.
Private Function Save_Setting_Data(ByVal aTrxCD As String, ByVal aKaeyol As String, ByVal aLesson As Integer, ByVal aWeek As Integer) As Boolean
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    
    Dim sStr        As String
    Dim bRet        As Boolean
    Dim sTmp        As String
    
    Dim ni          As Integer
    Dim nExe        As Integer
    
    
    bRet = False
    
    On Error GoTo ErrStmt
    
    basDataBase.DBConn.BeginTrans
    
    Set DBCmd = New ADODB.Command
    Set DBRec = New ADODB.Recordset
    Set DBParam = New ADODB.Parameter
    
    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
    
    sStr = ""
    sStr = sStr & "  INSERT INTO SDTRX11TB (ACID, TRXCD, KAEYOL, LESSON, WEEKS)"
    sStr = sStr & "  VALUES("
    sStr = sStr & "                 '" & Trim(basModule.SchCD) & "',"
    sStr = sStr & "                 '" & aTrxCD & "',"
    sStr = sStr & "                 '" & aKaeyol & "',"
    sStr = sStr & "                  " & Trim(CStr(aLesson)) & ","
    sStr = sStr & "                  " & Trim(CStr(aWeek))
    sStr = sStr & "         )"
    



'    '>> �п�
'        sTmp = Trim(basModule.SchCD)
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)): DBCmd.Parameters.Append DBParam
'    '>> ������ �ð�ǥ ����
'        sTmp = aTrxCD
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("TRXCD", adChar, adParamInput, nLength, Trim(sTmp)): DBCmd.Parameters.Append DBParam
'    '>> LESSON
'        nTmp = aLesson
'            Set DBParam = DBCmd.CreateParameter("LESSON", adDouble, adParamInput, , nTmp):   DBCmd.Parameters.Append DBParam
'    '>> WEEKS
'        nTmp = aWeek
'            Set DBParam = DBCmd.CreateParameter("WEEKS", adDouble, adParamInput, , nTmp):   DBCmd.Parameters.Append DBParam

    DBCmd.CommandText = sStr
    DBCmd.CommandType = adCmdText
    DBCmd.CommandTimeout = 30

    nExe = 0
    DBCmd.Execute nExe, , -1

    Do While basDataBase.DBConn.State And adStateExecuting
        DoEvents
    Loop

    If nExe = 1 Then
        basDataBase.DBConn.CommitTrans
        bRet = True
    Else
        basDataBase.DBConn.RollbackTrans
        bRet = False
    End If
        
        
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    Save_Setting_Data = bRet
    
    Exit Function
ErrStmt:
    
    If Err.Number = -2147217900 Then
        'MsgBox "update"
        On Error GoTo 0
        On Error GoTo ErrUpdate
        
        sStr = ""
        sStr = sStr & "  DELETE "
        sStr = sStr & "    FROM SDTRX11TB "
        sStr = sStr & "   WHERE ACID   = '" & Trim(basModule.SchCD) & "'"
        sStr = sStr & "     AND TRXCD  = '" & aTrxCD & "'"
        sStr = sStr & "     AND KAEYOL = '" & aKaeyol & "'"
        sStr = sStr & "     AND LESSON = " & Trim(CStr(aLesson))
        sStr = sStr & "     AND WEEKS  = " & Trim(CStr(aWeek))
        


    
    '    '>> �п�
    '        sTmp = Trim(basModule.SchCD)
    '        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
    '            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)): DBCmd.Parameters.Append DBParam
    '    '>> ������ �ð�ǥ ����
    '        sTmp = aTrxCD
    '        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
    '            Set DBParam = DBCmd.CreateParameter("TRXCD", adChar, adParamInput, nLength, Trim(sTmp)): DBCmd.Parameters.Append DBParam
    '    '>> LESSON
    '        nTmp = aLesson
    '            Set DBParam = DBCmd.CreateParameter("LESSON", adDouble, adParamInput, , nTmp):   DBCmd.Parameters.Append DBParam
    '    '>> WEEKS
    '        nTmp = aWeek
    '            Set DBParam = DBCmd.CreateParameter("WEEKS", adDouble, adParamInput, , nTmp):   DBCmd.Parameters.Append DBParam
    
        DBCmd.CommandText = sStr
        DBCmd.CommandType = adCmdText
        DBCmd.CommandTimeout = 30
    
        nExe = 0
        DBCmd.Execute nExe, , -1
    
        Do While basDataBase.DBConn.State And adStateExecuting
            DoEvents
        Loop
    
        If nExe = 1 Then
            
            sStr = ""
            sStr = sStr & "  INSERT INTO SDTRX11TB (ACID, TRXCD, KAEYOL, LESSON, WEEKS)"
            sStr = sStr & "  VALUES("
            sStr = sStr & "                 '" & Trim(basModule.SchCD) & "',"
            sStr = sStr & "                 '" & aTrxCD & "',"
            sStr = sStr & "                 '" & aKaeyol & "',"
            sStr = sStr & "                  " & Trim(CStr(aLesson)) & ","
            sStr = sStr & "                  " & Trim(CStr(aWeek))
            sStr = sStr & "         )"
            


        '    '>> �п�
        '        sTmp = Trim(basModule.SchCD)
        '        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
        '            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)): DBCmd.Parameters.Append DBParam
        '    '>> ������ �ð�ǥ ����
        '        sTmp = aTrxCD
        '        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
        '            Set DBParam = DBCmd.CreateParameter("TRXCD", adChar, adParamInput, nLength, Trim(sTmp)): DBCmd.Parameters.Append DBParam
        '    '>> LESSON
        '        nTmp = aLesson
        '            Set DBParam = DBCmd.CreateParameter("LESSON", adDouble, adParamInput, , nTmp):   DBCmd.Parameters.Append DBParam
        '    '>> WEEKS
        '        nTmp = aWeek
        '            Set DBParam = DBCmd.CreateParameter("WEEKS", adDouble, adParamInput, , nTmp):   DBCmd.Parameters.Append DBParam
        
            DBCmd.CommandText = sStr
            DBCmd.CommandType = adCmdText
            DBCmd.CommandTimeout = 30
        
            nExe = 0
            DBCmd.Execute nExe, , -1
        
            Do While basDataBase.DBConn.State And adStateExecuting
                DoEvents
            Loop
        
            If nExe = 1 Then
                basDataBase.DBConn.CommitTrans
                bRet = True
            Else
                basDataBase.DBConn.RollbackTrans
                bRet = False
            End If
        Else
            basDataBase.DBConn.RollbackTrans
        
            Set DBCmd = Nothing
            Set DBRec = Nothing
            
            On Error GoTo 0
            MsgBox "������ �ð����� ��Ͻ� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "������ �ð����� ���"
        End If
        
    Else
        basDataBase.DBConn.RollbackTrans
    
        Set DBCmd = Nothing
        Set DBRec = Nothing
        
        On Error GoTo 0
        MsgBox "������ �ð����� ��Ͻ� ������ �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "������ �ð����� ���"
        
    End If
    
    Save_Setting_Data = bRet
    
    Exit Function
ErrUpdate:
    basDataBase.DBConn.RollbackTrans
    
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    On Error GoTo 0
    MsgBox "������ �ð����� ��Ͽ����� �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "������ �ð����� ���"
    
    Save_Setting_Data = bRet
        
End Function


'>> ������
Private Function Delete_Setting_Data(ByVal aTrxCD As String, ByVal aKaeyol As String, ByVal aLesson As Integer, ByVal aWeek As Integer) As Boolean
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    
    Dim sStr        As String
    Dim bRet        As Boolean
    Dim sTmp        As String
    
    Dim ni          As Integer
    Dim nExe        As Integer
    
    bRet = False
    
    On Error GoTo ErrStmt
    
    basDataBase.DBConn.BeginTrans
    
    Set DBCmd = New ADODB.Command
    Set DBRec = New ADODB.Recordset
    Set DBParam = New ADODB.Parameter
    
    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
    
    sStr = ""
    sStr = sStr & "  DELETE "
    sStr = sStr & "    FROM SDTRX11TB "
    sStr = sStr & "   WHERE ACID   = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "     AND TRXCD  = '" & aTrxCD & "'"
    sStr = sStr & "     AND KAEYOL = '" & aKaeyol & "'"
    sStr = sStr & "     AND LESSON = " & Trim(CStr(aLesson))
    sStr = sStr & "     AND WEEKS  = " & Trim(CStr(aWeek))
    



'    '>> �п�
'        sTmp = Trim(basModule.SchCD)
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)): DBCmd.Parameters.Append DBParam
'    '>> ������ �ð�ǥ ����
'        sTmp = aTrxCD
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("TRXCD", adChar, adParamInput, nLength, Trim(sTmp)): DBCmd.Parameters.Append DBParam
'    '>> LESSON
'        nTmp = aLesson
'            Set DBParam = DBCmd.CreateParameter("LESSON", adDouble, adParamInput, , nTmp):   DBCmd.Parameters.Append DBParam
'    '>> WEEKS
'        nTmp = aWeek
'            Set DBParam = DBCmd.CreateParameter("WEEKS", adDouble, adParamInput, , nTmp):   DBCmd.Parameters.Append DBParam

    DBCmd.CommandText = sStr
    DBCmd.CommandType = adCmdText
    DBCmd.CommandTimeout = 30

    nExe = 0
    DBCmd.Execute nExe, , -1

    Do While basDataBase.DBConn.State And adStateExecuting
        DoEvents
    Loop

    If nExe = 1 Then
        basDataBase.DBConn.CommitTrans
        bRet = True
    Else
        basDataBase.DBConn.RollbackTrans
        bRet = False
    End If
    
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    Delete_Setting_Data = bRet
    
    Exit Function
ErrStmt:
    basDataBase.DBConn.RollbackTrans
    
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    On Error GoTo 0
    MsgBox "������ �ð����� ���������� �߻��Ͽ����ϴ�.", vbCritical + vbOKOnly, "������ �ð����� ����"
    
    Delete_Setting_Data = bRet
        
End Function

'===============================================================================================================================================================



Private Sub sprKeyiN_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyInsert
            '> ���� �߰�
            With sprKeyiN
                .MaxRows = .MaxRows + 1
                .Row = .MaxRows
            
                .Col = 1:   Call basFunction.Set_SprType_Text(sprKeyiN, "center", "center", 2, ""):         .Lock = False
                .Col = 2:   Call basFunction.Set_SprType_Text(sprKeyiN, "center", "center", 1, ""):         .Lock = False
                .Col = 3:   Call basFunction.Set_SprType_Text(sprKeyiN, "center", "center", 50, ""):        .Lock = False
                
                .Col = 4:   Call basFunction.Set_SprType_ChkBox(sprKeyiN):          .Lock = True
                
            End With
            
        Case vbKeyDelete
            '> ���� ����
            With sprKeyiN
                If sprKeyiN.MaxRows = 0 Then Exit Sub
                
                sprKeyiN.DeleteRows sprKeyiN.ActiveRow, 1
                sprKeyiN.MaxRows = sprKeyiN.MaxRows - 1
            End With
            
        Case vbKeyReturn
            With sprKeyiN
                .SetActiveCell 1, .ActiveRow
                
            End With
        
    End Select
End Sub





'## ����ϱ�
Private Sub cmdKeyiN_Time_Click()
    Dim nRow        As Long
    
    Dim nWeekDay    As Integer
    Dim nLesson     As Integer
    Dim sTrxCD      As String
    Dim sKaeyol     As String
    Dim nChk        As Integer
    
    Dim sTmp        As String
    Dim sSaveOK     As String
    
    sSaveOK = ""
    
    With sprKeyiN
        If .MaxRows = 0 Then Exit Sub
        
        For nRow = 1 To .MaxRows Step 1
        
            sKaeyol = Trim(Right(cboKaeyol_All.Text, 30))       '< �迭
        
            .Row = nRow
            .Col = 1
                '>> ���ϼ���
                Select Case Trim(.Text)
                    Case "��"
                        nWeekDay = 2
                    Case "ȭ"
                        nWeekDay = 3
                    Case "��"
                        nWeekDay = 4
                    Case "��"
                        nWeekDay = 5
                    Case "��"
                        nWeekDay = 6
                    Case "��"
                        nWeekDay = 7
                    Case "��"
                        nWeekDay = 1
                End Select
            .Col = 2
                If Trim(.Text) <> "" Then nLesson = Trim(.Text)
            .Col = 3
                sTrxCD = Get_TrxCD(Trim(basModule.SchCD), sKaeyol, UCase(Trim(.Text)))
            .Col = 4
                nChk = 1
                nChk = .Value
                
                        
                Select Case UCase(Left(Trim(sTrxCD), 2))
                    Case "A1", "A2", "B1", "B2", "C1", "C2", "PB"
                        Select Case nLesson
                            Case 1 To 10
                                Select Case nWeekDay
                                    Case 1 To 7
                                        If nChk = 0 Then
                                            '>> save ok -------------------------------------------------------------------------------------------------
                                            Select Case Find_Early_Save_Data(Left(Trim(sTrxCD), 1), nLesson, nWeekDay)
                                            '>> �ð�ǥ ���
                                                Case "IN"
                                                    If Save_Setting_Data(Trim(sTrxCD), sKaeyol, nLesson, nWeekDay) = True Then
                                                        .Row = nRow
                                                        .Col = .MaxCols
                                                            .Value = 1
                                                    End If
                                                Case "NOT"
                                                    ' no action
                                            End Select
                                            '------------------------------------------------------------------------------------------------------------
                                        End If
                                    Case Else
                                        'not
                                End Select
                            Case Else
                                'not
                        End Select
                    Case Else
                        'not
                End Select
        Next nRow
    End With
    
    
    '## ��ϳ��� ��ȸ
    cmdNewPB.Tag = "SAVE"
        Call cmdFindAll_Click
    cmdNewPB.Tag = ""
    
End Sub

Private Sub cmdKeyNew_Time_Click()
    sprKeyiN.MaxRows = 0
End Sub

Private Function Get_TrxCD(ByVal aAcID As String, ByVal aKaeyol As String, ByVal aTrxNM As String)
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    
    Dim sStr        As String
    Dim sRet        As String
    
    Dim ni          As Long
    
    On Error Resume Next
    
    sStr = ""
    sStr = sStr & "  SELECT TRXCD, TRXNM"
    sStr = sStr & "    From SDTRX01TB"
    sStr = sStr & "   WHERE ACID   = '" & Trim(aAcID) & "'"
    sStr = sStr & "     AND KAEYOL = '" & Trim(aKaeyol) & "'"
    sStr = sStr & "     AND REPLACE(TRXNM,' ','')  LIKE '" & Trim(aTrxNM) & "%'"
    
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
'    '>> lesson
'        nTmp = aLesson
'            Set DBParam = DBCmd.CreateParameter("LESSON", adDouble, adParamInput, , nTmp):   DBCmd.Parameters.Append DBParam
'    '>> week
'        nTmp = aWeek
'            Set DBParam = DBCmd.CreateParameter("WEEKS", adDouble, adParamInput, , nTmp):   DBCmd.Parameters.Append DBParam

    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic���·� ���ԵǸ� record count�� �� �� ����.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop
    
    sRet = ""
    With DBRec
        
        Select Case .RecordCount
            Case 0
                sRet = ""
                
            Case Is = 1
                .MoveFirst
                    
                sRet = ""
                If IsNull(.Fields("TRXCD")) = False Then sRet = Trim(.Fields("TRXCD"))
                
        End Select
    End With
    
    Get_TrxCD = sRet
    
End Function































