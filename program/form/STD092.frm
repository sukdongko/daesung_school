VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form STD092 
   Caption         =   "입학사정 >> 학생취소자 조회"
   ClientHeight    =   9720
   ClientLeft      =   1620
   ClientTop       =   2475
   ClientWidth     =   15330
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9720
   ScaleWidth      =   15330
   Begin MSComDlg.CommonDialog dlgFile 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame18 
      BackColor       =   &H0082C8E8&
      BorderStyle     =   0  '없음
      Caption         =   "Frame18"
      Height          =   9465
      Left            =   60
      TabIndex        =   11
      Top             =   30
      Width           =   15015
      Begin VB.Frame Frame19 
         BackColor       =   &H00D2EAF5&
         BorderStyle     =   0  '없음
         Caption         =   "Frame19"
         Height          =   9405
         Left            =   30
         TabIndex        =   12
         Top             =   30
         Width           =   14955
         Begin VB.CommandButton cmdAllStdData 
            Caption         =   "현재자료 엑셀로 받기"
            Height          =   435
            Left            =   12930
            TabIndex        =   8
            Top             =   60
            Width           =   1965
         End
         Begin VB.ComboBox cboinGbn 
            Height          =   300
            Left            =   5400
            Style           =   2  '드롭다운 목록
            TabIndex        =   3
            Top             =   127
            Width           =   885
         End
         Begin VB.ComboBox cboExmType 
            Height          =   300
            Left            =   4020
            Style           =   2  '드롭다운 목록
            TabIndex        =   2
            Top             =   127
            Width           =   855
         End
         Begin VB.ComboBox cboKaeyol_F 
            Height          =   300
            Left            =   2190
            Style           =   2  '드롭다운 목록
            TabIndex        =   1
            Top             =   127
            Width           =   915
         End
         Begin VB.CommandButton cmdFind 
            Caption         =   "조회하기(&F)"
            Height          =   480
            Left            =   390
            TabIndex        =   0
            Top             =   37
            Width           =   1245
         End
         Begin VB.TextBox txtStdNM_F 
            Height          =   345
            IMEMode         =   10  '한글 
            Left            =   9720
            TabIndex        =   6
            Text            =   "txtStdNM_F"
            Top             =   105
            Width           =   825
         End
         Begin FPSpread.vaSpread sprSTD_F 
            Height          =   8835
            Left            =   60
            TabIndex        =   9
            Top             =   570
            Width           =   14895
            _Version        =   393216
            _ExtentX        =   26273
            _ExtentY        =   15584
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
            MaxCols         =   34
            SpreadDesigner  =   "STD092.frx":0000
         End
         Begin EditLib.fpMask fpExmID_F 
            Height          =   345
            Left            =   7170
            TabIndex        =   4
            Top             =   105
            Width           =   765
            _Version        =   196608
            _ExtentX        =   1349
            _ExtentY        =   609
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
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
            Left            =   11400
            TabIndex        =   7
            Top             =   105
            Width           =   1215
            _Version        =   196608
            _ExtentX        =   2143
            _ExtentY        =   609
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
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
            Left            =   8310
            TabIndex        =   5
            Top             =   105
            Width           =   735
            _Version        =   196608
            _ExtentX        =   1296
            _ExtentY        =   609
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
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
         Begin VB.Label Label24 
            BackStyle       =   0  '투명
            Caption         =   ">> 조회기본항목"
            BeginProperty Font 
               Name            =   "굴림"
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
            TabIndex        =   19
            Top             =   135
            Width           =   1545
         End
         Begin VB.Label Label37 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "접수"
            Height          =   210
            Left            =   4890
            TabIndex        =   18
            Top             =   172
            Width           =   465
         End
         Begin VB.Label Label36 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "무/유시험"
            Height          =   210
            Left            =   3000
            TabIndex        =   17
            Top             =   172
            Width           =   975
         End
         Begin VB.Label Label31 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "계열"
            Height          =   210
            Left            =   1590
            TabIndex        =   16
            Top             =   172
            Width           =   525
         End
         Begin VB.Label Label27 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "생년월일"
            Height          =   210
            Left            =   10410
            TabIndex        =   15
            Top             =   172
            Width           =   975
         End
         Begin VB.Label Label26 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "학생명"
            Height          =   210
            Left            =   8670
            TabIndex        =   14
            Top             =   172
            Width           =   975
         End
         Begin VB.Label Label25 
            BackStyle       =   0  '투명
            Caption         =   "수험번호              부터"
            Height          =   210
            Left            =   6390
            TabIndex        =   13
            Top             =   172
            Width           =   2025
         End
      End
   End
   Begin FPSpread.vaSpread sprStdData 
      Height          =   165
      Left            =   0
      TabIndex        =   10
      Top             =   9330
      Width           =   9765
      _Version        =   393216
      _ExtentX        =   17224
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
      SpreadDesigner  =   "STD092.frx":20A9
   End
End
Attribute VB_Name = "STD092"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################################################
'   시 스 템  명 : 대성학원 입학사정, 반배정 & 시간표 프로그램
'   서브시스템명 :
'   모   듈   명 : STD011
'   모 듈  목 적 : 학생전체 조회
'
'   작   성   일 : 2007/12/13
'   작   성   자 : 유하균
' --------------------------------------------------------------------------------------------------------------
' --------------------------------------------------------------------------------------------------------------
'                 수     정     내     용
' --------------------------------------------------------------------------------------------------------------
'   1. 수정일 :
'   2. 내  용 :
'################################################################################################################

Option Explicit

Private Sub Form_Terminate()
    Unload Me
End Sub

Private Sub Form_Load()
    
    Me.Move 0, 0, 15255, 9980
    
    
    With sprSTD_F
        .ShadowColor = basModule.ShadowColor1
        .ShadowDark = basModule.ShadowDark1
        .ShadowText = basModule.ShadowText1
        .GridColor = basModule.GridColor1
        .GrayAreaBackColor = basModule.GrayAreaBackColor1
    End With
    
    With cboKaeyol_F
        .Clear
        .AddItem "전체" & Space(30) & "ALL"
        .AddItem "인문" & Space(30) & "01"
        .AddItem "자연" & Space(30) & "02"
        
    '<< 계열 >> : 2008.01.09
        If Trim(basModule.SchCD) = "N" Then             '< 노량진
            .AddItem "예체" & Space(30) & "03"
            .AddItem "수리(나)" & Space(30) & "04"
            .AddItem "인문수능" & Space(30) & "05"
            .AddItem "자연수능" & Space(30) & "06"
            
            .AddItem "인문-신" & Space(30) & "07"
            .AddItem "자연-신" & Space(30) & "08"
            '.AddItem "수능인문-신" & Space(30) & "09"
            '.AddItem "수능자연-신" & Space(30) & "10"
            
            .AddItem "편)인문" & Space(30) & "11"
            .AddItem "편)자연" & Space(30) & "12"
            .AddItem "편)예체" & Space(30) & "13"
            .AddItem "편)수리(나)" & Space(30) & "14"
            .AddItem "편)인문수능" & Space(30) & "15"
            .AddItem "편)자연수능" & Space(30) & "16"
        End If
    '<< 계열 >> : 2008.01.10
        
        If Trim(basModule.SchCD) = "K" Or Trim(basModule.SchCD) = "W" Or Trim(basModule.SchCD) = "Q" Then       '< 강남 2008.03.24
            .AddItem "주말법대" & Space(30) & "04"
            .AddItem "주말의대" & Space(30) & "05"
            
            .AddItem "야간법대" & Space(30) & "06"
            .AddItem "야간의대" & Space(30) & "07"
            
            .AddItem "선착순인문" & Space(30) & "11"
            .AddItem "선착순자연" & Space(30) & "12"
            
            .AddItem "선착순인문16" & Space(30) & "16"
            .AddItem "선착순자연17" & Space(30) & "17"
        End If
        
    '<< 계열 >> : 2008.01.09
        Select Case Trim(basModule.SchCD)
            Case "S"                                        '< 송파
'                .AddItem "예체능" & Space(30) & "03"
'
'                .AddItem "인문수능" & Space(30) & "05"
'                .AddItem "자연수능" & Space(30) & "06"
'
'                .AddItem "신설인문" & Space(30) & "11"
'                .AddItem "신설자연" & Space(30) & "12"

                .AddItem "인문프리미엄" & Space(30) & "18"
                .AddItem "자연프리미엄" & Space(30) & "19"

        End Select
        
        Select Case Trim(basModule.SchCD)
            Case "P", "J"                                   '< 양재
                .AddItem "신설인문" & Space(30) & "11"
                .AddItem "신설자연" & Space(30) & "12"
        End Select
        
        
    '<< 계열 >> : 2009.01.09
        Select Case Trim(basModule.SchCD)
            Case "B"                                    '< 부산
                .AddItem "수학선행인문" & Space(30) & "05"
                .AddItem "수학선행자연" & Space(30) & "06"
                
                .AddItem "연.고대인문" & Space(30) & "07"
                .AddItem "연.고대자연" & Space(30) & "08"
                
                .AddItem "심화인문" & Space(30) & "09"
                .AddItem "심화자연" & Space(30) & "10"

        End Select
        
        .ListIndex = 0
    End With
    
    With cboExmType
        .Clear
        .AddItem "전체" & Space(30) & "ALL"
        .AddItem "유시험" & Space(30) & "1"
        .AddItem "무시험" & Space(30) & "0"
        
        .ListIndex = 0
    End With
    
    With cboinGbn
        .Clear
        .AddItem "전체" & Space(30) & "ALL"
        .AddItem "인터넷" & Space(30) & "INT"
        .AddItem "학원" & Space(30) & "HAK"
        
        .ListIndex = 0
    End With
    
    Call init_Form
    
    
End Sub

Private Sub init_Form()
    Dim ni      As Integer
    
    
    '>> 조회부분
    fpExmID_F.Text = ""
    fpExmID_E.Text = ""
    
    txtStdNM_F.Text = ""
    fpBirth_ymd_F.Text = ""
    sprSTD_F.MaxRows = 0
    
End Sub




'>> 학생 조회하기
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
    sStr = sStr & "  SELECT A.SCHNO, A.EXMID, STDNM, SEL1_SCH , SEL2_SCH, SUBSTR(REPLACE(Birth_ymd,'-',''),1,4)||'-'||SUBSTR(REPLACE(Birth_ymd,'-',''),5,2) ||'-'||SUBSTR(REPLACE(Birth_ymd,'-',''),7,2) AS Birth_ymd,"
    
    '<< 계열 >> : 2008.01.09
    If Trim(basModule.SchCD) = "N" Then
        sStr = sStr & "     DECODE(KAEYOL,'01','인문',"
        sStr = sStr & "                   '02','자연',"
        sStr = sStr & "                   '03','예체',"
        sStr = sStr & "                   '04','수리(나)',"
        sStr = sStr & "                   '05','인문수능',"
        sStr = sStr & "                   '06','자연수능',"
        sStr = sStr & "                   '07','신설인문',"
        sStr = sStr & "                   '08','신설자연',"
        sStr = sStr & "                   '09','신설수능인문',"
        sStr = sStr & "                   '10','신설수능자연',"
        
        sStr = sStr & "                   '11','편)인문',"
        sStr = sStr & "                   '12','편)자연',"
        sStr = sStr & "                   '13','편)예체',"
        sStr = sStr & "                   '14','편)수리(나)',"
        sStr = sStr & "                   '15','편)인문수능',"
        sStr = sStr & "                   '16','편)자연수능'"
        sStr = sStr & "            ) AS GAEYUL,"
'<< 계열 >> : 2008.01.10
    ElseIf Trim(basModule.SchCD) = "K" Or Trim(basModule.SchCD) = "W" Or Trim(basModule.SchCD) = "Q" Then       '< 2008.03.24
        sStr = sStr & "     DECODE(KAEYOL,'01','인문',"
        sStr = sStr & "                   '02','자연',"
        
        sStr = sStr & "                   '04','주말법대',"
        sStr = sStr & "                   '05','주말의대',"
        sStr = sStr & "                   '06','야간법대',"
        sStr = sStr & "                   '07','야간의대',"
        
        sStr = sStr & "                   '11','선착순인문',"
        sStr = sStr & "                   '12','선착순자연',"
        
        sStr = sStr & "                   '16','선착순인문16',"
        sStr = sStr & "                   '17','선착순자연17'"
        sStr = sStr & "            ) AS GAEYUL,"
    ElseIf Trim(basModule.SchCD) = "S" Then
        sStr = sStr & "     DECODE(KAEYOL,'01','인문',"
        sStr = sStr & "                   '02','자연',"
        sStr = sStr & "                   '03','예체능',"
        sStr = sStr & "                   '05','수능인문',"
        sStr = sStr & "                   '06','수능자연',"
        
        sStr = sStr & "                   '11','신설인문',"
        sStr = sStr & "                   '12','신설자연'"
        
        sStr = sStr & "            ) AS GAEYUL,"
        
    ElseIf Trim(basModule.SchCD) = "P" Then
        sStr = sStr & "     DECODE(KAEYOL,'01','인문',"
        sStr = sStr & "                   '02','자연',"
        sStr = sStr & "                   '03','특별인문',"
        sStr = sStr & "                   '04','특별자연'"
        sStr = sStr & "            ) AS GAEYUL,"
    
    ElseIf Trim(basModule.SchCD) = "J" Then
        sStr = sStr & "     DECODE(KAEYOL,'01','인문',"
        sStr = sStr & "                   '02','자연',"
        sStr = sStr & "                   '11','신설인문',"
        sStr = sStr & "                   '12','신설자연'"
        sStr = sStr & "            ) AS GAEYUL,"
        
    ElseIf Trim(basModule.SchCD) = "B" Then
        sStr = sStr & "     DECODE(KAEYOL,'01','인문',"
        sStr = sStr & "                   '02','자연',"
        sStr = sStr & "                   '05','특별인문',"
        sStr = sStr & "                   '06','특별자연',"
        sStr = sStr & "                   '07','연고대인문',"
        sStr = sStr & "                   '08','연고대자연',"
        sStr = sStr & "                   '09','심화인문',"
        sStr = sStr & "                   '10','심화자연'"
        sStr = sStr & "            ) AS GAEYUL,"
        
    Else
        sStr = sStr & "     DECODE(KAEYOL,'01','인문',"
        sStr = sStr & "                   '02','자연'"
        sStr = sStr & "            ) AS GAEYUL,"
    End If
    
    sStr = sStr & "     /* 사탐, 과탐 분리 */"
    sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'" & constSatamCodes(0) & "|') > 0 THEN          /* 사탐-한국사 */"
    sStr = sStr & "             '" & constSatamCodes(0) & "'"
    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'51|') > 0 THEN     /* 과탐-물리1 */"
    sStr = sStr & "             '51'"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             '00'"
    sStr = sStr & "         END END SEL1,"
    sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'" & constSatamCodes(1) & "|') > 0 THEN          /* 사탐-세계사 */"
    sStr = sStr & "             '" & constSatamCodes(1) & "'"
    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'52|') > 0 THEN     /* 과탐-화학1 */"
    sStr = sStr & "             '52'"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             '00'"
    sStr = sStr & "         END END SEL2,"
    sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'" & constSatamCodes(2) & "|') > 0 THEN          /* 사탐-동아시아사 */"
    sStr = sStr & "             '" & constSatamCodes(2) & "'"
    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'53|') > 0 THEN     /* 과탐-생명과학1 */"
    sStr = sStr & "             '53'"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             '00'"
    sStr = sStr & "         END END SEL3,"
    sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'" & constSatamCodes(3) & "|') > 0 THEN          /* 사탐-한국지리 */"
    sStr = sStr & "             '" & constSatamCodes(3) & "'"
    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'54|') > 0 THEN     /* 과탐-지구과학1 */"
    sStr = sStr & "             '54'"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             '00'"
    sStr = sStr & "         END END SEL4,"
    sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'" & constSatamCodes(4) & "|') > 0 THEN          /* 사탐-세계지리 */"
    sStr = sStr & "             '" & constSatamCodes(4) & "'"
    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'55|') > 0 THEN     /* 과탐-물리2 */"
    sStr = sStr & "             '55'"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             '00'"
    sStr = sStr & "         END END SEL5,"
    sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'" & constSatamCodes(5) & "|') > 0 THEN          /* 사탐-생활과윤리 */"
    sStr = sStr & "             '" & constSatamCodes(5) & "'"
    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'56|') > 0 THEN     /* 과탐-화학2 */"
    sStr = sStr & "             '56'"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             '00'"
    sStr = sStr & "         END END SEL6,"
    sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'" & constSatamCodes(6) & "|') > 0 THEN          /* 사탐-윤리사상 */"
    sStr = sStr & "             '" & constSatamCodes(6) & "'"
    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'57|') > 0 THEN     /* 과탐-생명과학2 */"
    sStr = sStr & "             '57'"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             '00'"
    sStr = sStr & "         END END SEL7,"
    sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'" & constSatamCodes(7) & "|') > 0 THEN          /* 사탐-법과정치 */"
    sStr = sStr & "             '" & constSatamCodes(7) & "'"
    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'58|') > 0 THEN     /* 과탐-지구과학2 */"
    sStr = sStr & "             '58'"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             '00'"
    sStr = sStr & "         END END SEL8,"
    sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'" & constSatamCodes(8) & "|') > 0 THEN          /* 사탐-경제 */"
    sStr = sStr & "             '" & constSatamCodes(8) & "'"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             '00'"
    sStr = sStr & "         END SEL9,"
    sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'" & constSatamCodes(9) & "|') > 0 THEN          /* 사탐-사회문화 */"
    sStr = sStr & "             '" & constSatamCodes(9) & "'"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             '00'"
    sStr = sStr & "         END SEL10,"
    sStr = sStr & " '' AS SEL11,"
    sStr = sStr & "  "
    sStr = sStr & "      /* 제2외국어 & 수리 */"
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
    sStr = sStr & "      /* 논술 */"
    sStr = sStr & "         CASE WHEN INSTR(SEL5,'91|') > 0 THEN         /* 언어 */"
    sStr = sStr & "             '91'"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             '00'"
    sStr = sStr & "         END SEL_N1,"
    sStr = sStr & "         CASE WHEN INSTR(SEL5,'92|') > 0 THEN         /* 수리 */"
    sStr = sStr & "             '92'"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             '00'"
    sStr = sStr & "         END SEL_N2,"
    sStr = sStr & "         CASE WHEN INSTR(SEL5,'93|') > 0 THEN         /* 외국어 */"      '< 변경
    sStr = sStr & "             '93'"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             '00'"
    sStr = sStr & "         END SEL_N3,"
    sStr = sStr & "         CASE WHEN INSTR(SEL5,'94|') > 0 THEN         /*  */"            '< 변경
    sStr = sStr & "             '94'"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             '00'"
    sStr = sStr & "         END SEL_N4, "
    sStr = sStr & "         GET_INTERNET_TOT_STD_INWON('" & Trim(basModule.SchCD) & "') AS PAYTOT, "        '< 전체집계 하는 함수
    sStr = sStr & "         K_NUM, M_NUM, E_NUM, TOT_NUM, "
    sStr = sStr & "         ZIP, ADDR1, ADDR2, TEL, CEL, "
    sStr = sStr & "         TO_CHAR(REGDATE,'YYYY-MM-DD HH24:MI:SS') AS REGDATE, "
    sStr = sStr & "         TO_CHAR(TIMESTAMP,'YYYY-MM-DD HH24:MI:SS') AS TIMESTAMP "
    
    sStr = sStr & "    FROM CLSTD91TB A, CLSTD92TB B"
    sStr = sStr & "   WHERE A.SCHNO = B.SCHNO "
    sStr = sStr & "     AND A.ACID  = '" & Trim(basModule.SchCD) & "'"
    
    '>> 유/무시험 체크
    If Trim(Right(cboExmType.Text, 30)) = "0" Then
        sStr = sStr & " AND A.EXMTYPE = '0'"
    ElseIf Trim(Right(cboExmType.Text, 30)) = "1" Then
        sStr = sStr & " AND A.EXMTYPE = '1'"
    End If
            
    '>> 인터넷/학원
    If Trim(Right(cboinGbn.Text, 30)) = "INT" Then          '< 인터넷 접수
        sStr = sStr & " AND A.R_WAY = '2'"
    ElseIf Trim(Right(cboinGbn.Text, 30)) = "HAK" Then      '< 학원 접수
        sStr = sStr & " AND A.R_WAY IN('1','3') "
    End If

    '>> 수험번호
    If Trim(fpExmID_F.UnFmtText) <> "" And Trim(fpExmID_E.UnFmtText) <> "" Then
        sStr = sStr & " AND A.EXMID BETWEEN '" & Trim(fpExmID_F.UnFmtText) & "'"
        sStr = sStr & "                 AND '" & Trim(fpExmID_E.UnFmtText) & "'"
    ElseIf Trim(fpExmID_F.UnFmtText) <> "" And Trim(fpExmID_E.UnFmtText) = "" Then
        sStr = sStr & " AND A.EXMID BETWEEN '" & Trim(fpExmID_F.UnFmtText) & "'"
        sStr = sStr & "                 AND '99999'"
    ElseIf Trim(fpExmID_F.UnFmtText) = "" And Trim(fpExmID_E.UnFmtText) <> "" Then
        sStr = sStr & " AND A.EXMID BETWEEN '00000'"
        sStr = sStr & "                 AND '" & Trim(fpExmID_E.UnFmtText) & "'"
    Else
        'no action
    End If
            
    If Trim(Right(cboKaeyol_F.Text, 30)) <> "ALL" Then      ' 인문
        sStr = sStr & " AND A.KAEYOL = '" & Trim(Right(cboKaeyol_F.Text, 30)) & "'"
    End If
    
    If Trim(txtStdNM_F.Text) <> "" Then
        sStr = sStr & " AND A.STDNM LIKE '%" & Trim(txtStdNM_F.Text) & "%'"
    End If
    
    If Trim(fpBirth_ymd_F.UnFmtText) <> "" Then
        sStr = sStr & " AND A.Birth_ymd LIKE '" & Trim(fpBirth_ymd_F.UnFmtText) & "%'"
    End If
            
    sStr = sStr & "     AND CL_CLOSE IS NULL "
    sStr = sStr & "     AND BIGO2 IS NULL"                  '< 2008.12. 수능본 학생은 년도가 들어가고 아니면 NULL
    
    sStr = sStr & "   ORDER BY SCHNO "
    
    
    Set DBCmd = New ADODB.Command
    Set DBRec = New ADODB.Recordset
    Set DBParam = New ADODB.Parameter
    
    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
    DBCmd.CommandText = sStr
    DBCmd.CommandType = adCmdText
    DBCmd.CommandTimeout = 30
    


    
    '>> 분원
'        sTmp = Trim(basModule.SchCD)
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
        
    '>> 수험번호
'        If Trim(fpExmID_F.UnFmtText) > "" Then
'            sTmp = Trim(fpExmID_F.UnFmtText)
'            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'                Set DBParam = DBCmd.CreateParameter("EXMID", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
'        End If
        
'    '>> 학생명
'        If Trim(txtStdNM_F.Text) > "" Then
'            sTmp = "%" & Trim(txtStdNM_F.Text) & "%"
'            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'                Set DBParam = DBCmd.CreateParameter("STDNM", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
'        End If
'
'    '>> 주민번호
'        If Trim(fpBirth_ymd_F.UnFmtText) > "" Then
'            sTmp = "%" & Trim(fpBirth_ymd_F.UnFmtText) & "%"
'            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'                Set DBParam = DBCmd.CreateParameter("Birth_ymd", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
'        End If
'
'    '>> 지망학원
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
    
    
    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic형태로 열게되면 record count를 할 수 없음.
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
                                sTmp = "노량진"
                            Case "K"
                                sTmp = "강남"
                            Case "S"
                                sTmp = "송파"
                            Case "P"
                                sTmp = "송파 M"
                            Case "M"
                                sTmp = "강남 M"
                            
                            Case "W"
                                sTmp = "주말법의대"
                            Case "Q"
                                sTmp = "야간법의대"
                                
                            Case "J"
                                sTmp = "양재"
                            Case "B"
                                sTmp = "부산"
                                
                        End Select
                    End If
                    Call basFunction.Set_SprType_Text(sprSTD_F, "CENTER", "LEFT", LenB(sTmp), sTmp)
                
                
                sprSTD_F.Col = 5
                    sTmp = " "
                    If IsNull(.Fields("SEL2_SCH")) = False Then
                        Select Case Trim(.Fields("SEL2_SCH"))
                            Case "N"
                                sTmp = "노량진"
                            Case "K"
                                sTmp = "강남"
                            Case "S"
                                sTmp = "송파"
                            Case "P"
                                sTmp = "송파 M"
                            Case "M"
                                sTmp = "강남 M"
                            
                            Case "W"
                                sTmp = "주말법의대"
                            Case "Q"
                                sTmp = "야간법의대"
                                
                            Case "J"
                                sTmp = "양재"
                            Case "B"
                                sTmp = "부산"
                                
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
                    nTmp = 0:   If IsNumeric(.Fields("K_NUM")) = True Then nTmp = Trim(.Fields("K_NUM"))
                    Call basFunction.Set_SprType_Numeric(sprSTD_F, 0, 0, 999999, "", nTmp)
                sprSTD_F.Col = 8
                    nTmp = 0:   If IsNumeric(.Fields("M_NUM")) = True Then nTmp = Trim(.Fields("M_NUM"))
                    Call basFunction.Set_SprType_Numeric(sprSTD_F, 0, 0, 999999, "", nTmp)
                sprSTD_F.Col = 9
                    nTmp = 0:   If IsNumeric(.Fields("E_NUM")) = True Then nTmp = Trim(.Fields("E_NUM"))
                    Call basFunction.Set_SprType_Numeric(sprSTD_F, 0, 0, 999999, "", nTmp)
                sprSTD_F.Col = 10
                    nTmp = 0:   If IsNumeric(.Fields("TOT_NUM")) = True Then nTmp = Trim(.Fields("TOT_NUM"))
                    Call basFunction.Set_SprType_Numeric(sprSTD_F, 0, 0, 999999, "", nTmp)
                
                sprSTD_F.Col = 11
                    sTmp = " ":  If IsNull(.Fields("GAEYUL")) = False Then sTmp = Trim(.Fields("GAEYUL"))
                    Call basFunction.Set_SprType_Text(sprSTD_F, "CENTER", "LEFT", LenB(sTmp), sTmp)
                
                sprSTD_F.SetCellBorder sprSTD_F.Col, sprSTD_F.Row, sprSTD_F.Col, sprSTD_F.Row, 2, basModule.SectionColor2, CellBorderStyleSolid
                
                '>> 선택과목 (사탐/ 과탐)
                For ni = 1 To SATAM_COUNT Step 1
                
                    If ni Mod 4 = 1 Then
                        sprSTD_F.SetCellBorder sprSTD_F.Col, sprSTD_F.Row, sprSTD_F.Col, sprSTD_F.Row, 2, basModule.SectionColor2, CellBorderStyleSolid
                    End If
                
                    sprSTD_F.Col = sprSTD_F.Col + 1
                    
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
                                    
                                    Case "51":   sTmp = "물1"
                                    Case "52":   sTmp = "화1"
                                    Case "53":   sTmp = "생1"
                                    Case "54":   sTmp = "지1"
                                    Case "55":   sTmp = "물2"
                                    Case "56":   sTmp = "화2"
                                    Case "57":   sTmp = "생2"
                                    Case "58":   sTmp = "지2"
                                    
                                End Select
                            End If
                            Call basFunction.Set_SprType_Text(sprSTD_F, "CENTER", "LEFT", LenB(sTmp), sTmp)
                        End If
                    End If
                Next ni
                
                '사탐과목하나 늘면서 빈칸으로 처리
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
                        
                            Case "31":  sTmp = "독어"
                            Case "32":  sTmp = "일어"
                            Case "33":  sTmp = "에스파냐어"
                            Case "34":  sTmp = "불어"
                            Case "35":  sTmp = "중국어"
                            Case "36":  sTmp = "한문"
                            
                            Case "81":  sTmp = "미적분"
                            Case "82":  sTmp = "이산수학"
                            Case "83":  sTmp = "확률통계"
                            Case "84":  sTmp = "수리나형"
                            
                        End Select
                        Call basFunction.Set_SprType_Text(sprSTD_F, "CENTER", "LEFT", LenB(sTmp), sTmp)
                    End If
                End If
                
                sprSTD_F.SetCellBorder sprSTD_F.Col, sprSTD_F.Row, sprSTD_F.Col, sprSTD_F.Row, 2, basModule.SectionColor1, CellBorderStyleSolid
                
            '>> 논술
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
                                    Case "91":  sTmp = "언어"
                                    Case "92":  sTmp = "수리"
                                    Case "93":  sTmp = "외국어"     '< 변경
                                    Case "94":  sTmp = ""           '< 변경
                                    
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
                    sTmp = " ":  If IsNull(.Fields("TIMESTAMP")) = False Then sTmp = Trim(.Fields("TIMESTAMP"))
                    Call basFunction.Set_SprType_Text(sprSTD_F, "CENTER", "LEFT", LenB(sTmp), sTmp)
                
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
    
    MsgBox "학생 조회하였습니다.", vbInformation + vbOKOnly, "학생조회"
    
    sprSTD_F.SetFocus
    
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    cmdFind.Enabled = True
    
    Exit Sub
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    cmdFind.Enabled = True
    
    MsgBox "학생조회시 에러가 발생하였습니다." & vbCrLf & _
           Trim(CStr(Err.Number)) & ":" & Err.Description, vbCritical + vbOKOnly, "학생조회"
           
    On Error GoTo 0
End Sub


'>> 학생선택
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














'## 전체학생 데이터 받기
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
    
    '> 초기화
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
            MsgBox "선택한 파일이 없습니다.", vbExclamation + vbOKOnly, Me.Caption
            Exit Sub
        End If
        
        sExcelFileName = .fileName
        
        ni = InStrRev(sExcelFileName, "\", -1, vbTextCompare)
        sExcelLogFile = Mid(sExcelFileName, 1, ni) & "\" & Mid(sExcelFileName, ni + 1, Len(sExcelFileName) - ni + 1 - 5)
        
    End With
    
    On Error GoTo 0
    
    On Error GoTo ErrStmt
   
    '## 헤더만들기
    sprStdData.MaxRows = sprStdData.MaxRows + 1
    sprStdData.Row = sprStdData.MaxRows
    
    For ni = 1 To sprSTD_F.MaxCols Step 1
        sprStdData.Col = ni
        
        sprSTD_F.Row = SpreadHeader
        sprSTD_F.Col = ni
        sTmp = " ":     If IsNull(sprSTD_F.Text) = False Then sTmp = Trim(sprSTD_F.Text)
            Call basFunction.Set_SprType_Text(sprStdData, "center", "left", basFunction.LenKor(sTmp), sTmp)
    Next ni
    
    For nRec = 1 To sprSTD_F.MaxRows Step 1
        sprStdData.MaxRows = sprStdData.MaxRows + 1
        sprStdData.Row = sprStdData.MaxRows
        
        For ni = 1 To sprSTD_F.MaxCols Step 1
            sprSTD_F.Row = nRec
            sprSTD_F.Col = ni
            sTmp = " ":     If IsNull(sprSTD_F.Text) = False Then sTmp = Trim(sprSTD_F.Text)
            
            sprStdData.Col = ni
                Call basFunction.Set_SprType_Text(sprStdData, "center", "left", basFunction.LenKor(sTmp), sTmp)
        Next ni
    Next nRec
    
    nRet = sprStdData.ExportToExcel(sExcelFileName, "Sheet1", sExcelLogFile)
    MsgBox "엑셀자료 작성완료하였습니다.", vbInformation + vbOKOnly, "전체학생 조회"
    
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    Exit Sub
    
ErrStmt1:
    MsgBox "저장할 엑셀명을 등록하세요.", vbExclamation + vbOKOnly, Me.Caption
    Exit Sub
    
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    MsgBox "전체학생 조회시 에러가 발생하였습니다." & vbCrLf & _
           Trim(CStr(Err.Number)) & ":" & Trim(Err.Description), vbCritical + vbOKOnly, "전체학생 조회"
    
    On Error GoTo 0
End Sub







