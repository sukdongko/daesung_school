VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form STD040 
   Caption         =   "입학사정 >> 합격생 ▶ 시간표 작업으로 등록"
   ClientHeight    =   9660
   ClientLeft      =   1710
   ClientTop       =   3585
   ClientWidth     =   16020
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9660
   ScaleWidth      =   16020
   Begin VB.Frame Frame3 
      BackColor       =   &H00C6AD84&
      BorderStyle     =   0  '없음
      Caption         =   "Frame3"
      Height          =   8925
      Left            =   60
      TabIndex        =   21
      Top             =   690
      Width           =   15465
      Begin VB.Frame Frame4 
         BackColor       =   &H00F7EFE7&
         BorderStyle     =   0  '없음
         Caption         =   "Frame4"
         Height          =   8865
         Left            =   30
         TabIndex        =   22
         Top             =   30
         Width           =   15405
         Begin EditLib.fpMask fpOK 
            Height          =   375
            Left            =   10980
            TabIndex        =   8
            Top             =   2460
            Width           =   795
            _Version        =   196608
            _ExtentX        =   1402
            _ExtentY        =   661
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
            Mask            =   "9999"
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
         Begin VB.CommandButton cmdFinish 
            Caption         =   "입학사정 완료하기"
            Height          =   600
            Left            =   7770
            TabIndex        =   7
            Top             =   2340
            Width           =   3075
         End
         Begin VB.CommandButton cmdOK_Cancel 
            Caption         =   "합격 취소하기"
            Height          =   600
            Left            =   7770
            TabIndex        =   10
            Top             =   6030
            Width           =   3075
         End
         Begin VB.CommandButton cmdDel 
            Caption         =   "학생 삭제하기"
            Height          =   600
            Left            =   7770
            TabIndex        =   9
            Top             =   4080
            Width           =   3075
         End
         Begin VB.CommandButton cmdOK 
            Caption         =   "합격생 ▶ 시간표 작업으로 등록"
            Height          =   600
            Left            =   7770
            TabIndex        =   6
            Top             =   570
            Width           =   3105
         End
         Begin VB.CheckBox chkDel 
            BackColor       =   &H00F7EFE7&
            Caption         =   "삭제"
            Height          =   225
            Left            =   6420
            TabIndex        =   12
            Top             =   150
            Width           =   915
         End
         Begin VB.CheckBox chkOK 
            BackColor       =   &H00F7EFE7&
            Caption         =   "선택"
            Height          =   225
            Left            =   660
            TabIndex        =   11
            Top             =   150
            Width           =   915
         End
         Begin FPSpread.vaSpread sprData 
            Height          =   8655
            Left            =   120
            TabIndex        =   13
            Top             =   120
            Width           =   7485
            _Version        =   393216
            _ExtentX        =   13203
            _ExtentY        =   15266
            _StockProps     =   64
            AllowDragDrop   =   -1  'True
            AllowMultiBlocks=   -1  'True
            AllowUserFormulas=   -1  'True
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
            Protect         =   0   'False
            SpreadDesigner  =   "STD040.frx":0000
         End
         Begin VB.Label Label4 
            BackStyle       =   0  '투명
            Caption         =   $"STD040.frx":19A6
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   6165
            Left            =   7980
            TabIndex        =   23
            Top             =   1230
            Width           =   6945
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0082C8E8&
      BorderStyle     =   0  '없음
      Caption         =   "Frame1"
      Height          =   615
      Left            =   60
      TabIndex        =   14
      Top             =   30
      Width           =   15465
      Begin VB.Frame Frame2 
         BackColor       =   &H00D2EAF5&
         BorderStyle     =   0  '없음
         Caption         =   "Frame2"
         Height          =   555
         Left            =   30
         TabIndex        =   15
         Top             =   30
         Width           =   15405
         Begin VB.CommandButton cmdFind 
            Caption         =   "조회하기 (&F)"
            Height          =   450
            Left            =   480
            TabIndex        =   0
            Top             =   60
            Width           =   1365
         End
         Begin VB.TextBox txtStdNM 
            Height          =   345
            Left            =   8400
            TabIndex        =   4
            Text            =   "txtStdNM"
            Top             =   90
            Width           =   1605
         End
         Begin VB.ComboBox cboKaeyol 
            Height          =   300
            Left            =   2850
            Style           =   2  '드롭다운 목록
            TabIndex        =   1
            Top             =   105
            Width           =   1275
         End
         Begin EditLib.fpMask fpBirth_ymd 
            Height          =   345
            Left            =   10890
            TabIndex        =   5
            Top             =   90
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
         Begin EditLib.fpMask fpExmID_S 
            Height          =   345
            Left            =   4950
            TabIndex        =   2
            Top             =   90
            Width           =   795
            _Version        =   196608
            _ExtentX        =   1402
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
         Begin EditLib.fpMask fpExmID_E 
            Height          =   345
            Left            =   6240
            TabIndex        =   3
            Top             =   90
            Width           =   825
            _Version        =   196608
            _ExtentX        =   1455
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
         Begin EditLib.fpLongInteger fpTotCnt 
            Height          =   345
            Left            =   14220
            TabIndex        =   24
            Top             =   90
            Width           =   825
            _Version        =   196608
            _ExtentX        =   1455
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
         Begin VB.Label Label5 
            BackStyle       =   0  '투명
            Caption         =   "조회인원"
            Height          =   210
            Left            =   13290
            TabIndex        =   25
            Top             =   150
            Width           =   975
         End
         Begin VB.Label Label1 
            BackStyle       =   0  '투명
            Caption         =   "수험번호               부터               까지"
            Height          =   210
            Left            =   4200
            TabIndex        =   20
            Top             =   150
            Width           =   3405
         End
         Begin VB.Label Label3 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "생년월일"
            Height          =   210
            Left            =   9870
            TabIndex        =   19
            Top             =   150
            Width           =   975
         End
         Begin VB.Label Label2 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "학생명"
            Height          =   210
            Left            =   7380
            TabIndex        =   18
            Top             =   150
            Width           =   975
         End
         Begin VB.Label Label28 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "계  열"
            Height          =   210
            Left            =   1860
            TabIndex        =   17
            Top             =   150
            Width           =   975
         End
         Begin VB.Label Label24 
            BackStyle       =   0  '투명
            Caption         =   ">> 조회항목"
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
            TabIndex        =   16
            Top             =   150
            Width           =   2625
         End
      End
   End
End
Attribute VB_Name = "STD040"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################################################
'   시 스 템  명 : 대성학원 입학사정, 반배정 & 시간표 프로그램
'   서브시스템명 :
'   모   듈   명 : STD040
'   모 듈  목 적 : 합격생 ▶ 시간표 작업으로 등록
'
'   작   성   일 : 2007/08/29
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
    Dim sSort       As String
    
    Dim sData       As String * 255
    Dim sGbn        As String
    Dim nRtn        As Long
    
    Me.Move 0, 0, 15700, 9980
    
    Me.Tag = "LOAD"
        With sprData
            .ShadowColor = basModule.ShadowColor1
            .ShadowDark = basModule.ShadowDark1
            .ShadowText = basModule.ShadowText1
            .GridColor = basModule.GridColor1
            .GrayAreaBackColor = basModule.GrayAreaBackColor1
        End With
        
'>> 계열
        With cboKaeyol
            .Clear
            .AddItem "전체" & Space(30) & "XX"
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
            If Trim(basModule.SchCD) = "K" Or Trim(basModule.SchCD) = "W" Or Trim(basModule.SchCD) = "Q" Then           '< 강남 2008.03.24
                .AddItem "주말법대" & Space(30) & "04"
                .AddItem "주말의대" & Space(30) & "05"
                
                .AddItem "야간법대" & Space(30) & "06"
                .AddItem "야간의대" & Space(30) & "07"
            
                .AddItem "선착순인문" & Space(30) & "11"
                .AddItem "선착순자연" & Space(30) & "12"
                
                .AddItem "선착순인문16" & Space(30) & "16"
                .AddItem "선착순자연17" & Space(30) & "17"
            End If
        '<< 계열 >> : 2008.02.15
            If Trim(basModule.SchCD) = "S" Then             '< 송파
''                .AddItem "예체능" & Space(30) & "03"
''
''                .AddItem "인문수능" & Space(30) & "05"
''                .AddItem "자연수능" & Space(30) & "06"
''
''                .AddItem "신설인문" & Space(30) & "11"
''                .AddItem "신설자연" & Space(30) & "12"
                
                .AddItem "인문프리미엄" & Space(30) & "18"
                .AddItem "자연프리미엄" & Space(30) & "19"

            End If
        '<< 계열 >> : 2008.02.15
            If Trim(basModule.SchCD) = "P" Then             '< 마송
                .AddItem "특별인문" & Space(30) & "03"
                .AddItem "특별자연" & Space(30) & "04"
            End If
            
            
            If Trim(basModule.SchCD) = "J" Then             '< 양재
                .AddItem "신설인문" & Space(30) & "11"
                .AddItem "신설자연" & Space(30) & "12"
                
                .AddItem "인문프리미엄" & Space(30) & "18"
                .AddItem "자연프리미엄" & Space(30) & "19"
            End If
            
            
        '<< 계열 >> : 2009.01.09
            If Trim(basModule.SchCD) = "B" Then             '< 부산
                .AddItem "수학선행인문" & Space(30) & "05"
                .AddItem "수학선행자연" & Space(30) & "06"
                
                .AddItem "연.고대인문" & Space(30) & "07"
                .AddItem "연.고대자연" & Space(30) & "08"
                
                .AddItem "심화인문" & Space(30) & "09"
                .AddItem "심화자연" & Space(30) & "10"
            End If
            
            .ListIndex = 0
        End With
            
        sprData.Tag = "0"
            
        Call init_Form
        
    Me.Tag = ""
End Sub

Private Sub init_Form()
    
    fpExmID_S.Text = ""
    fpExmID_E.Text = ""
    
    txtStdNM.Text = ""
    fpBirth_ymd.Text = ""
    
    sprData.MaxRows = 0
    
    fpTotCnt.value = 0
    fpOK.Text = ""
    
End Sub

'>> 조회조건의 학생검색
Private Sub cmdFind_Click()
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    
    Dim nLength     As Long
    Dim sStr        As String
    Dim ni          As Integer
    Dim nRec        As Long
    Dim sGbn        As String
    
    Dim sTmp        As String
    Dim nTmp        As Long
    
    On Error GoTo ErrStmt
    
    fpTotCnt.value = 0
    fpOK.Text = ""
    
    chkOK.value = 0
    chkDel.value = 0
    sprData.MaxRows = 0
    
    sStr = ""
    sStr = sStr & "  SELECT 0 AS SEL,"
    sStr = sStr & "         SCHNO, STDNM, ACID, EXMID, SUBSTR(Birth_ymd,1,4) ||'-'||SUBSTR(Birth_ymd,5,2) ||'-'||SUBSTR(Birth_ymd,7,2) AS Birth_ymd,"
    sStr = sStr & "         0 AS DEL"
    sStr = sStr & "    From CLSTD01TB"
    sStr = sStr & "   WHERE CY_ACNT > ' ' "
    sStr = sStr & "     AND TOT_AMT > 0 "
    
    sStr = sStr & "     AND (PASS1 = ? OR "
    sStr = sStr & "          PASS2 = ? OR "
    sStr = sStr & "          PASS3 = ? OR "
    sStr = sStr & "          PASS4 = ? ) "
'>> 계열
'    Select Case Trim(Right(cboKaeyol, 30))
'        Case "XX"
'            ' no action
'        Case "01", "03"
'            sStr = sStr & "AND SEL1 > ' ' "
'        Case "02", "04"             '< 2008.02.15
'            sStr = sStr & "AND SEL3 > ' ' "
'    End Select
    Select Case Trim(Right(cboKaeyol.Text, 30))         '< 2008.02.15
        Case "XX"
            ' no action
        Case Else
            sStr = sStr & "AND KAEYOL = '" & Trim(Right(cboKaeyol.Text, 30))
    End Select
    
'>> 수험번호
    If Trim(fpExmID_S.UnFmtText) > " " And Trim(fpExmID_E.UnFmtText) > " " Then
        sStr = sStr & " AND EXMID BETWEEN ? AND ? "
    ElseIf Trim(fpExmID_S.UnFmtText) > " " And Trim(fpExmID_E.UnFmtText) = " " Then
        sStr = sStr & " AND EXMID BETWEEN ? AND '99999' "
    ElseIf Trim(fpExmID_S.UnFmtText) = " " And Trim(fpExmID_E.UnFmtText) > " " Then
        sStr = sStr & " AND EXMID BETWEEN '00000' AND ? "
    ElseIf Trim(fpExmID_S.UnFmtText) = " " And Trim(fpExmID_E.UnFmtText) = " " Then
        ' no action
    End If
'>> 학생명
    If Trim(txtStdNM.Text) > " " Then
        sStr = sStr & " AND STDNM LIKE ? "
    End If
'>> 주민번호
    If Trim(fpBirth_ymd.UnFmtText) > " " Then
        sStr = sStr & " AND Birth_ymd LIKE ? "
    End If
'>> 완료여부 : 저장되면 YYMM값이 들어감.
    sStr = sStr & " AND CL_CLOSE IS NULL "
    sStr = sStr & " AND BIGO2 IS NULL"                      '< 2008.12. 수능본 학생은 년도가 들어가고 아니면 NULL
    
    sStr = sStr & " ORDER BY ACID, EXMID "
    
    Set DBCmd = New ADODB.Command
    Set DBRec = New ADODB.Recordset
    Set DBParam = New ADODB.Parameter
    
    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
    DBCmd.CommandText = sStr
    DBCmd.CommandType = adCmdText
    DBCmd.CommandTimeout = 30
    
    '>> 분원
'        sTmp = Trim(basModule.SchCD)
'            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'                Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
        
        For ni = 1 To 4 Step 1
            sTmp = Trim(basModule.SchCD)
                sGbn = "PASS" & Trim(CStr(ni))
            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
                Set DBParam = DBCmd.CreateParameter(sGbn, adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
        Next ni

    '>> 수험번호
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
    '>> 학생명
        If Trim(txtStdNM.Text) > " " Then
            sTmp = "%" & Trim(txtStdNM.Text) & "%"
            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
                Set DBParam = DBCmd.CreateParameter("STDNM", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
        End If
    '>> 주민번호
        If Trim(fpBirth_ymd.UnFmtText) > " " Then
            sTmp = "%" & Trim(fpBirth_ymd.UnFmtText) & "%"
            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
                Set DBParam = DBCmd.CreateParameter("Birth_ymd", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
        End If
    
    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic형태로 열게되면 record count를 할 수 없음.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop
    
    With DBRec
        If .RecordCount > 0 Then
            .MoveFirst
            
            For nRec = 1 To .RecordCount Step 1
                sprData.MaxRows = sprData.MaxRows + 1
                sprData.Row = sprData.MaxRows
                
                fpTotCnt.value = sprData.Row
                
                sprData.Col = 1
                    nTmp = 0:   If IsNull(.Fields("SEL")) = False Then nTmp = CLng(.Fields("SEL"))
                        Call basFunction.Set_SprType_ChkBox(sprData):       sprData.value = nTmp
                sprData.Col = sprData.Col + 1
                    sTmp = " ": If IsNull(.Fields("SCHNO")) = False Then sTmp = Trim(.Fields("SCHNO"))
                        Call basFunction.Set_SprType_Text(sprData, "CENTER", "LEFT", LenB(sTmp), sTmp)
                sprData.Col = sprData.Col + 1
                    sTmp = " ": If IsNull(.Fields("STDNM")) = False Then sTmp = Trim(.Fields("STDNM"))
                        Call basFunction.Set_SprType_Text(sprData, "CENTER", "LEFT", LenB(sTmp), sTmp)
                sprData.Col = sprData.Col + 1
                    sTmp = " ": If IsNull(.Fields("ACID")) = False Then sTmp = Trim(.Fields("ACID"))
                        Call basFunction.Set_SprType_Text(sprData, "CENTER", "LEFT", LenB(sTmp), sTmp)
                sprData.Col = sprData.Col + 1
                    sTmp = " ": If IsNull(.Fields("EXMID")) = False Then sTmp = Trim(.Fields("EXMID"))
                        Call basFunction.Set_SprType_Text(sprData, "CENTER", "LEFT", LenB(sTmp), sTmp)
                sprData.Col = sprData.Col + 1
                    sTmp = " ": If IsNull(.Fields("Birth_ymd")) = False Then sTmp = Trim(.Fields("Birth_ymd"))
                        Call basFunction.Set_SprType_Text(sprData, "CENTER", "LEFT", LenB(sTmp), sTmp)
                sprData.Col = sprData.Col + 1
                    nTmp = 0:   If IsNull(.Fields("DEL")) = False Then nTmp = CLng(.Fields("DEL"))
                        Call basFunction.Set_SprType_ChkBox(sprData):       sprData.value = nTmp
                
                .MoveNext
            Next nRec
            
            sprData.Row = 1:       sprData.Row2 = sprData.MaxRows
            sprData.Col = 1:       sprData.Col2 = sprData.MaxCols
            sprData.BlockMode = True
                sprData.BackColor = basModule.BackColor1
                sprData.BackColorStyle = BackColorStyleUnderGrid
                
                sprData.Lock = True
                sprData.Protect = True
            sprData.BlockMode = False
            
        End If
    End With
    
    MsgBox "학생 조회하였습니다.", vbInformation + vbOKOnly, "합격생 ▶ 시간표 작업으로 등록"
    
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    Exit Sub
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    On Error GoTo 0
    MsgBox "합격처리 및 확인 조회시 에러가 발생하였습니다.", vbCritical + vbOKOnly, "합격생 ▶ 시간표 작업으로 등록"
End Sub

'>> 선택 ## multi 선택
Private Sub sprData_Click(ByVal Col As Long, ByVal Row As Long)
    Dim nRow        As Long
    
    If Row < 1 Then Exit Sub

    With sprData
        If .MaxRows < 1 Then Exit Sub

        sprData.Enabled = False
        
            Select Case Col
                Case 1 To 6
                    If .Tag = "0" Then
                        .Row = CLng(.Tag):  .Row2 = .Row
                        .Col = 1:           .Col2 = .MaxCols
                        .BlockMode = True
                            .BackColor = basModule.BackColor1
                            .BackColorStyle = BackColorStyleUnderGrid
                        .BlockMode = False
                        
                        .Row = CLng(.Tag)
                            .Col = 1
                                .value = 0
                                
'                        For nRow = 1 To .MaxRows Step 1
'                            .Row = nRow
'                            .Col = 1
'                                .Value = 0
'                        Next nRow
                        
                        .Row = Row:     .Row2 = .Row
                        .Col = 1:       .Col2 = .MaxCols
                        .BlockMode = True
                            .BackColor = basModule.SelectColor1
                            .BackColorStyle = BackColorStyleUnderGrid
                        .BlockMode = False
                        
                        .Col = 1:       .value = 1
                        
                        .Tag = Trim(CStr(Row))
                    ElseIf .Tag > "0" Then
                        .Row = Row
                        .Col = 1
                        If .value = 1 Then
                            .value = 0
                            
                            .Row = Row:     .Row2 = .Row
                            .Col = 1:       .Col2 = .MaxCols
                            .BlockMode = True
                            .BackColor = basModule.BackColor1
                            .BackColorStyle = BackColorStyleUnderGrid
                            .BlockMode = False
                            
                            .Tag = Trim(CStr(Row))
                        Else
                            .value = 1
                            
                            .Row = Row:     .Row2 = .Row
                            .Col = 1:       .Col2 = .MaxCols
                            .BlockMode = True
                            .BackColor = basModule.SelectColor1
                            .BackColorStyle = BackColorStyleUnderGrid
                            .BlockMode = False
                            
                            .Tag = Trim(CStr(Row))
                        End If
                    
                    End If
                    
                    For nRow = 1 To .MaxRows Step 1
                        .Row = nRow
                        .Col = .MaxCols
                        .value = 0
                    Next nRow
                Case Else
                    If .Tag = "0" Then
                        .Row = CLng(.Tag):  .Row2 = .Row
                        .Col = 1:           .Col2 = .MaxCols
                        .BlockMode = True
                            .BackColor = basModule.BackColor1
                            .BackColorStyle = BackColorStyleUnderGrid
                        .BlockMode = False
                        
                        .Row = CLng(.Tag)
                            .Col = .MaxCols
                                .value = 0
                        
'                        For nRow = 1 To .MaxRows Step 1
'                            .Row = nRow
'                            .Col = .MaxCols
'                                .Value = 0
'                        Next nRow
                        
                        .Row = Row:     .Row2 = .Row
                        .Col = 1:       .Col2 = .MaxCols
                        .BlockMode = True
                        .BackColor = basModule.SelectColor1
                        .BackColorStyle = BackColorStyleUnderGrid
                        .BlockMode = False
                        
                        .Col = .MaxCols:    .value = 1
                        
                        .Tag = Trim(CStr(Row))
                    ElseIf .Tag > "0" Then
                        .Row = Row
                        .Col = .MaxCols
                        If .value = 1 Then
                            .value = 0
                            
                            .Row = Row:     .Row2 = .Row
                            .Col = 1:       .Col2 = .MaxCols
                            .BlockMode = True
                            .BackColor = basModule.BackColor1
                            .BackColorStyle = BackColorStyleUnderGrid
                            .BlockMode = False
                            
                            .Tag = Trim(CStr(Row))
                        Else
                            .value = 1
                            
                            .Row = Row:     .Row2 = .Row
                            .Col = 1:       .Col2 = .MaxCols
                            .BlockMode = True
                            .BackColor = basModule.SelectColor1
                            .BackColorStyle = BackColorStyleUnderGrid
                            .BlockMode = False
                            
                            .Tag = Trim(CStr(Row))
                        End If
                    
                    End If
                    
                    For nRow = 1 To .MaxRows Step 1
                        .Row = nRow
                        .Col = 1
                        .value = 0
                    Next nRow
                    
            End Select
        
        sprData.Enabled = True

    End With
End Sub

Private Sub sprData_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim nS      As Long
    Dim nE      As Long
    
    Dim nRow    As Long
    
    With sprData
    
        If .MaxRows = 0 Then Exit Sub
        
        Select Case Shift
'            Case 0
'                Call sprData_Click(1, .ActiveRow)
                
            Case 1          '<< shift
                Select Case .ActiveCol
                    Case 1 To 6
                        If Button = vbLeftButton Then
                            If .Tag > "0" Then              '<< 1. 선택하고 2. shift를 눌러 멀티로 선택한 경우
                                nS = CLng(.Tag)
                                nE = .ActiveRow
                                
                                If nS > nE Then
                                    nS = .ActiveRow
                                    nE = CLng(.Tag)
                                End If
                                
                                .Row = nS:  .Row2 = nE
                                .Col = 1:   .Col2 = .MaxCols
                                .BlockMode = True
                                    .BackColor = basModule.SelectColor1
                                    .BackColorStyle = BackColorStyleUnderGrid
                                .BlockMode = False
                                
                                For nRow = nS To nE Step 1
                                    .Row = nRow
                                    .Col = 1
                                        .value = 1
                                Next nRow
                                
                                .Tag = "0"
                                
                                For nRow = 1 To .MaxRows Step 1
                                    .Row = nRow
                                    .Col = .MaxCols
                                    .value = 0
                                Next nRow
                                
                            End If
                        End If

                    Case Else
                        If Button = vbLeftButton Then
                            If .Tag > "0" Then              '<< 1. 선택하고 2. shift를 눌러 멀티로 선택한 경우
                                nS = CLng(.Tag)
                                nE = .ActiveRow
                                
                                If nS > nE Then
                                    nS = .ActiveRow
                                    nE = CLng(.Tag)
                                End If
                                
                                .Row = nS:  .Row2 = nE
                                .Col = 1:   .Col2 = .MaxCols
                                .BlockMode = True
                                    .BackColor = basModule.SelectColor1
                                    .BackColorStyle = BackColorStyleUnderGrid
                                .BlockMode = False
                                
                                For nRow = nS To nE Step 1
                                    .Row = nRow
                                    .Col = .MaxCols
                                        .value = 1
                                Next nRow
                                
                                .Tag = "0"
                                
                                For nRow = 1 To .MaxRows Step 1
                                    .Row = nRow
                                    .Col = 1
                                    .value = 0
                                Next nRow
                                
                            End If
                        End If
                End Select
            
        End Select
    
    End With
End Sub

'>> 전체선택
Private Sub chkDel_Click()
    Dim ni      As Long
    
    With sprData
        If .MaxRows = 0 Then Exit Sub
            
        If chkDel.value = 0 Then
            For ni = 1 To .MaxRows Step 1
                .Row = ni
                .Col = .MaxCols
                    .value = 0
            Next ni
            
            .Row = 1:   .Row2 = .MaxRows
            .Col = 1:   .MaxCols = .MaxCols
            .BlockMode = True
                .BackColor = basModule.BackColor1
                .BackColorStyle = BackColorStyleUnderGrid
            .BlockMode = False
        Else
            chkOK.value = 0
            
            For ni = 1 To .MaxRows Step 1
                .Row = ni
                .Col = .MaxCols
                    .value = 1
                .Col = 1
                    .value = 0
            Next ni
            
            .Row = 1:   .Row2 = .MaxRows
            .Col = 1:   .MaxCols = .MaxCols
            .BlockMode = True
                .BackColor = basModule.SelectColor1
                .BackColorStyle = BackColorStyleUnderGrid
            .BlockMode = False
        End If
        
    End With
End Sub

Private Sub chkOK_Click()
    Dim ni      As Long
    
    With sprData
        If .MaxRows = 0 Then Exit Sub
            
        If chkOK.value = 0 Then
            For ni = 1 To .MaxRows Step 1
                .Row = ni
                .Col = 1
                    .value = 0
            Next ni
            
            .Row = 1:   .Row2 = .MaxRows
            .Col = 1:   .MaxCols = .MaxCols
            .BlockMode = True
                .BackColor = basModule.BackColor1
                .BackColorStyle = BackColorStyleUnderGrid
            .BlockMode = False
        Else
            chkDel.value = 0
            
            For ni = 1 To .MaxRows Step 1
                .Row = ni
                .Col = 1
                    .value = 1
                .Col = .MaxCols
                    .value = 0
            Next ni
            
            .Row = 1:   .Row2 = .MaxRows
            .Col = 1:   .MaxCols = .MaxCols
            .BlockMode = True
                .BackColor = basModule.SelectColor1
                .BackColorStyle = BackColorStyleUnderGrid
            .BlockMode = False
        End If
        
    End With
End Sub




'>> 합격생 ▶ 시간표 작업으로 등록 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Private Sub cmdOK_Click()
'## procedure
    Dim bRet        As Boolean
    Dim ni          As Long
    
    Dim nCnt        As Long
    
    '>> 조건체크
    With sprData
        If .MaxRows = 0 Then Exit Sub
        
        For ni = 1 To .MaxRows Step 1
            .Row = ni
            .Col = 1                    '<< 선택
            If .value = 1 Then
                nCnt = nCnt + 1
            End If
        Next ni
        
        If nCnt = 0 Then
            MsgBox "선택 1개 이상하십시요.", vbExclamation + vbOKOnly, "합격생▶시간표작업 등록하기"
            Exit Sub
        End If
    End With
    
    On Error GoTo ErrStmt
    
    cmdDel.Enabled = False
        bRet = Save_STD_Schedule
        
    cmdDel.Enabled = True
    
    If bRet = True Then
        MsgBox "합격생▶시간표작업 등록하였습니다.", vbInformation + vbOKOnly, "합격생▶시간표작업 등록하기"
    Else
        MsgBox "합격생▶시간표작업 등록시 에러가 발생하였습니다.", vbCritical + vbOKOnly, "합격생▶시간표작업 등록하기"
    End If
    
    Exit Sub
ErrStmt:
    MsgBox "합격생▶시간표작업 등록시 오류가 발생하였습니다.", vbCritical + vbOKOnly, "합격생▶시간표작업 등록학기"
    On Error GoTo 0
    
End Sub

'>> 합격생▶시간표작업 등록하기
Private Function Save_STD_Schedule() As Boolean
    Dim bRet        As Boolean
    
    Dim DBCmd       As ADODB.Command
    Dim DBParam     As ADODB.Parameter
    
    Dim ni          As Long
    
    Dim nLength     As Byte
    Dim sTmp        As String
    Dim nTmp        As Double
    Dim nRec        As Long
    
    bRet = False
    
    On Error GoTo ErrStmt
    
    basDataBase.DBConn.BeginTrans
    
    Set DBCmd = New ADODB.Command
    Set DBParam = New ADODB.Parameter
    
    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
    
    For nRec = sprData.MaxRows To 1 Step -1
    
        sprData.Row = nRec
        sprData.Col = 1
    
        If sprData.value = 1 Then
        
        '>> 기존 파라미터가 남아 있으면 메모리에서 삭제함.
            For ni = 0 To DBCmd.Parameters.count - 1 Step 1
                DBCmd.Parameters.Delete (0)
            Next ni
                
        '>> 시스템코드
            sprData.Col = 2:    sTmp = Trim(sprData.Text)
            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
                Set DBParam = DBCmd.CreateParameter("V_SCHNO", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
                
        '>> 학원코드
            
            sTmp = basModule.SchCD
            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
                Set DBParam = DBCmd.CreateParameter("V_ACID", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
        
        '>> 데이터 등록
            DBCmd.CommandType = adCmdStoredProc
            DBCmd.CommandText = "PG_STD.PROC_STD_SAVE_SCHEDULE"
            DBCmd.CommandTimeout = 30
            
            DBCmd.Execute
        
            Do While basDataBase.DBConn.State And adStateExecuting
                DoEvents
            Loop
            
            With sprData
                .Row = nRec:    .Row2 = .Row
                .Col = 1:       .Col2 = .MaxCols
                .BlockMode = True
                    .BackColor = basModule.BackColor1
                    .BackColorStyle = BackColorStyleUnderGrid
                .BlockMode = False
                
                .Col = 1:       .value = 0
            End With
            
        End If
    Next nRec
    
    chkOK.value = 0
    
    Set DBCmd = Nothing
    Set DBParam = Nothing
    
    basDataBase.DBConn.CommitTrans
    
    Save_STD_Schedule = True
    Exit Function
    
ErrStmt:
    basDataBase.DBConn.RollbackTrans
    
    Set DBCmd = Nothing
    Set DBParam = Nothing
    
    Save_STD_Schedule = False
    
    On Error GoTo 0
End Function



'>> 입학사정 완료하기 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'   2007.12.21 수정
Private Sub cmdFinish_Click()
'## Update
    Dim sTmp        As String
    Dim ni          As Integer
    
    Dim bRet        As Boolean
    Dim nCnt        As Long
    Dim nRec        As Long
    
    On Error GoTo ErrStmt
    
    If Trim(fpOK.UnFmtText) = "" Then
        MsgBox "완료시점의 월을 입력하십시요." & vbCrLf & _
               "월에 2번 처리하시는 경우엔 3번째자리를 1로 넣어주십시요." & vbCrLf & _
               "예) 기본 0802   두번 0812", vbExclamation + vbOKOnly, "입학사정 완료하기"
        Exit Sub
    End If
    
    If MsgBox("학생 합격자 내용을 모두 완료하시겠습니까?" & vbCrLf & _
              "완료시엔 기존 등록된 학생은 더이상 조회하실수 없습니다.", vbQuestion + vbYesNo, "입학사정 완료하기") = vbNo Then
        Exit Sub
    End If
    
    bRet = False
    
'    '>> 조건체크
'    With sprData
'        If .MaxRows = 0 Then Exit Sub
'
'        For ni = 1 To .MaxRows Step 1
'            .Row = ni
'            .Col = 1
'            If .Value = 1 Then
'                nCnt = nCnt + 1
'            End If
'        Next ni
'
'        If nCnt = 0 Then
'            MsgBox "선택 1개 이상하십시요.", vbExclamation + vbOKOnly, "입학사정 완료하기"
'            Exit Sub
'        End If
'    End With
    
    cmdFinish.Enabled = False
        bRet = Finish_STD_Data
        
    cmdFinish.Enabled = True
    
'    If bRet = True Then
'        With sprData
'            For nRec = .MaxRows To 1 Step -1
'                .Row = nRec
'                .Col = 1
'                If .Value = 1 Then
'                    .DeleteRows .Row, 1
'                    .MaxRows = .MaxRows - 1
'                End If
'            Next nRec
'        End With
'
'        MsgBox "입학사정 완료하였습니다.", vbInformation + vbOKOnly, "입학사정 완료하기"
'    Else
'        MsgBox "입학사정 완료등록시 에러가 발생하였습니다.", vbCritical + vbOKOnly, "입학사정 완료하기"
'    End If
    
    
    If bRet = True Then
        MsgBox "입학사정 완료하였습니다.", vbInformation + vbOKOnly, "입학사정 완료하기"
    Else
        MsgBox "입학사정 완료등록시 에러가 발생하였습니다.", vbCritical + vbOKOnly, "입학사정 완료하기"
    End If

    Exit Sub
ErrStmt:
    MsgBox "입학사정 완료시 오류가 발생하였습니다.", vbCritical + vbOKOnly, "입학사정 완료하기"
    On Error GoTo 0
End Sub

Private Function Finish_STD_Data() As Boolean
    Dim DBCmd       As ADODB.Command
    
    Dim sStr        As String
    Dim nExe        As Long
    
    Dim ni          As Integer
    
    On Error GoTo ErrStmt
    
    basDataBase.DBConn.BeginTrans
    
    Set DBCmd = New ADODB.Command
    
    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
            

        
    sStr = ""
    sStr = sStr & "  Update CLSTD01TB"
    sStr = sStr & "     SET CL_CLOSE = '" & Trim(fpOK.UnFmtText) & "'"
    sStr = sStr & "   WHERE SCHNO IN (SELECT SCHNO"
    sStr = sStr & "                     From CLSTD01TB"
    sStr = sStr & "                    WHERE ACID = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "                      AND (PASS1 = '" & Trim(basModule.SchCD) & "' OR"
    sStr = sStr & "                           PASS2 = '" & Trim(basModule.SchCD) & "' OR"
    sStr = sStr & "                           PASS3 = '" & Trim(basModule.SchCD) & "' OR"
    sStr = sStr & "                           PASS4 = '" & Trim(basModule.SchCD) & "')"
    sStr = sStr & "                      AND CY_ACNT > ' ' "
    sStr = sStr & "                      AND TOT_AMT > 0 "
    sStr = sStr & "                   Union"
    sStr = sStr & "                   SELECT SCHNO"
    sStr = sStr & "                     From CLSTD01TB"
    sStr = sStr & "                    WHERE (PASS1 = '" & Trim(basModule.SchCD) & "' OR"
    sStr = sStr & "                           PASS2 = '" & Trim(basModule.SchCD) & "' OR"
    sStr = sStr & "                           PASS3 = '" & Trim(basModule.SchCD) & "' OR"
    sStr = sStr & "                           PASS4 = '" & Trim(basModule.SchCD) & "')"
    sStr = sStr & "                      AND CY_ACNT > ' ' "
    sStr = sStr & "                      AND TOT_AMT > 0 "
    sStr = sStr & "                  )"
    sStr = sStr & "     AND CL_CLOSE IS NULL "
            
    DBCmd.CommandText = sStr
    DBCmd.CommandType = adCmdText
    DBCmd.CommandTimeout = 30
    
    DBCmd.Execute nExe, , -1
            
    Do While basDataBase.DBConn.State And adStateExecuting
        DoEvents
    Loop
    
    If MsgBox("전체 학생인원이 " & Trim(CStr(nExe)) & " 이 맞습니까?", vbQuestion + vbYesNo, "입학사정 완료하기") = vbYes Then
    
        sStr = ""
        sStr = sStr & "  Update CLSTD01TB"
        sStr = sStr & "     SET CL_CLOSE = '" & Trim(fpOK.UnFmtText) & "'"
        sStr = sStr & "   WHERE ACID = '" & Trim(basModule.SchCD) & "'"
        sStr = sStr & "     AND CL_CLOSE IS NULL "
                
        DBCmd.CommandText = sStr
        DBCmd.CommandType = adCmdText
        DBCmd.CommandTimeout = 30
        
        DBCmd.Execute nExe, , -1
        
        Do While basDataBase.DBConn.State And adStateExecuting
            DoEvents
        Loop
    
        If nExe > 0 Then
            Finish_STD_Data = True
            basDataBase.DBConn.CommitTrans
        Else
            Finish_STD_Data = False
            basDataBase.DBConn.RollbackTrans
        End If
    Else
        Finish_STD_Data = False
        basDataBase.DBConn.RollbackTrans
    End If
    
    Set DBCmd = Nothing
    Exit Function
    
ErrStmt:
    basDataBase.DBConn.RollbackTrans
    
    Set DBCmd = Nothing
    Finish_STD_Data = False
End Function


'Private Function Finish_STD_Data() As Boolean
'    Dim bRet        As Boolean
'
'    Dim DBCmd       As ADODB.Command
'    Dim DBParam     As ADODB.Parameter
'
'    Dim ni          As Long
'
'    Dim nLength     As Byte
'    Dim sTmp        As String
'    Dim nTmp        As Double
'
'    Dim nRow        As Long
'    Dim sStr        As String
'    Dim nEXE        As Integer
'
'    Dim nRec        As Long                                 '<< 처리해야 할 수
'    Dim nTot        As Long                                 '<< 처리한 수
'
'    bRet = False
'    nRec = 0
'    nTot = 0
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
'    For nRow = 1 To sprData.MaxRows Step 1
'
'        sprData.Row = nRow
'        sprData.Col = 1
'
'        If sprData.Value = 1 Then
'
'            nRec = nRec + 1
'

'
'            sStr = ""
'            sStr = sStr & "  Update CLSTD01TB"
'            sStr = sStr & "     SET CL_CLOSE = ? "
'            sStr = sStr & "   WHERE SCHNO    = ? "
'            sStr = sStr & "     AND ACID     = ? "
'
'            '>> 작업완료
'                sTmp = Format(Now, "YYMM")
'                nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'                    Set DBParam = DBCmd.CreateParameter("CY_ACNT", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
'            '>> 학생코드
'                sprData.Col = 2
'                sTmp = Trim(sprData.Text)
'                nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'                    Set DBParam = DBCmd.CreateParameter("SCHHO", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
'            '>> 학원코드 분류
'                sTmp = Trim(basModule.SchCD)
'                nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'                    Set DBParam = DBCmd.CreateParameter("ACID", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
'
'            DBCmd.CommandText = sStr
'            DBCmd.CommandType = adCmdText
'            DBCmd.CommandTimeout = 30
'
'            DBCmd.Execute nEXE, , -1
'
'            nTot = nTot + nEXE
'
'            Do While basDataBase.DBConn.State And adStateExecuting
'                DoEvents
'            Loop
'
'        End If
'    Next nRow
'
'    If nRec = nTot Then
'        Finish_STD_Data = True
'    Else
'        Finish_STD_Data = False
'    End If
'
'    Set DBCmd = Nothing
'    Set DBParam = Nothing
'
'    basDataBase.DBConn.CommitTrans
'    Exit Function
'
'ErrStmt:
'    basDataBase.DBConn.RollbackTrans
'
'    Set DBCmd = Nothing
'    Set DBParam = Nothing
'
'    Finish_STD_Data = False
'End Function




'>> 학생 삭제하기 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Private Sub cmdDel_Click()
'## procedure
    Dim bRet        As Boolean
    Dim ni          As Long
    
    Dim nCnt        As Long
    
    '>> 조건체크
    With sprData
        If .MaxRows = 0 Then Exit Sub
        
        For ni = 1 To .MaxRows Step 1
            .Row = ni
            .Col = .MaxCols
            If .value = 1 Then
                nCnt = nCnt + 1
            End If
        Next ni
        
        If nCnt = 0 Then
            MsgBox "선택 1개 이상하십시요.", vbExclamation + vbOKOnly, "학생 삭제하기"
            Exit Sub
        End If
    End With
    
    On Error GoTo ErrStmt
    
    
    cmdDel.Enabled = False
        bRet = Delete_StdOut
        
    cmdDel.Enabled = True
    
    If bRet = True Then
        MsgBox "학생 삭제하였습니다.", vbInformation + vbOKOnly, "학생 삭제하기"
    Else
        MsgBox "학생 삭제시 에러가 발생하였습니다.", vbCritical + vbOKOnly, "학생 삭제하기"
    End If
    
    Exit Sub
ErrStmt:
    MsgBox "학생삭제시 오류가 발생하였습니다.", vbCritical + vbOKOnly, "학생 삭제학기"
    On Error GoTo 0
    
End Sub

'>> 학생삭제
Private Function Delete_StdOut() As Boolean
    Dim bRet        As Boolean
    
    Dim DBCmd       As ADODB.Command
    Dim DBParam     As ADODB.Parameter
    
    Dim ni          As Long
    
    Dim nLength     As Byte
    Dim sTmp        As String
    Dim nTmp        As Double
    Dim nRec        As Long
    
    bRet = False
    
    On Error GoTo ErrStmt
    
    basDataBase.DBConn.BeginTrans
    
    Set DBCmd = New ADODB.Command
    Set DBParam = New ADODB.Parameter
    
    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
    
    
    For nRec = sprData.MaxRows To 1 Step -1
    
        sprData.Row = nRec
        sprData.Col = sprData.MaxCols
    
        If sprData.value = 1 Then
        
        '>> 기존 파라미터가 남아 있으면 메모리에서 삭제함.
            For ni = 0 To DBCmd.Parameters.count - 1 Step 1
                DBCmd.Parameters.Delete (0)
            Next ni
                
        '>> 시스템코드
            sprData.Col = 2:    sTmp = Trim(sprData.Text)
            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
                Set DBParam = DBCmd.CreateParameter("V_SCHNO", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
                
        '>> 학원코드
            sTmp = basModule.SchCD
            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
                Set DBParam = DBCmd.CreateParameter("V_ACID", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam

        '>> 데이터 등록
            DBCmd.CommandType = adCmdStoredProc
            DBCmd.CommandText = "PG_STD.PROC_STD_DELETE"
            DBCmd.CommandTimeout = 30
        
            DBCmd.Execute
        
            Do While basDataBase.DBConn.State And adStateExecuting
                DoEvents
            Loop
            
            sprData.DeleteRows sprData.Row, 1
            sprData.MaxRows = sprData.MaxRows - 1
        End If
    Next nRec
    
    Delete_StdOut = True
    
    Set DBCmd = Nothing
    Set DBParam = Nothing
    
    basDataBase.DBConn.CommitTrans
    Exit Function
    
ErrStmt:
    basDataBase.DBConn.RollbackTrans
    
    Set DBCmd = Nothing
    Set DBParam = Nothing
    
    Delete_StdOut = False
End Function

'>> 합격 취소하기 <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Private Sub cmdOK_Cancel_Click()
'## Update
Dim sTmp        As String
    Dim ni          As Integer
    
    Dim bRet        As Boolean
    Dim nCnt        As Long
    Dim nRec        As Long
    
    On Error GoTo ErrStmt
    
    bRet = False
    
    '>> 조건체크
    With sprData
        If .MaxRows = 0 Then Exit Sub
        
        For ni = 1 To .MaxRows Step 1
            .Row = ni
            .Col = .MaxCols
            If .value = 1 Then
                nCnt = nCnt + 1
            End If
        Next ni
        
        If nCnt = 0 Then
            MsgBox "선택 1개 이상하십시요.", vbExclamation + vbOKOnly, "합격 취소하기"
            Exit Sub
        End If
    End With
    
    cmdOK_Cancel.Enabled = False
        bRet = Cancel_STD_Data
        
    cmdOK_Cancel.Enabled = True
    
    If bRet = True Then
        With sprData
            For nRec = .MaxRows To 1 Step -1
                .Row = nRec
                .Col = .MaxCols
                If .value = 1 Then
                    .DeleteRows .Row, 1
                    .MaxRows = .MaxRows - 1
                End If
            Next nRec
        End With
        
        MsgBox "합격 취소하였습니다.", vbInformation + vbOKOnly, "합격 취소하기"
    Else
        MsgBox "합격 취소시 에러가 발생하였습니다.", vbCritical + vbOKOnly, "합격 취소하기"
    End If
    
    Exit Sub
ErrStmt:
    MsgBox "합격 취소시 오류가 발생하였습니다.", vbCritical + vbOKOnly, "합격 취소하기"
    On Error GoTo 0
End Sub

Private Function Cancel_STD_Data() As Boolean
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
    
    Dim nRec        As Long                                 '<< 처리해야 할 수
    Dim nTot        As Long                                 '<< 처리한 수
    
    Dim sSchNO      As String
    Dim sAcID       As String
    
    bRet = False
    nRec = 0
    nTot = 0
    
    On Error GoTo ErrStmt
    
    basDataBase.DBConn.BeginTrans
    
    Set DBCmd = New ADODB.Command
    Set DBParam = New ADODB.Parameter
    
    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
    
    For nRow = 1 To sprData.MaxRows Step 1
        
        sprData.Row = nRow
        sprData.Col = sprData.MaxCols
        
        If sprData.value = 1 Then
        
            nRec = nRec + 1
            
            '>> 기존 파라미터가 남아 있으면 메모리에서 삭제함.
            For ni = 0 To DBCmd.Parameters.count - 1 Step 1
                DBCmd.Parameters.Delete (0)
            Next ni
        
            sStr = ""
            sStr = sStr & " UPDATE CLSTD01TB "
            sStr = sStr & "    SET EXMID = '', "
            sStr = sStr & "        PASS1 = '', "
            sStr = sStr & "        PASS2 = '', "
            sStr = sStr & "        PASS3 = '', "
            sStr = sStr & "        PASS4 = '', "
            
            sStr = sStr & "        CY_ACNT = '', "
            sStr = sStr & "        TOT_AMT = 0 , "
            
            sStr = sStr & "        BASE_AMT1  = 0 , "
            sStr = sStr & "        BASE_AMT2  = 0 , "
            sStr = sStr & "        BASE_AMT3  = 0 , "
            sStr = sStr & "        BASE_AMT4  = 0 , "
            sStr = sStr & "        BASE_AMT5  = 0 , "
            sStr = sStr & "        BASE_AMT6  = 0 , "
            sStr = sStr & "        BASE_AMT7  = 0 , "
            sStr = sStr & "        BASE_AMT8  = 0 , "
            sStr = sStr & "        BASE_AMT9  = 0 , "
            sStr = sStr & "        BASE_AMT10 = 0 , "
            
            sStr = sStr & "        TAMGU_AMT1  = 0 , "
            sStr = sStr & "        TAMGU_AMT2  = 0 , "
            sStr = sStr & "        TAMGU_AMT3  = 0 , "
            sStr = sStr & "        TAMGU_AMT4  = 0 , "
            sStr = sStr & "        TAMGU_AMT5  = 0 , "
            sStr = sStr & "        TAMGU_AMT6  = 0 , "
            sStr = sStr & "        TAMGU_AMT7  = 0 , "
            sStr = sStr & "        TAMGU_AMT8  = 0 , "
            sStr = sStr & "        TAMGU_AMT9  = 0 , "
            sStr = sStr & "        TAMGU_AMT10 = 0 , "
            sStr = sStr & "        TAMGU_AMT11 = 0 , "
            sStr = sStr & "        TAMGU_AMT12 = 0  "
            
            sStr = sStr & "  WHERE SCHNO = ? "
            
            '>> 학생코드
                sprData.Col = 2
                sTmp = Trim(sprData.Text):      sSchNO = sTmp
                nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
                    Set DBParam = DBCmd.CreateParameter("SCHHO", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
'            '>> 학원코드 분류
'                sTmp = Trim(basModule.SchCD):   sAcID = sTmp
'                nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'                    Set DBParam = DBCmd.CreateParameter("ACID", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
            
            DBCmd.CommandText = sStr
            DBCmd.CommandType = adCmdText
            DBCmd.CommandTimeout = 30
            
            DBCmd.Execute nExe, , -1
            
            Do While basDataBase.DBConn.State And adStateExecuting
                DoEvents
            Loop
            
            nTot = nTot + nExe
            
            If nExe > 0 Then
                
                '## 시간표에 들어간 학생 삭제
                Call Cancel_STD_to_Delete_TTLtable_STD(sSchNO, sAcID)
                
            End If
            
        
        End If
    Next nRow
    
    If nRec = nTot Then
        Cancel_STD_Data = True
        basDataBase.DBConn.CommitTrans
    Else
        Cancel_STD_Data = False
        basDataBase.DBConn.RollbackTrans
    End If
    
    Set DBCmd = Nothing
    Set DBParam = Nothing
    
    
    Exit Function
    
ErrStmt:
    basDataBase.DBConn.RollbackTrans
    
    Set DBCmd = Nothing
    Set DBParam = Nothing
    
    Cancel_STD_Data = False
End Function


Private Sub Cancel_STD_to_Delete_TTLtable_STD(ByVal aSchNO As String, ByVal aAcID As String)
    
    Dim bRet        As Boolean
    
    Dim DBCmd       As ADODB.Command
    Dim DBParam     As ADODB.Parameter
    
    Dim ni          As Long
    
    Dim nLength     As Byte
    Dim sTmp        As String
    Dim nTmp        As Double
    Dim nRec        As Long
    
    Dim sStr        As String
    Dim nExe        As Integer
    
    On Error Resume Next
    
    Set DBCmd = New ADODB.Command
    Set DBParam = New ADODB.Parameter
    
    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
    
    sStr = ""
    sStr = sStr & "  DELETE "
    sStr = sStr & "    FROM CLTTL01TB "
    sStr = sStr & "   WHERE SCHNO    = ? "
    sStr = sStr & "     AND ACID     = ? "
            
    '>> 학생코드
        sTmp = aSchNO
        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
            Set DBParam = DBCmd.CreateParameter("SCHHO", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
    '>> 학원코드 분류
        sTmp = aAcID
        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
            Set DBParam = DBCmd.CreateParameter("ACID", adVarChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
            
    DBCmd.CommandText = sStr
    DBCmd.CommandType = adCmdText
    DBCmd.CommandTimeout = 30
    
    DBCmd.Execute nExe, , -1
            
    Do While basDataBase.DBConn.State And adStateExecuting
        DoEvents
    Loop
    
    Set DBCmd = Nothing
    Set DBParam = Nothing
    
    On Error GoTo 0
    
End Sub










