VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form MAT010 
   Caption         =   "입학사정 >> 입학원서 출력 >> 수학 집중 클리닉"
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
      BorderStyle     =   0  '없음
      Caption         =   "Frame2"
      Height          =   495
      Left            =   30
      TabIndex        =   39
      Top             =   0
      Width           =   14085
      Begin VB.Frame Frame1 
         BackColor       =   &H00D2EAF5&
         BorderStyle     =   0  '없음
         Caption         =   "Frame1"
         Height          =   435
         Left            =   30
         TabIndex        =   40
         Top             =   30
         Width           =   14025
         Begin VB.CommandButton cmdPrintAll 
            Caption         =   "전체페이지 출력"
            Height          =   375
            Left            =   10740
            TabIndex        =   48
            Top             =   30
            Width           =   1515
         End
         Begin VB.CommandButton cmdPrint 
            Caption         =   "현재페이지 출력"
            Height          =   375
            Left            =   9150
            TabIndex        =   47
            Top             =   30
            Width           =   1515
         End
         Begin VB.ComboBox cboKaeyol 
            Height          =   300
            Left            =   870
            Style           =   2  '드롭다운 목록
            TabIndex        =   46
            Top             =   67
            Width           =   1155
         End
         Begin VB.CommandButton cmdFind 
            Caption         =   "학생 조회"
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
            Caption         =   "◀"
            Height          =   375
            Left            =   12390
            TabIndex        =   42
            Top             =   30
            Width           =   405
         End
         Begin VB.CommandButton cmdShiftRight 
            Caption         =   "▶"
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
         Begin VB.Label NonPrintLbl 
            BackStyle       =   0  '투명
            Caption         =   "계열"
            BeginProperty Font 
               Name            =   "굴림체"
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
            BackStyle       =   0  '투명
            Caption         =   "수험번호          부터          까지"
            BeginProperty Font 
               Name            =   "굴림체"
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
            BackStyle       =   0  '투명
            Caption         =   "학생명"
            BeginProperty Font 
               Name            =   "굴림체"
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
      Begin VB.TextBox 수험번호 
         Alignment       =   1  '오른쪽 맞춤
         Appearance      =   0  '평면
         BorderStyle     =   0  '없음
         BeginProperty Font 
            Name            =   "돋움체"
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
         Appearance      =   0  '평면
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   9870
         Left            =   0
         ScaleHeight     =   9840
         ScaleWidth      =   13770
         TabIndex        =   1
         Top             =   0
         Width           =   13800
         Begin VB.TextBox 보호자주소1 
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
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
            Text            =   "보호자주소1"
            Top             =   6630
            Width           =   3135
         End
         Begin VB.TextBox 접수계열 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
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
            Text            =   "접수계열2"
            Top             =   2100
            Width           =   1515
         End
         Begin VB.TextBox 접수계열 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "돋움체"
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
            Text            =   "예.체능계"
            Top             =   330
            Width           =   1530
         End
         Begin VB.TextBox 생년월일 
            BorderStyle     =   0  '없음
            Height          =   225
            Left            =   7710
            TabIndex        =   58
            Text            =   "생년월일"
            Top             =   3570
            Width           =   2895
         End
         Begin VB.TextBox 학년1 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
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
            Text            =   "학년1"
            Top             =   7830
            Width           =   1545
         End
         Begin VB.TextBox 등급 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
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
            Text            =   "6월 평가원"
            Top             =   8790
            Width           =   1395
         End
         Begin VB.TextBox 학년 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림체"
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
            Text            =   "학년"
            Top             =   2970
            Width           =   660
         End
         Begin VB.TextBox 학생성명 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림체"
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
            Text            =   "홍길동"
            Top             =   3570
            Width           =   1545
         End
         Begin VB.TextBox 접수계열 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림체"
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
            Text            =   "계열"
            Top             =   2940
            Width           =   1860
         End
         Begin VB.TextBox 학생이메일 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
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
         Begin VB.TextBox 학생우편번호 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
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
         Begin VB.TextBox 보호자우편번호 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
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
         Begin VB.TextBox 보호자성명 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
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
            Text            =   "홍길동"
            Top             =   6150
            Width           =   1545
         End
         Begin VB.TextBox 보호자주소2 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
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
            Text            =   "서울 중구 신당동 떡복이집..................."
            Top             =   6870
            Width           =   3870
         End
         Begin VB.TextBox 보호자직업 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
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
            Text            =   "삼호물산주식회사"
            Top             =   6780
            Width           =   2325
         End
         Begin VB.TextBox 보호자연락처_휴대폰 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
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
         Begin VB.TextBox 보호자연락처_직장 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
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
         Begin VB.TextBox 학생주소1 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            Height          =   195
            Left            =   2460
            TabIndex        =   8
            Text            =   "서울 송파구 삼전동"
            Top             =   4350
            Width           =   3915
         End
         Begin VB.TextBox 학생출신고 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
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
            Text            =   "나사렛종고"
            Top             =   5490
            Width           =   3990
         End
         Begin VB.TextBox 학생연락처_집 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
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
         Begin VB.TextBox 학생연락처_휴대폰 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림"
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
         Begin VB.TextBox 학생주소2 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            Height          =   195
            Left            =   2460
            TabIndex        =   4
            Text            =   "53-21 쌍용빌라 나동 201호 "
            Top             =   4890
            Width           =   3960
         End
         Begin VB.TextBox 수리선택 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림체"
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
            Text            =   "수리선택"
            Top             =   8340
            Width           =   1080
         End
         Begin VB.TextBox 수리성적 
            Appearance      =   0  '평면
            BorderStyle     =   0  '없음
            BeginProperty Font 
               Name            =   "굴림체"
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
            Text            =   "수리성적"
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
            BackStyle       =   0  '투명
            Caption         =   "관   계"
            BeginProperty Font 
               Name            =   "굴림"
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
         Begin VB.Label 관계 
            BackStyle       =   0  '투명
            Caption         =   "관계"
            BeginProperty Font 
               Name            =   "굴림"
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
            BackStyle       =   0  '투명
            Caption         =   "*휴대폰"
            BeginProperty Font 
               Name            =   "굴림"
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
            BackStyle       =   0  '투명
            Caption         =   "휴대폰"
            BeginProperty Font 
               Name            =   "굴림"
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
            BackStyle       =   0  '투명
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "굴림"
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
            BackStyle       =   0  '투명
            Caption         =   "(학생의)"
            BeginProperty Font 
               Name            =   "굴림"
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
            Alignment       =   2  '가운데 맞춤
            BackStyle       =   0  '투명
            Caption         =   "(SMS 수신"
            BeginProperty Font 
               Name            =   "굴림"
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
            BackStyle       =   0  '투명
            Caption         =   "번호)"
            BeginProperty Font 
               Name            =   "굴림"
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
            BackStyle       =   0  '투명
            Caption         =   "@ 굵은선 안에만 기재하시오. (*는 필수정보이고 그 외에는 선택정보입니다.)"
            BeginProperty Font 
               Name            =   "굴림"
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
            BackStyle       =   1  '투명하지 않음
            BorderStyle     =   0  '투명
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
            BorderStyle     =   3  '점
            Index           =   8
            X1              =   810
            X2              =   5700
            Y1              =   7440
            Y2              =   7440
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            BorderStyle     =   3  '점
            Index           =   7
            X1              =   8010
            X2              =   12990
            Y1              =   7440
            Y2              =   7440
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '투명
            Caption         =   "점선 아랫부분은 기재하지 마시오"
            BeginProperty Font 
               Name            =   "굴림"
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
            BackStyle       =   0  '투명
            Caption         =   "2013년 학적카드"
            BeginProperty Font 
               Name            =   "굴림"
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
            BackStyle       =   0  '투명
            Caption         =   "학번:"
            BeginProperty Font 
               Name            =   "굴림"
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
         Begin VB.Label 성별 
            BackStyle       =   0  '투명
            Caption         =   "남"
            BeginProperty Font 
               Name            =   "굴림"
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
            BackStyle       =   0  '투명
            Caption         =   "성   별"
            BeginProperty Font 
               Name            =   "굴림"
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
            BackStyle       =   0  '투명
            Caption         =   "직   업"
            BeginProperty Font 
               Name            =   "굴림"
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
            BorderStyle     =   3  '점
            Index           =   6
            X1              =   60
            X2              =   13680
            Y1              =   1140
            Y2              =   1140
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '투명
            Caption         =   "수학 집중 클리닉 입학원서"
            BeginProperty Font 
               Name            =   "굴림"
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
            BackStyle       =   0  '투명
            Caption         =   "※수험번호"
            BeginProperty Font 
               Name            =   "굴림체"
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
            BackStyle       =   0  '투명
            Caption         =   "수리 선택"
            BeginProperty Font 
               Name            =   "굴림"
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
            BackStyle       =   0  '투명
            Caption         =   "  구   분"
            BeginProperty Font 
               Name            =   "굴림"
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
            BackStyle       =   0  '투명
            Caption         =   "계   열"
            BeginProperty Font 
               Name            =   "굴림"
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
            BackStyle       =   0  '투명
            Caption         =   "성   명"
            BeginProperty Font 
               Name            =   "굴림"
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
            BackStyle       =   0  '투명
            Caption         =   "생년월일"
            BeginProperty Font 
               Name            =   "굴림"
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
            BackStyle       =   0  '투명
            Caption         =   "학   년"
            BeginProperty Font 
               Name            =   "굴림"
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
            BackStyle       =   0  '투명
            Caption         =   "주   소"
            BeginProperty Font 
               Name            =   "굴림"
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
            BackStyle       =   0  '투명
            Caption         =   "전   화"
            BeginProperty Font 
               Name            =   "굴림"
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
            BackStyle       =   0  '투명
            Caption         =   "E-mail"
            BeginProperty Font 
               Name            =   "굴림"
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
            BackStyle       =   0  '투명
            Caption         =   "재학교"
            BeginProperty Font 
               Name            =   "굴림"
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
            BackStyle       =   0  '투명
            Caption         =   "(출신교)"
            BeginProperty Font 
               Name            =   "굴림"
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
            BackStyle       =   0  '투명
            Caption         =   "직장 전화"
            BeginProperty Font 
               Name            =   "굴림"
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
            BackStyle       =   0  '투명
            Caption         =   "성   명"
            BeginProperty Font 
               Name            =   "굴림"
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
            BackStyle       =   0  '투명
            Caption         =   "주   소"
            BeginProperty Font 
               Name            =   "굴림"
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
            BorderStyle     =   3  '점
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
            BackStyle       =   0  '투명
            Caption         =   "(근무처)"
            BeginProperty Font 
               Name            =   "굴림"
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
            BackStyle       =   0  '투명
            Caption         =   "호"
            BeginProperty Font 
               Name            =   "굴림"
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
            BackStyle       =   0  '투명
            Caption         =   "자"
            BeginProperty Font 
               Name            =   "굴림"
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
            BackStyle       =   0  '투명
            Caption         =   "보"
            BeginProperty Font 
               Name            =   "굴림"
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
            BackStyle       =   0  '투명
            Caption         =   "생"
            BeginProperty Font 
               Name            =   "굴림"
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
            BackStyle       =   0  '투명
            Caption         =   "학"
            BeginProperty Font 
               Name            =   "굴림"
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
            BackStyle       =   1  '투명하지 않음
            BorderStyle     =   0  '투명
            Height          =   660
            Index           =   4
            Left            =   10260
            Top             =   5880
            Width           =   2715
         End
         Begin VB.Shape FillBOXs 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  '투명하지 않음
            BorderStyle     =   0  '투명
            Height          =   4455
            Index           =   1
            Left            =   6540
            Top             =   2730
            Width           =   960
         End
         Begin VB.Shape FillBOXs 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  '투명하지 않음
            BorderStyle     =   0  '투명
            Height          =   675
            Index           =   0
            Left            =   4380
            Top             =   3360
            Width           =   885
         End
         Begin VB.Shape FillBOXs 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  '투명하지 않음
            BorderStyle     =   0  '투명
            Height          =   4440
            Index           =   2
            Left            =   1320
            Top             =   2730
            Width           =   960
         End
         Begin VB.Shape FillBOXs 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  '투명하지 않음
            BorderStyle     =   0  '투명
            Height          =   795
            Index           =   6
            Left            =   9420
            Top             =   240
            Width           =   1320
         End
         Begin VB.Shape Shape2 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  '투명하지 않음
            BorderColor     =   &H00000000&
            BorderStyle     =   0  '투명
            Height          =   1455
            Left            =   810
            Top             =   7680
            Width           =   1485
         End
         Begin VB.Shape FillBOXs 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  '투명하지 않음
            BorderStyle     =   0  '투명
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
'   시 스 템  명 : 대성학원 입학사정, 반배정 & 시간표 프로그램
'   서브시스템명 :
'   모   듈   명 : MAT010
'   모 듈  목 적 : 수학 집중 클리닉 입학원서
'
'   작   성   일 : 2009/11/26
'   작   성   자 : 유하균
' --------------------------------------------------------------------------------------------------------------
' --------------------------------------------------------------------------------------------------------------
'                 수     정     내     용
' --------------------------------------------------------------------------------------------------------------
'   1. 수정일 :
'   2. 내  용 :
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
    kaeyol      As String
    
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

Private sSavePath   As String       '<< image 경로
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
        
        '>> 계열
        With cboKaeyol
            .Clear
            .AddItem "인문" & Space(30) & "11"
            .AddItem "자연" & Space(30) & "12"
            .AddItem "전체" & Space(30) & "XX"
            .ListIndex = 2
        End With
        
        ReDim uSTD(0) As tSTD
        
    Me.Tag = ""
    
    성별.Caption = ""
    등급.Text = "2013 수능"
    
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
    'Width = 4890   ' 높이와 너비를 설정합니다.
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


'>> 학생 조회
Private Sub cmdFind_Click()
    
'    Select Case Trim(basModule.SchCD)
'        Case "N"
'
'        Case Else
'            MsgBox "노량진 대성학원이 아닌 경우 출력물이 다를 수 있습니다.", vbExclamation + vbOKOnly, "학생조회"
'    End Select
    
    On Error GoTo ErrStmt
    ReDim uSTD(0) As tSTD
    
    cmdFind.Enabled = False
        Call Get_STD_Data
        
    cmdFind.Enabled = True
    
    Exit Sub
ErrStmt:
    MsgBox "학생조회시 오류가 발생하였습니다.", vbCritical + vbOKOnly, "학생조회"
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
    
    
    '<< 초기 작업 : 제약조건
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
    
    
'>> 계열
    Select Case Trim(Right(cboKaeyol, 30))
        Case "XX"
            sStr = sStr & "     AND KEYOL IN ('11','12')"
        Case "11", "13"
            sStr = sStr & "     AND KEYOL = '11' "
        Case "12"
            sStr = sStr & "     AND KEYOL = '12' "
    End Select
    
'>> 수험번호
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
    
'>> 학생명
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
    
    
    
    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic형태로 열게되면 record count를 할 수 없음.
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
                        uSTD(nRec).EX_YN = "무시험"
                    ElseIf uSTD(nRec).EX_YN = "2" Then
                        uSTD(nRec).EX_YN = "유시험"
                    End If
                End If
                
                If IsNull(.Fields("KEYOL")) = False Then uSTD(nRec).kaeyol = .Fields("KEYOL")
                
                If IsNull(.Fields("SEL1")) = False Then uSTD(nRec).SEL1 = .Fields("SEL1")
                If IsNull(.Fields("SEL2")) = False Then uSTD(nRec).SEL2 = .Fields("SEL2")
                If IsNull(.Fields("SEL3")) = False Then uSTD(nRec).SEL3 = .Fields("SEL3")
                If IsNull(.Fields("SEL4")) = False Then uSTD(nRec).SEL4 = .Fields("SEL4")
                If IsNull(.Fields("SEL5")) = False Then uSTD(nRec).SEL5 = .Fields("SEL5")
                
                If IsNull(.Fields("GRADE_KOR")) = False Then uSTD(nRec).K_LEV = .Fields("GRADE_KOR")
                If IsNull(.Fields("GRADE_MAT")) = False Then uSTD(nRec).M_LEV = .Fields("GRADE_MAT")
                If IsNull(.Fields("GRADE_ENG")) = False Then uSTD(nRec).E_LEV = .Fields("GRADE_ENG")
                'If IsNull(.Fields("GTOT")) = False Then uSTD(nRec).TOT_NUM = .Fields("GTOT")
                
                '## 지망학원 - WINTER는 필요없음.
                If IsNull(.Fields("ACACD")) = False Then uSTD(nRec).ACID = .Fields("ACACD")
                    Select Case Trim(.Fields("ACACD"))
                        Case "N"
                            uSTD(nRec).SEL1_SCH = "노량진"
                        Case "K"
                            uSTD(nRec).SEL1_SCH = "강남"
                        Case "S"
                            uSTD(nRec).SEL1_SCH = "송파"
                        Case "P"
                            uSTD(nRec).SEL1_SCH = "송파 M"
                        Case "M"
                            uSTD(nRec).SEL1_SCH = "강남 M"
                            
                        Case "W"
                            uSTD(nRec).SEL1_SCH = "주말법의대"
                        Case "Q"
                            uSTD(nRec).SEL1_SCH = "야간법의대"
                            
                    End Select
                
                If IsNull(.Fields("ACACD")) = False Then uSTD(nRec).ACID = .Fields("ACACD")
                    Select Case Trim(.Fields("ACACD"))
                        Case "N"
                            uSTD(nRec).SEL2_SCH = "노량진"
                        Case "K"
                            uSTD(nRec).SEL2_SCH = "강남"
                        Case "S"
                            uSTD(nRec).SEL2_SCH = "송파"
                        Case "P"
                            uSTD(nRec).SEL2_SCH = "송파 M"
                        Case "M"
                            uSTD(nRec).SEL2_SCH = "강남 M"
                            
                        Case "W"
                            uSTD(nRec).SEL2_SCH = "주말법의대"
                        Case "Q"
                            uSTD(nRec).SEL2_SCH = "야간법의대"
                            
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
            
            Call Get_STD_image              '<< 이미지 자료 가져오기
            
            Call Std_Data_Show(1)           '<< 학생자료 화면 보이기
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
    MsgBox "학생조회시 에러가 발생하였습니다.", vbCritical + vbOKOnly, "학생조회"
End Sub


'## 서버의 이미지 가져오기
Private Sub Get_STD_image()
    
    Dim bData()     As Byte
    Dim f           As Integer
    Dim nRec        As Long

    Dim sLocalFile  As String
    Dim sSourceUrl  As String

    On Error Resume Next

    f = FreeFile()
    
    For nRec = 1 To UBound(uSTD) Step 1
    
        '2010.12.20 김한욱 노량진,송파,양재 경우 사진파일이 수험번호로 저장
        
        Select Case Trim(SchCD)
            Case "N"
                sLocalFile = sSavePath & "\" & uSTD(nRec).SU_NO & ".jpg"                    '수험번호
            Case "S"
                sLocalFile = sSavePath & "\" & uSTD(nRec).SU_NO & ".jpg"                    '수험번호
            Case "J"
                sLocalFile = sSavePath & "\" & uSTD(nRec).SU_NO & ".jpg"                    '수험번호
            Case Else
                sLocalFile = sSavePath & "\" & uSTD(nRec).ORD_NO & ".jpg"                   '<< unique key : ORD_NO
        End Select
        
        If Dir(sLocalFile, vbNormal) = "" Then                                                '<< 학생 이미지 없는 것만 받음
            If uSTD(nRec).PHOTO_PATH > " " Then
            
            
                Select Case Trim(basModule.SchCD)
                    Case "B"
                        sSourceUrl = "http://www.dsnschool.net" & uSTD(nRec).PHOTO_PATH        '<< 서버의 이미지 경로
                        
                    Case Else
                        sSourceUrl = "http://www.dshw.co.kr" & uSTD(nRec).PHOTO_PATH        '<< 서버의 이미지 경로
                        
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











'>> scroll 이동
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
            학년.Text = "재수생"
            학년1.Text = "재수생"
        Else
            학년.Text = .HAKYUN & "학년"
            학년1.Text = .HAKYUN & "학년"
        End If
        
        
        Select Case Trim(.kaeyol)   '<< 계열: 01,02,03-인문,자연,예체   06,05-수능인문,자연  06,07 -강남법대,의대
            Case "11"
                    접수계열(0).Text = "인 문 계"
                    접수계열(1).Text = "인 문 계"
                    접수계열(2).Text = "인 문 계"
            Case "12"
                    접수계열(0).Text = "자 연 계"
                    접수계열(1).Text = "자 연 계"
                    접수계열(2).Text = "자 연 계"
            Case Else
                    접수계열(0).Text = ""
                    접수계열(1).Text = ""
                    접수계열(2).Text = ""
        End Select
        
        
        If Trim(.PTS1) = "1" Then
            수리선택.Text = "가형"
        ElseIf Trim(.PTS1) = "2" Then
            수리선택.Text = "나형"
        End If
        
        '점수종류 3학년 이하면 원점수, 재수생이면 표준점수
        If CInt(.HAKYUN) < 4 Then
            수리성적.Text = "수리 원점수"   '1,2,3학년
        Else
            수리성적.Text = "수리 표준점수" '재수생
        End If
        
         If Trim(.MATJUM) = "0" Or Trim(.MATJUM) = "" Then
            수리성적.Text = "X"
        Else
            수리성적.Text = 수리성적.Text & " " & Trim(.MATJUM) & " 점"
        End If
        
        
        '등급
        Select Case Trim(.ETC1)
            Case "1"
                등급.Text = "2013 수능"
            Case "2"
                등급.Text = "6월 평가원"
            Case "3"
                등급.Text = "9월 평가원"
            Case "4"
                등급.Text = "고2 대성모의고사"
            Case "5"
                등급.Text = "고2 교육청모의고사"
            Case "9"
                등급.Text = "내신등급"
            Case Else
                등급.Text = "X"
        End Select
        
        
'        '2011년 수학 클리닉 점수 종류 구분 노량진 송파 양재
'        '김한욱 시작
'        Select Case Trim(SchCD)
'            Case "N"
'            Select Case Trim(.ETC1)
'                Case "1"
'                        M_Title(0).Text = "12학년도수능"        '2012.11.22일 고동석 수정 원래 3
'                Case "2"
'                        M_Title(0).Text = "6월 평가원"          '2012.11.22일 고동석 수정 원래 1
'                Case "3"
'                        M_Title(0).Text = "9월 평가원"          '2012.11.22일 고동석 수정 원래 2
'                Case Else
'                        M_Title(0).Text = "원점수"
'            End Select
'
'            Case "S"
'            Select Case Trim(.ETC1)
'                Case "1"
'                        M_Title(0).Text = "12학년도수능"
'                Case "2"
'                        M_Title(0).Text = "6월 평가원"
'                Case "3"
'                        M_Title(0).Text = "9월 평가원"
'                Case Else
'                        M_Title(0).Text = "원점수"
'            End Select
'
'            Case "J"
'            Select Case Trim(.ETC1)
'                Case "1"
'                        M_Title(0).Text = "12학년도수능"
'                Case "2"
'                        M_Title(0).Text = "6월 평가원"
'                Case "3"
'                        M_Title(0).Text = "9월 평가원"
'                Case Else
'                        M_Title(0).Text = "원점수"
'            End Select
'        '김한욱 끝
'        End Select
        
       
    
        
        수험번호.Text = .SU_NO
        학생성명.Text = .STDNM
        생년월일.Text = .Birth
        성별.Caption = IIf(.SEX = 1, "남", "여")
        학생우편번호.Text = "(" & .ZIP & ")"
        학생주소1.Text = .ADDR1
        학생주소2.Text = .ADDR2
        
        학생출신고.Text = .HIGH_SCH
        학생이메일.Text = .EMAIL
        학생연락처_집.Text = .TEL
        학생연락처_휴대폰.Text = .CEL
        
        보호자성명.Text = .PRNT_NM
        관계.Caption = IIf(.PRNT_RLTN = 1, "부", "모")
        보호자연락처_휴대폰.Text = .PRNT_CEL
        보호자우편번호.Text = "(" & .PRNT_ZIP & ")"
        보호자주소1.Text = .PRNT_ADDR1
        보호자주소2.Text = .PRNT_ADDR2
        
        
        보호자직업.Text = .PRNT_JOB
        보호자연락처_직장.Text = .PRNT_TEL
                       
        
        '<< 과목
        Call Div_Gwamok_NM("SEL1", .SEL1)
        Call Div_Gwamok_NM("SEL4", .SEL4)
        
        '<< 석차
        sTmp = ""
        sTmp = sTmp & "영어(" & Trim(.E_LEV) & "), "
        sTmp = sTmp & "수학(" & Trim(.M_LEV) & ")"
        '학교석차.Text = sTmp
        
        '제2지망.Text = .SEL2_SCH
        
        '2010.12.20 김한욱 노량진,송파,양재의 경우 수험번호로 된 사진 표시
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

'<< 과목넣기 : 배열로 되어있으니 유의할 것!!
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
                        sTmp = "국사"
                    Case "2"
                        sTmp = "윤리"
                    Case "3"
                        sTmp = "경제"
                    Case "4"
                        sTmp = "한국근현대"
                    Case "5"
                        sTmp = "세계사"
                    Case "6"
                        sTmp = "경제지리"
                    Case "7"
                        sTmp = "한국지리"
                    Case "8"
                        sTmp = "정치"
                    Case "9"
                        sTmp = "사회문화"
                    Case "10"
                        sTmp = "법과사회"
                    Case "11"
                        sTmp = "세계지리"
                End Select

                
            Case "SEL4"
            
                sTmp = ""
                Select Case Trim(sDiv(ni))
                    Case "1"
                        sTmp = "물리"
                    Case "2"
                        sTmp = "화학"
                    Case "3"
                        sTmp = "생명과학"
                    Case "4"
                        sTmp = "지구과학"
                End Select

                
            Case Else
                ' skip
                
                
        End Select
    Next ni
    
End Sub

'>> 이미지 받은파일 체크 : 체크시 이상이 있는 경우엔 default 값을 보여줌.
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
        MaxSize = LOF(f)                                        '<< 파일의 바이트 크기를 구합니다.
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




'## 전체 출력
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
        MsgBox "인쇄취소합니다.", vbExclamation + vbOKOnly, "현재페이지 인쇄하기"
        Exit Sub
    End If
    
    nRec = 0
    cmdPrint.Tag = "ALL"
    
    Do
        nRec = nRec + 1
        txtPage.Text = Trim(CStr(nRec)) & "/" & Trim(CStr(UBound(uSTD)))
        
        
        Call Std_Data_Show(nRec)                                '<< 학생자료 화면 보이기
        Me.Tag = "LOAD"
            VScroll1.value = nRec
            Call CmdPrint_Click:        DoEvents                '<< 1명 출력
            
        Me.Tag = ""

    Loop Until nRec = UBound(uSTD)
    
    cmdPrint.Tag = ""
    MsgBox "출력을 완료하였습니다.", vbInformation + vbOKOnly, "전체출력"
    
    Exit Sub
ErrStmt:
    On Error GoTo 0
    cmdPrint.Tag = ""
    
    MsgBox "출력시 에러가 발생하였습니다.", vbCritical + vbOKOnly, "전체출력"
    
End Sub

'## 현재 페이지 출력 : 1명 출력
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
    
    '<< 현재 페이지만 출력하면,
    If cmdPrint.Tag = "" Then
        bChk = False
        With dlgPrint
            .CancelError = True
            .ShowPrinter
            
            bChk = True
        End With
        
ErrPrint:
        If bChk = False Then
            MsgBox "인쇄취소합니다.", vbExclamation + vbOKOnly, "현재페이지 인쇄하기"
            Exit Sub
        End If
    End If
    
    '****************************************************************************************
    ' 프린터 출력초기화를 한다.
    ' PrintStartDoc (Width, Height, PaperSize, Orientation,TopMargin,LeftMargin
    '****************************************************************************************
    pRate = 1.15
    basFunction.PrintStartDoc pReportViewer.Width * pRate, pReportViewer.Height * pRate, vbPRPSA4, vbPRORLandscape, 1, 1


    '********************************************************************
    '  컬렉션을 이용하여 CONTROL을 배열로 처리한다.
    ' ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
    '  ※ 아래의 순서를 절대루 바꾸지 말것....   boss
    '********************************************************************
    Dim UsrCtl      As Control

    For Each UsrCtl In Me
        With UsrCtl
               If (UCase(TypeName(UsrCtl)) = "SHAPE" And UCase(UsrCtl.Name) = "FILLBOXS") Then
                '********************************************************************
                '  테두리 없는 사각 박스를 만들고 내부색을 칠한다.
                '********************************************************************
                 Printer.DrawWidth = 1                   ' 선의 굵기
                 Printer.FillStyle = vbFSTransparent     ' 단색
                 Printer.FillColor = &HC1F1FF            ' 색갈 칠하기
                 PrintFilledBox .Left * pRate, .Top * pRate, .Width * pRate, .Height * pRate, &HC1F1FF
             End If
        End With
    Next

    For Each UsrCtl In Me
        With UsrCtl
             If (UCase(TypeName(UsrCtl)) = "SHAPE" And UCase(UsrCtl.Name) = "BOXS") Then
                '********************************************************************
                '  line를 이용한 box만들기(기본적으로 shape는 출력시 line를 이용한다)
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
                         '  박스/line를 긋는다.
                         '********************************************************************
                          Printer.DrawStyle = IIf(UsrCtl.BorderStyle = 3, 2, UsrCtl.BorderStyle)
                          Printer.DrawWidth = IIf(UsrCtl.BorderStyle = 3, 1, UsrCtl.BorderWidth * 4)
                          Printer.FillStyle = vbFSTransparent
                          PrintLine .X1 * pRate, .Y1 * pRate, .X2 * pRate, .Y2 * pRate

                    Case "LABEL"
                          '********************************************************************
                          '  Label을 그대로 출력 한다(속성)
                          '  단) transparent는 true로 처리하고 실행한다.
                          '  SetBkMode(Printer.hdc, TRANSPARENT)문장은 MS버그를 처리하기 위함
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
                         '  데이터 출력 (DATA는 TEXTBOX로 처리 한다.)
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
                          '  사진출력
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

    Printer.EndDoc     ' 프린터로 보낸다

End Sub




















'## 사진 업로드
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

    If 학생성명.Text = "" Then
        MsgBox "학생을 조회하십시요.", vbExclamation + vbOKOnly, "사진 업로드"
        Exit Sub
    End If
    If UBound(uSTD) < 1 Then
        MsgBox "학생을 조회하십시요.", vbExclamation + vbOKOnly, "사진 업로드"
        Exit Sub
    End If
    
    '수험번호.tag
    
    With uSTD(VScroll1.value)
        sOrdNO = .ORD_NO
        sSchNO = .SCHNO
        sExmID = .EXMID
        
        
        simageFile = ""
        sPhotoPath = ""
        
        bRet = ""
        If Trim(sPhotoPath) = "" Then           '< 이미지가 없는 경우엔 강제로 생성
            bRet = Make_image_Path(sOrdNO, sExmID, simageFile)
            
            If bRet = "" Then
                MsgBox "경로 생성에 문제가 있습니다." & vbCrLf & _
                       "관리자에게 문의하십시요.", vbExclamation + vbOKOnly, "사진 업로드"
                Exit Sub
            Else
                sFileLocation = bRet
            End If
        End If
    End With
    
    '<< 파일 지우기 >>
    If Trim(txtPage) > " " Then
        sDiv = Split(txtPage.Text, "/", -1, vbTextCompare)
        
        nS = CLng(sDiv(0))
        sLocalFile = sSavePath & "\" & uSTD(nS).ORD_NO & ".jpg"       '<< unique key : ord_no
        If Dir(sLocalFile) > " " Then
            Kill sLocalFile
        End If
    End If
    
    '파일 넣기
    Load INT900
    Call INT900.Save_Photo(sFileLocation, sSchNO)
    INT900.Show
    
End Sub


'## 이미지 없는 경우 강제를 생성
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





