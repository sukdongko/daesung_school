VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form MAT_TEMP_FORM 
   Caption         =   "Form4"
   ClientHeight    =   12960
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15570
   LinkTopic       =   "Form4"
   ScaleHeight     =   12960
   ScaleWidth      =   15570
   StartUpPosition =   3  'Windows 기본값
   Begin VB.Frame 기본정보프레임 
      Caption         =   "Frame1"
      Height          =   7245
      Left            =   60
      TabIndex        =   8
      Top             =   2370
      Width           =   12585
      Begin VB.PictureBox pReportControl 
         Height          =   10095
         Left            =   0
         ScaleHeight     =   10035
         ScaleWidth      =   14040
         TabIndex        =   9
         Top             =   90
         Width           =   14100
         Begin VB.PictureBox pReportViewer 
            Appearance      =   0  '평면
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   15870
            Left            =   -600
            ScaleHeight     =   15840
            ScaleWidth      =   13770
            TabIndex        =   10
            Top             =   -1200
            Width           =   13800
            Begin VB.TextBox 생년월일 
               Height          =   465
               Left            =   7050
               TabIndex        =   29
               Text            =   "Text1"
               Top             =   3000
               Width           =   3165
            End
            Begin VB.TextBox 주민번호 
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
               Left            =   7725
               TabIndex        =   28
               Text            =   "690120 - 1473730"
               Top             =   3120
               Width           =   1620
            End
            Begin VB.TextBox 학생주소2 
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
               Left            =   2340
               TabIndex        =   27
               Text            =   "53-21 쌍용빌라 나동 201호 "
               Top             =   4560
               Width           =   4050
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
               Left            =   7755
               TabIndex        =   26
               Text            =   "011-9490-8607"
               Top             =   4635
               Width           =   2955
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
               Left            =   7755
               TabIndex        =   25
               Text            =   "02-2104-8600"
               Top             =   3825
               Width           =   2955
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
               Left            =   2445
               TabIndex        =   24
               Text            =   "나사렛종고"
               Top             =   5385
               Width           =   3990
            End
            Begin VB.TextBox 학생주소1 
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
               Left            =   2340
               TabIndex        =   23
               Text            =   "서울 송파구 삼전동"
               Top             =   3960
               Width           =   4095
            End
            Begin VB.TextBox 보호자주소2 
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
               Left            =   2340
               TabIndex        =   22
               Text            =   "서울 중구 신당동 떡복이집..................."
               Top             =   7905
               Width           =   4110
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
               Left            =   11010
               TabIndex        =   21
               Text            =   "02-2104-8600"
               Top             =   7485
               Width           =   1470
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
               Left            =   7815
               TabIndex        =   20
               Text            =   "011-9490-8607"
               Top             =   7485
               Width           =   1605
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
               Left            =   7755
               TabIndex        =   19
               Text            =   "삼호물산주식회사"
               Top             =   6315
               Width           =   2325
            End
            Begin VB.TextBox 보호자주소1 
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
               Left            =   2340
               TabIndex        =   18
               Text            =   "서울 중구 신당동 떡복이집..................."
               Top             =   7350
               Width           =   4140
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
               Left            =   2340
               TabIndex        =   17
               Text            =   "홍길동"
               Top             =   6315
               Width           =   1545
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
               Left            =   2355
               TabIndex        =   16
               Text            =   "(100-100)"
               Top             =   7140
               Width           =   1005
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
               Left            =   2340
               TabIndex        =   15
               Text            =   "(100-100)"
               Top             =   3645
               Width           =   1005
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
               Left            =   7755
               TabIndex        =   14
               Text            =   "iiiboss_12345@mail.naver.com"
               Top             =   5370
               Width           =   2955
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
               Left            =   8070
               TabIndex        =   13
               Text            =   "계열"
               Top             =   1860
               Width           =   1080
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
               Left            =   2430
               TabIndex        =   12
               Text            =   "홍길동"
               Top             =   2835
               Width           =   1545
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
               TabIndex        =   11
               Text            =   "학년"
               Top             =   1875
               Width           =   660
            End
            Begin VB.Shape FillBOXs 
               BackColor       =   &H00E0E0E0&
               BackStyle       =   1  '투명하지 않음
               BorderStyle     =   0  '투명
               Height          =   6780
               Index           =   2
               Left            =   1320
               Top             =   1590
               Width           =   960
            End
            Begin VB.Shape FillBOXs 
               BackColor       =   &H00E0E0E0&
               BackStyle       =   1  '투명하지 않음
               BorderStyle     =   0  '투명
               Height          =   765
               Index           =   0
               Left            =   6540
               Top             =   1590
               Width           =   960
            End
            Begin VB.Shape FillBOXs 
               BackColor       =   &H00E0E0E0&
               BackStyle       =   1  '투명하지 않음
               BorderStyle     =   0  '투명
               Height          =   570
               Index           =   5
               Left            =   6540
               Top             =   2340
               Width           =   4170
            End
            Begin VB.Shape FillBOXs 
               BackColor       =   &H00E0E0E0&
               BackStyle       =   1  '투명하지 않음
               BorderStyle     =   0  '투명
               Height          =   4875
               Index           =   1
               Left            =   6540
               Top             =   3510
               Width           =   960
            End
            Begin VB.Shape FillBOXs 
               BackColor       =   &H00E0E0E0&
               BackStyle       =   1  '투명하지 않음
               BorderStyle     =   0  '투명
               Height          =   930
               Index           =   4
               Left            =   10260
               Top             =   5940
               Width           =   2715
            End
            Begin VB.Label Label1 
               BackColor       =   &H80000005&
               Caption         =   "2012년"
               BeginProperty Font 
                  Name            =   "굴림"
                  Size            =   12
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   1
               Left            =   840
               TabIndex        =   51
               Top             =   1280
               Width           =   1695
            End
            Begin VB.Shape Boxs 
               BorderColor     =   &H00FF0000&
               BorderWidth     =   2
               Height          =   6795
               Index           =   2
               Left            =   750
               Top             =   1590
               Width           =   12195
            End
            Begin VB.Line Lines 
               BorderColor     =   &H00FF0000&
               BorderWidth     =   2
               Index           =   14
               X1              =   6540
               X2              =   6540
               Y1              =   1590
               Y2              =   8370
            End
            Begin VB.Line Lines 
               BorderColor     =   &H00FF0000&
               Index           =   11
               X1              =   1335
               X2              =   10710
               Y1              =   2340
               Y2              =   2340
            End
            Begin VB.Line Lines 
               BorderColor     =   &H00FF0000&
               Index           =   2
               X1              =   1335
               X2              =   10710
               Y1              =   3510
               Y2              =   3510
            End
            Begin VB.Line Lines 
               BorderColor     =   &H00FF0000&
               BorderWidth     =   2
               Index           =   1
               X1              =   780
               X2              =   12960
               Y1              =   5940
               Y2              =   5940
            End
            Begin VB.Line Lines 
               BorderColor     =   &H00FF0000&
               Index           =   50
               X1              =   1350
               X2              =   12990
               Y1              =   6870
               Y2              =   6870
            End
            Begin VB.Line Lines 
               BorderColor     =   &H00FF0000&
               Index           =   46
               X1              =   1320
               X2              =   1320
               Y1              =   1590
               Y2              =   8370
            End
            Begin VB.Line Lines 
               BorderColor     =   &H00FF0000&
               Index           =   45
               X1              =   7500
               X2              =   7500
               Y1              =   1575
               Y2              =   2340
            End
            Begin VB.Line Lines 
               BorderColor     =   &H00FF0000&
               Index           =   10
               X1              =   10710
               X2              =   10710
               Y1              =   1590
               Y2              =   4350
            End
            Begin VB.Line Lines 
               BorderColor     =   &H00FF0000&
               Index           =   44
               X1              =   2280
               X2              =   2280
               Y1              =   1590
               Y2              =   8370
            End
            Begin VB.Line Lines 
               BorderColor     =   &H00FF0000&
               Index           =   42
               X1              =   7500
               X2              =   7500
               Y1              =   3510
               Y2              =   8370
            End
            Begin VB.Line Lines 
               BorderColor     =   &H00FF0000&
               Index           =   37
               X1              =   6540
               X2              =   10725
               Y1              =   2910
               Y2              =   2910
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
               TabIndex        =   50
               Top             =   2775
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
               TabIndex        =   49
               Top             =   4080
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
               Left            =   945
               TabIndex        =   48
               Top             =   6615
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
               Left            =   945
               TabIndex        =   47
               Top             =   7485
               Width           =   195
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
               Left            =   945
               TabIndex        =   46
               Top             =   7050
               Width           =   195
            End
            Begin VB.Label Labels 
               BackStyle       =   0  '투명
               Caption         =   "휴대폰"
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
               Index           =   12
               Left            =   6705
               TabIndex        =   45
               Top             =   7485
               Width           =   645
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
               Left            =   6705
               TabIndex        =   44
               Top             =   6495
               Width           =   615
            End
            Begin VB.Line Lines 
               BorderColor     =   &H00FF0000&
               Index           =   0
               X1              =   1335
               X2              =   12945
               Y1              =   5025
               Y2              =   5025
            End
            Begin VB.Line Lines 
               BorderColor     =   &H00FF0000&
               BorderStyle     =   3  '점
               Index           =   13
               X1              =   2310
               X2              =   6510
               Y1              =   4320
               Y2              =   4320
            End
            Begin VB.Line Lines 
               BorderColor     =   &H00FF0000&
               Index           =   4
               X1              =   6525
               X2              =   12960
               Y1              =   4335
               Y2              =   4335
            End
            Begin VB.Line Lines 
               BorderColor     =   &H00FF0000&
               BorderStyle     =   3  '점
               Index           =   12
               X1              =   2280
               X2              =   6525
               Y1              =   7710
               Y2              =   7710
            End
            Begin VB.Label Labels 
               BackStyle       =   0  '투명
               Caption         =   "주    소"
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
               Index           =   17
               Left            =   1485
               TabIndex        =   43
               Top             =   7575
               Width           =   645
            End
            Begin VB.Label Labels 
               BackStyle       =   0  '투명
               Caption         =   "성    명"
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
               Index           =   18
               Left            =   1485
               TabIndex        =   42
               Top             =   6315
               Width           =   645
            End
            Begin VB.Label Labels 
               BackStyle       =   0  '투명
               Caption         =   "직    업"
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
               Index           =   19
               Left            =   6705
               TabIndex        =   41
               Top             =   6240
               Width           =   645
            End
            Begin VB.Line Lines 
               BorderColor     =   &H00FF0000&
               Index           =   16
               X1              =   10260
               X2              =   10260
               Y1              =   5925
               Y2              =   8370
            End
            Begin VB.Label Labels 
               BackStyle       =   0  '투명
               Caption         =   "직장 전화"
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
               Index           =   20
               Left            =   11190
               TabIndex        =   40
               Top             =   6315
               Width           =   900
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
               TabIndex        =   39
               Top             =   5475
               Width           =   615
            End
            Begin VB.Label Labels 
               BackStyle       =   0  '투명
               Caption         =   "재학교"
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
               Index           =   21
               Left            =   1485
               TabIndex        =   38
               Top             =   5220
               Width           =   645
            End
            Begin VB.Label Labels 
               BackStyle       =   0  '투명
               Caption         =   "E-mail"
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
               Index           =   13
               Left            =   6705
               TabIndex        =   37
               Top             =   5370
               Width           =   645
            End
            Begin VB.Label Labels 
               BackStyle       =   0  '투명
               Caption         =   "휴대폰"
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
               Index           =   14
               Left            =   6705
               TabIndex        =   36
               Top             =   4620
               Width           =   645
            End
            Begin VB.Label Labels 
               BackStyle       =   0  '투명
               Caption         =   "전    화"
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
               Index           =   15
               Left            =   6705
               TabIndex        =   35
               Top             =   3840
               Width           =   645
            End
            Begin VB.Label Labels 
               BackStyle       =   0  '투명
               Caption         =   "주    소"
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
               Index           =   22
               Left            =   1485
               TabIndex        =   34
               Top             =   4290
               Width           =   645
            End
            Begin VB.Label Labels 
               BackStyle       =   0  '투명
               Caption         =   "학    년"
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
               Index           =   25
               Left            =   1485
               TabIndex        =   33
               Top             =   1875
               Width           =   645
            End
            Begin VB.Label Labels 
               BackStyle       =   0  '투명
               Caption         =   "주민등록번호"
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
               Index           =   6
               Left            =   7935
               TabIndex        =   32
               Top             =   2520
               Width           =   1230
            End
            Begin VB.Label Labels 
               BackStyle       =   0  '투명
               Caption         =   "성    명"
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
               Index           =   16
               Left            =   1485
               TabIndex        =   31
               Top             =   2835
               Width           =   645
            End
            Begin VB.Label Labels 
               BackStyle       =   0  '투명
               Caption         =   "계    열"
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
               Index           =   2
               Left            =   6720
               TabIndex        =   30
               Top             =   1845
               Width           =   645
            End
            Begin VB.Image Photo 
               Height          =   2625
               Left            =   10830
               Picture         =   "TEMP_FORM.frx":0000
               Stretch         =   -1  'True
               Top             =   1650
               Width           =   2085
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
               Picture         =   "TEMP_FORM.frx":1406
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Frame3"
      Height          =   2295
      Left            =   30
      TabIndex        =   0
      Top             =   60
      Width           =   12975
      Begin VB.TextBox 접수계열2 
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
         Left            =   165
         TabIndex        =   1
         Text            =   "접수계열2"
         Top             =   1620
         Width           =   1515
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
         Index           =   323
         Left            =   30
         TabIndex        =   7
         Top             =   870
         Width           =   4065
      End
      Begin VB.Label Labels 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  '투명
         Caption         =   "2013년"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   322
         Left            =   0
         TabIndex        =   6
         Top             =   90
         Width           =   900
      End
      Begin VB.Label Labels 
         BackStyle       =   0  '투명
         Caption         =   "대성학원 입학원서"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   321
         Left            =   1120
         TabIndex        =   5
         Top             =   0
         Width           =   3585
      End
      Begin VB.Line Lines 
         BorderColor     =   &H00FF0000&
         Index           =   52
         X1              =   1890
         X2              =   4560
         Y1              =   2070
         Y2              =   2070
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
         Index           =   320
         Left            =   1950
         TabIndex        =   4
         Top             =   1710
         Width           =   760
      End
      Begin VB.Shape Boxs 
         BorderColor     =   &H00FF0000&
         Height          =   585
         Index           =   3
         Left            =   30
         Top             =   1500
         Width           =   1755
      End
      Begin VB.Label Label1 
         Alignment       =   1  '오른쪽 맞춤
         BackStyle       =   0  '투명
         Caption         =   "@ 굵은선 안에만 기재하시오. "
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   8.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   5940
         TabIndex        =   3
         Top             =   1920
         Width           =   3630
      End
      Begin VB.Label Label4 
         BackStyle       =   0  '투명
         Caption         =   "(*는 필수정보이고 그 외에는 선택정보입니다.)"
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
         Left            =   9540
         TabIndex        =   2
         Top             =   1920
         Width           =   3465
      End
   End
End
Attribute VB_Name = "MAT_TEMP_FORM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Frame3_DragDrop(Source As Control, X As Single, Y As Single)
End Sub

