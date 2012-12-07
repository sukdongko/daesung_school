VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form PRT050 
   Caption         =   "Ω√∞£«• √‚∑¬ >> ∫Û æÁΩƒ¡ˆ √‚∑¬"
   ClientHeight    =   11175
   ClientLeft      =   1395
   ClientTop       =   3390
   ClientWidth     =   16095
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   11175
   ScaleWidth      =   16095
   Begin VB.PictureBox pReportControl 
      BorderStyle     =   0  'æ¯¿Ω
      Height          =   9765
      Left            =   30
      ScaleHeight     =   9765
      ScaleWidth      =   14445
      TabIndex        =   40
      Top             =   540
      Width           =   14445
      Begin VB.VScrollBar VScroll1 
         Height          =   9765
         Left            =   14220
         Max             =   1
         TabIndex        =   368
         Top             =   0
         Width           =   225
      End
      Begin VB.PictureBox pReportViewer 
         BackColor       =   &H00FFFFFF&
         Height          =   9765
         Left            =   0
         ScaleHeight     =   9705
         ScaleWidth      =   14175
         TabIndex        =   41
         Top             =   0
         Width           =   14235
         Begin VB.TextBox MR 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'æ¯¿Ω
            Height          =   270
            Index           =   9
            Left            =   7920
            TabIndex        =   35
            Text            =   "MR"
            Top             =   9060
            Width           =   5895
         End
         Begin VB.TextBox MR 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'æ¯¿Ω
            Height          =   270
            Index           =   8
            Left            =   7920
            TabIndex        =   34
            Text            =   "MR"
            Top             =   8790
            Width           =   5895
         End
         Begin VB.TextBox MR 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'æ¯¿Ω
            Height          =   270
            Index           =   7
            Left            =   7920
            TabIndex        =   33
            Text            =   "MR"
            Top             =   8520
            Width           =   5895
         End
         Begin VB.TextBox ML 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'æ¯¿Ω
            Height          =   270
            Index           =   9
            Left            =   720
            TabIndex        =   24
            Text            =   "ML"
            Top             =   9060
            Width           =   5895
         End
         Begin VB.TextBox ML 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'æ¯¿Ω
            Height          =   270
            Index           =   8
            Left            =   720
            TabIndex        =   23
            Text            =   "ML"
            Top             =   8790
            Width           =   5895
         End
         Begin VB.TextBox ML 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'æ¯¿Ω
            Height          =   270
            Index           =   7
            Left            =   720
            TabIndex        =   22
            Text            =   "ML"
            Top             =   8520
            Width           =   5895
         End
         Begin VB.TextBox RTB 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'æ¯¿Ω
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   0
            Left            =   9810
            TabIndex        =   25
            Text            =   "RTB"
            Top             =   5160
            Width           =   3225
         End
         Begin VB.TextBox RTB 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'æ¯¿Ω
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   2
            Left            =   8700
            TabIndex        =   45
            Text            =   "RTB"
            Top             =   5250
            Width           =   615
         End
         Begin VB.TextBox RTB 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'æ¯¿Ω
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   1
            Left            =   8700
            TabIndex        =   44
            Text            =   "RTB"
            Top             =   5040
            Width           =   645
         End
         Begin VB.TextBox LTB 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'æ¯¿Ω
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   0
            Left            =   2580
            TabIndex        =   14
            Text            =   "LTB"
            Top             =   5160
            Width           =   3225
         End
         Begin VB.TextBox LTB 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'æ¯¿Ω
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   2
            Left            =   1500
            TabIndex        =   43
            Text            =   "LTB"
            Top             =   5250
            Width           =   615
         End
         Begin VB.TextBox LTB 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'æ¯¿Ω
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   1
            Left            =   1500
            TabIndex        =   42
            Text            =   "LTB"
            Top             =   5040
            Width           =   645
         End
         Begin VB.TextBox MR 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'æ¯¿Ω
            Height          =   270
            Index           =   6
            Left            =   7920
            TabIndex        =   32
            Text            =   "MR"
            Top             =   8250
            Width           =   5895
         End
         Begin VB.TextBox MR 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'æ¯¿Ω
            Height          =   270
            Index           =   5
            Left            =   7920
            TabIndex        =   31
            Text            =   "MR"
            Top             =   7980
            Width           =   5895
         End
         Begin VB.TextBox MR 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'æ¯¿Ω
            Height          =   270
            Index           =   4
            Left            =   7920
            TabIndex        =   30
            Text            =   "MR"
            Top             =   7710
            Width           =   5895
         End
         Begin VB.TextBox MR 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'æ¯¿Ω
            Height          =   270
            Index           =   3
            Left            =   7920
            TabIndex        =   29
            Text            =   "MR"
            Top             =   7440
            Width           =   5895
         End
         Begin VB.TextBox MR 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'æ¯¿Ω
            Height          =   270
            Index           =   2
            Left            =   7920
            TabIndex        =   28
            Text            =   "MR"
            Top             =   7170
            Width           =   5895
         End
         Begin VB.TextBox MR 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'æ¯¿Ω
            Height          =   270
            Index           =   1
            Left            =   7920
            TabIndex        =   27
            Text            =   "MR"
            Top             =   6900
            Width           =   5895
         End
         Begin VB.TextBox MR 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'æ¯¿Ω
            Height          =   270
            Index           =   0
            Left            =   7920
            TabIndex        =   26
            Text            =   "MR"
            Top             =   6630
            Width           =   5895
         End
         Begin VB.TextBox ML 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'æ¯¿Ω
            Height          =   270
            Index           =   6
            Left            =   720
            TabIndex        =   21
            Text            =   "ML"
            Top             =   8250
            Width           =   5895
         End
         Begin VB.TextBox ML 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'æ¯¿Ω
            Height          =   270
            Index           =   5
            Left            =   720
            TabIndex        =   20
            Text            =   "ML"
            Top             =   7980
            Width           =   5895
         End
         Begin VB.TextBox ML 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'æ¯¿Ω
            Height          =   270
            Index           =   4
            Left            =   720
            TabIndex        =   19
            Text            =   "ML"
            Top             =   7710
            Width           =   5895
         End
         Begin VB.TextBox ML 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'æ¯¿Ω
            Height          =   270
            Index           =   3
            Left            =   720
            TabIndex        =   18
            Text            =   "ML"
            Top             =   7440
            Width           =   5895
         End
         Begin VB.TextBox ML 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'æ¯¿Ω
            Height          =   270
            Index           =   2
            Left            =   720
            TabIndex        =   17
            Text            =   "ML"
            Top             =   7170
            Width           =   5895
         End
         Begin VB.TextBox ML 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'æ¯¿Ω
            Height          =   270
            Index           =   1
            Left            =   720
            TabIndex        =   16
            Text            =   "ML"
            Top             =   6900
            Width           =   5895
         End
         Begin VB.TextBox ML 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'æ¯¿Ω
            Height          =   270
            Index           =   0
            Left            =   720
            TabIndex        =   15
            Text            =   "ML"
            Top             =   6630
            Width           =   5895
         End
         Begin VB.Label RHD 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¥„¥Á : ¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   3
            Left            =   12330
            TabIndex        =   367
            Top             =   1020
            Width           =   1215
         End
         Begin VB.Label RHD 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "±≥Ω« : 100 »£"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   2
            Left            =   10950
            TabIndex        =   366
            Top             =   1020
            Width           =   1125
         End
         Begin VB.Label RHD 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "π› : ææÓøµø™π›"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   1
            Left            =   9330
            TabIndex        =   365
            Top             =   1020
            Width           =   1395
         End
         Begin VB.Label RHD 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "∞Ëø≠ : ¿ŒπÆ∞Ë"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   0
            Left            =   7950
            TabIndex        =   364
            Top             =   1020
            Width           =   1245
         End
         Begin VB.Label LHD 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¥„¥„ : ¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   3
            Left            =   5100
            TabIndex        =   363
            Top             =   1020
            Width           =   1215
         End
         Begin VB.Label LHD 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "±≥Ω« : 100 »£"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   2
            Left            =   3750
            TabIndex        =   362
            Top             =   1020
            Width           =   1125
         End
         Begin VB.Label LHD 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "π› : ææÓøµø™π›"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   1
            Left            =   2160
            TabIndex        =   361
            Top             =   1020
            Width           =   1395
         End
         Begin VB.Label LHD 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "∞Ëø≠ : ¿ŒπÆ∞Ë"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   0
            Left            =   750
            TabIndex        =   360
            Top             =   1020
            Width           =   1245
         End
         Begin VB.Label RC 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "08:00"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   21
            Left            =   8700
            TabIndex        =   359
            Top             =   2100
            Width           =   645
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            BorderStyle     =   3  '¡°
            Index           =   15
            X1              =   7260
            X2              =   7260
            Y1              =   90
            Y2              =   9660
         End
         Begin VB.Label LT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   102
            Left            =   2190
            TabIndex        =   358
            Top             =   6090
            Width           =   645
         End
         Begin VB.Label LS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   102
            Left            =   2190
            TabIndex        =   357
            Top             =   5880
            Width           =   645
         End
         Begin VB.Label LC 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "08:00"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   101
            Left            =   1500
            TabIndex        =   356
            Top             =   5880
            Width           =   645
         End
         Begin VB.Label LC 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "08:00"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   102
            Left            =   1500
            TabIndex        =   355
            Top             =   6090
            Width           =   645
         End
         Begin VB.Label LT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   92
            Left            =   2190
            TabIndex        =   354
            Top             =   5670
            Width           =   645
         End
         Begin VB.Label LS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   92
            Left            =   2190
            TabIndex        =   353
            Top             =   5460
            Width           =   645
         End
         Begin VB.Label LC 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "08:00"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   91
            Left            =   1500
            TabIndex        =   352
            Top             =   5460
            Width           =   645
         End
         Begin VB.Label LC 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "08:00"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   92
            Left            =   1500
            TabIndex        =   351
            Top             =   5670
            Width           =   645
         End
         Begin VB.Label LT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   82
            Left            =   2190
            TabIndex        =   350
            Top             =   4830
            Width           =   645
         End
         Begin VB.Label LS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   82
            Left            =   2190
            TabIndex        =   349
            Top             =   4620
            Width           =   645
         End
         Begin VB.Label LC 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "08:00"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   81
            Left            =   1500
            TabIndex        =   348
            Top             =   4620
            Width           =   645
         End
         Begin VB.Label LC 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "08:00"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   82
            Left            =   1500
            TabIndex        =   347
            Top             =   4830
            Width           =   645
         End
         Begin VB.Label LT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   72
            Left            =   2190
            TabIndex        =   346
            Top             =   4410
            Width           =   645
         End
         Begin VB.Label LS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   72
            Left            =   2190
            TabIndex        =   345
            Top             =   4200
            Width           =   645
         End
         Begin VB.Label LC 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "08:00"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   71
            Left            =   1500
            TabIndex        =   344
            Top             =   4200
            Width           =   645
         End
         Begin VB.Label LC 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "08:00"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   72
            Left            =   1500
            TabIndex        =   343
            Top             =   4410
            Width           =   645
         End
         Begin VB.Label LT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   62
            Left            =   2190
            TabIndex        =   342
            Top             =   3960
            Width           =   645
         End
         Begin VB.Label LS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   62
            Left            =   2190
            TabIndex        =   341
            Top             =   3750
            Width           =   645
         End
         Begin VB.Label LC 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "08:00"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   61
            Left            =   1500
            TabIndex        =   340
            Top             =   3750
            Width           =   645
         End
         Begin VB.Label LC 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "08:00"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   62
            Left            =   1500
            TabIndex        =   339
            Top             =   3960
            Width           =   645
         End
         Begin VB.Label LT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   52
            Left            =   2190
            TabIndex        =   338
            Top             =   3540
            Width           =   645
         End
         Begin VB.Label LS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   52
            Left            =   2190
            TabIndex        =   337
            Top             =   3330
            Width           =   645
         End
         Begin VB.Label LC 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "08:00"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   51
            Left            =   1500
            TabIndex        =   336
            Top             =   3330
            Width           =   645
         End
         Begin VB.Label LC 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "08:00"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   52
            Left            =   1500
            TabIndex        =   335
            Top             =   3540
            Width           =   645
         End
         Begin VB.Label LT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   42
            Left            =   2190
            TabIndex        =   334
            Top             =   3120
            Width           =   645
         End
         Begin VB.Label LS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   42
            Left            =   2190
            TabIndex        =   333
            Top             =   2910
            Width           =   645
         End
         Begin VB.Label LC 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "08:00"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   41
            Left            =   1500
            TabIndex        =   332
            Top             =   2940
            Width           =   645
         End
         Begin VB.Label LC 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "08:00"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   42
            Left            =   1500
            TabIndex        =   331
            Top             =   3150
            Width           =   645
         End
         Begin VB.Label LT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   32
            Left            =   2190
            TabIndex        =   330
            Top             =   2700
            Width           =   645
         End
         Begin VB.Label LS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   32
            Left            =   2190
            TabIndex        =   329
            Top             =   2490
            Width           =   645
         End
         Begin VB.Label LC 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "08:00"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   31
            Left            =   1500
            TabIndex        =   328
            Top             =   2520
            Width           =   645
         End
         Begin VB.Label LC 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "08:00"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   32
            Left            =   1500
            TabIndex        =   327
            Top             =   2730
            Width           =   645
         End
         Begin VB.Label LT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   22
            Left            =   2190
            TabIndex        =   326
            Top             =   2280
            Width           =   645
         End
         Begin VB.Label LS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   22
            Left            =   2190
            TabIndex        =   325
            Top             =   2070
            Width           =   645
         End
         Begin VB.Label LC 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "08:00"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   21
            Left            =   1500
            TabIndex        =   324
            Top             =   2100
            Width           =   645
         End
         Begin VB.Label LC 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "08:00"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   22
            Left            =   1500
            TabIndex        =   323
            Top             =   2310
            Width           =   645
         End
         Begin VB.Label LT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   12
            Left            =   2190
            TabIndex        =   322
            Top             =   1860
            Width           =   645
         End
         Begin VB.Shape Boxs 
            BorderColor     =   &H00FF0000&
            BorderWidth     =   2
            Height          =   5055
            Index           =   0
            Left            =   720
            Top             =   1260
            Width           =   5865
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            BorderWidth     =   2
            Index           =   28
            X1              =   720
            X2              =   6570
            Y1              =   1620
            Y2              =   1620
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "Ω√   ∞£   «•"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   20.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   60
            Left            =   2460
            TabIndex        =   321
            Top             =   300
            Width           =   2235
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            BorderWidth     =   2
            Index           =   0
            X1              =   2160
            X2              =   2160
            Y1              =   1260
            Y2              =   6300
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   3
            X1              =   5820
            X2              =   5820
            Y1              =   1260
            Y2              =   5010
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   4
            X1              =   1410
            X2              =   1410
            Y1              =   1620
            Y2              =   5010
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   1
            X1              =   5100
            X2              =   5100
            Y1              =   1260
            Y2              =   5010
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   2
            X1              =   4350
            X2              =   4350
            Y1              =   1260
            Y2              =   5010
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   5
            X1              =   3600
            X2              =   3600
            Y1              =   1260
            Y2              =   5010
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   6
            X1              =   2880
            X2              =   2880
            Y1              =   1260
            Y2              =   5010
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   7
            X1              =   720
            X2              =   6570
            Y1              =   2040
            Y2              =   2040
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   9
            X1              =   720
            X2              =   6570
            Y1              =   2460
            Y2              =   2460
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   10
            X1              =   720
            X2              =   6570
            Y1              =   2880
            Y2              =   2880
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   11
            X1              =   720
            X2              =   6570
            Y1              =   3300
            Y2              =   3300
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   14
            X1              =   720
            X2              =   6570
            Y1              =   3720
            Y2              =   3720
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   16
            X1              =   720
            X2              =   6570
            Y1              =   4140
            Y2              =   4140
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   17
            X1              =   720
            X2              =   6570
            Y1              =   4590
            Y2              =   4590
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   20
            X1              =   720
            X2              =   6570
            Y1              =   5010
            Y2              =   5010
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   21
            X1              =   720
            X2              =   6570
            Y1              =   5850
            Y2              =   5850
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   8
            X1              =   720
            X2              =   6570
            Y1              =   5430
            Y2              =   5430
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "ø˘"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   0
            Left            =   2400
            TabIndex        =   320
            Top             =   1350
            Width           =   315
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "»≠"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   1
            Left            =   3150
            TabIndex        =   319
            Top             =   1350
            Width           =   315
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "ºˆ"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   2
            Left            =   3870
            TabIndex        =   318
            Top             =   1350
            Width           =   315
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "∏Ò"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   3
            Left            =   4620
            TabIndex        =   317
            Top             =   1350
            Width           =   315
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "±›"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   4
            Left            =   5340
            TabIndex        =   316
            Top             =   1350
            Width           =   315
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "≈‰"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   5
            Left            =   6060
            TabIndex        =   315
            Top             =   1350
            Width           =   315
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "1±≥Ω√"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   6
            Left            =   780
            TabIndex        =   314
            Top             =   1770
            Width           =   585
         End
         Begin VB.Label LS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   12
            Left            =   2190
            TabIndex        =   313
            Top             =   1650
            Width           =   645
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "2±≥Ω√"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   8
            Left            =   780
            TabIndex        =   312
            Top             =   2190
            Width           =   585
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "3±≥Ω√"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   9
            Left            =   780
            TabIndex        =   311
            Top             =   2610
            Width           =   585
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "4±≥Ω√"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   10
            Left            =   780
            TabIndex        =   310
            Top             =   3000
            Width           =   585
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "5±≥Ω√"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   12
            Left            =   780
            TabIndex        =   309
            Top             =   3420
            Width           =   585
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "6±≥Ω√"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   13
            Left            =   780
            TabIndex        =   308
            Top             =   3870
            Width           =   585
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "7±≥Ω√"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   14
            Left            =   780
            TabIndex        =   307
            Top             =   4260
            Width           =   585
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "8±≥Ω√"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   15
            Left            =   780
            TabIndex        =   306
            Top             =   4680
            Width           =   585
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "9±≥Ω√"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   16
            Left            =   780
            TabIndex        =   305
            Top             =   5520
            Width           =   585
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "10±≥Ω√"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   17
            Left            =   750
            TabIndex        =   304
            Top             =   5940
            Width           =   705
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   22
            X1              =   1410
            X2              =   1410
            Y1              =   5430
            Y2              =   6300
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   23
            X1              =   2880
            X2              =   2880
            Y1              =   5430
            Y2              =   6300
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   24
            X1              =   3600
            X2              =   3600
            Y1              =   5430
            Y2              =   6300
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   25
            X1              =   4350
            X2              =   4350
            Y1              =   5430
            Y2              =   6300
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   26
            X1              =   5100
            X2              =   5100
            Y1              =   5430
            Y2              =   6300
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   27
            X1              =   5820
            X2              =   5820
            Y1              =   5430
            Y2              =   6270
         End
         Begin VB.Label LC 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "08:00"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   11
            Left            =   1500
            TabIndex        =   303
            Top             =   1680
            Width           =   645
         End
         Begin VB.Label LC 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "08:00"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   12
            Left            =   1500
            TabIndex        =   302
            Top             =   1890
            Width           =   645
         End
         Begin VB.Label LT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   103
            Left            =   2910
            TabIndex        =   301
            Top             =   6090
            Width           =   645
         End
         Begin VB.Label LS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   103
            Left            =   2910
            TabIndex        =   300
            Top             =   5880
            Width           =   645
         End
         Begin VB.Label LT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   93
            Left            =   2910
            TabIndex        =   299
            Top             =   5670
            Width           =   645
         End
         Begin VB.Label LS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   93
            Left            =   2910
            TabIndex        =   298
            Top             =   5460
            Width           =   645
         End
         Begin VB.Label LT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   83
            Left            =   2910
            TabIndex        =   297
            Top             =   4830
            Width           =   645
         End
         Begin VB.Label LS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   83
            Left            =   2910
            TabIndex        =   296
            Top             =   4620
            Width           =   645
         End
         Begin VB.Label LT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   73
            Left            =   2910
            TabIndex        =   295
            Top             =   4410
            Width           =   645
         End
         Begin VB.Label LS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   73
            Left            =   2910
            TabIndex        =   294
            Top             =   4200
            Width           =   645
         End
         Begin VB.Label LT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   63
            Left            =   2910
            TabIndex        =   293
            Top             =   3960
            Width           =   645
         End
         Begin VB.Label LS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   63
            Left            =   2910
            TabIndex        =   292
            Top             =   3750
            Width           =   645
         End
         Begin VB.Label LT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   53
            Left            =   2910
            TabIndex        =   291
            Top             =   3540
            Width           =   645
         End
         Begin VB.Label LS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   53
            Left            =   2910
            TabIndex        =   290
            Top             =   3330
            Width           =   645
         End
         Begin VB.Label LT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   43
            Left            =   2910
            TabIndex        =   289
            Top             =   3120
            Width           =   645
         End
         Begin VB.Label LS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   43
            Left            =   2910
            TabIndex        =   288
            Top             =   2910
            Width           =   645
         End
         Begin VB.Label LT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   33
            Left            =   2910
            TabIndex        =   287
            Top             =   2700
            Width           =   645
         End
         Begin VB.Label LS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   33
            Left            =   2910
            TabIndex        =   286
            Top             =   2490
            Width           =   645
         End
         Begin VB.Label LT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   23
            Left            =   2910
            TabIndex        =   285
            Top             =   2280
            Width           =   645
         End
         Begin VB.Label LS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   23
            Left            =   2910
            TabIndex        =   284
            Top             =   2070
            Width           =   645
         End
         Begin VB.Label LT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   13
            Left            =   2910
            TabIndex        =   283
            Top             =   1860
            Width           =   645
         End
         Begin VB.Label LS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   13
            Left            =   2910
            TabIndex        =   282
            Top             =   1650
            Width           =   645
         End
         Begin VB.Label LT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   104
            Left            =   3630
            TabIndex        =   281
            Top             =   6090
            Width           =   645
         End
         Begin VB.Label LS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   104
            Left            =   3630
            TabIndex        =   280
            Top             =   5880
            Width           =   645
         End
         Begin VB.Label LT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   94
            Left            =   3630
            TabIndex        =   279
            Top             =   5670
            Width           =   645
         End
         Begin VB.Label LS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   94
            Left            =   3630
            TabIndex        =   278
            Top             =   5460
            Width           =   645
         End
         Begin VB.Label LT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   84
            Left            =   3630
            TabIndex        =   277
            Top             =   4830
            Width           =   645
         End
         Begin VB.Label LS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   84
            Left            =   3630
            TabIndex        =   276
            Top             =   4620
            Width           =   645
         End
         Begin VB.Label LT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   74
            Left            =   3630
            TabIndex        =   275
            Top             =   4410
            Width           =   645
         End
         Begin VB.Label LS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   74
            Left            =   3630
            TabIndex        =   274
            Top             =   4200
            Width           =   645
         End
         Begin VB.Label LT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   64
            Left            =   3630
            TabIndex        =   273
            Top             =   3960
            Width           =   645
         End
         Begin VB.Label LS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   64
            Left            =   3630
            TabIndex        =   272
            Top             =   3750
            Width           =   645
         End
         Begin VB.Label LT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   54
            Left            =   3630
            TabIndex        =   271
            Top             =   3540
            Width           =   645
         End
         Begin VB.Label LS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   54
            Left            =   3630
            TabIndex        =   270
            Top             =   3330
            Width           =   645
         End
         Begin VB.Label LT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   44
            Left            =   3630
            TabIndex        =   269
            Top             =   3120
            Width           =   645
         End
         Begin VB.Label LS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   44
            Left            =   3630
            TabIndex        =   268
            Top             =   2910
            Width           =   645
         End
         Begin VB.Label LT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   34
            Left            =   3630
            TabIndex        =   267
            Top             =   2700
            Width           =   645
         End
         Begin VB.Label LS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   34
            Left            =   3630
            TabIndex        =   266
            Top             =   2490
            Width           =   645
         End
         Begin VB.Label LT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   24
            Left            =   3630
            TabIndex        =   265
            Top             =   2280
            Width           =   645
         End
         Begin VB.Label LS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   24
            Left            =   3630
            TabIndex        =   264
            Top             =   2070
            Width           =   645
         End
         Begin VB.Label LT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   14
            Left            =   3630
            TabIndex        =   263
            Top             =   1860
            Width           =   645
         End
         Begin VB.Label LS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   14
            Left            =   3630
            TabIndex        =   262
            Top             =   1650
            Width           =   645
         End
         Begin VB.Label LT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   105
            Left            =   4380
            TabIndex        =   261
            Top             =   6090
            Width           =   645
         End
         Begin VB.Label LS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   105
            Left            =   4380
            TabIndex        =   260
            Top             =   5880
            Width           =   645
         End
         Begin VB.Label LT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   95
            Left            =   4380
            TabIndex        =   259
            Top             =   5670
            Width           =   645
         End
         Begin VB.Label LS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   95
            Left            =   4380
            TabIndex        =   258
            Top             =   5460
            Width           =   645
         End
         Begin VB.Label LT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   85
            Left            =   4380
            TabIndex        =   257
            Top             =   4830
            Width           =   645
         End
         Begin VB.Label LS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   85
            Left            =   4380
            TabIndex        =   256
            Top             =   4620
            Width           =   645
         End
         Begin VB.Label LT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   75
            Left            =   4380
            TabIndex        =   255
            Top             =   4410
            Width           =   645
         End
         Begin VB.Label LS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   75
            Left            =   4380
            TabIndex        =   254
            Top             =   4200
            Width           =   645
         End
         Begin VB.Label LT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   65
            Left            =   4380
            TabIndex        =   253
            Top             =   3960
            Width           =   645
         End
         Begin VB.Label LS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   65
            Left            =   4380
            TabIndex        =   252
            Top             =   3750
            Width           =   645
         End
         Begin VB.Label LT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   55
            Left            =   4380
            TabIndex        =   251
            Top             =   3540
            Width           =   645
         End
         Begin VB.Label LS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   55
            Left            =   4380
            TabIndex        =   250
            Top             =   3330
            Width           =   645
         End
         Begin VB.Label LT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   45
            Left            =   4380
            TabIndex        =   249
            Top             =   3120
            Width           =   645
         End
         Begin VB.Label LS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   45
            Left            =   4380
            TabIndex        =   248
            Top             =   2910
            Width           =   645
         End
         Begin VB.Label LT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   35
            Left            =   4380
            TabIndex        =   247
            Top             =   2700
            Width           =   645
         End
         Begin VB.Label LS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   35
            Left            =   4380
            TabIndex        =   246
            Top             =   2490
            Width           =   645
         End
         Begin VB.Label LT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   25
            Left            =   4380
            TabIndex        =   245
            Top             =   2280
            Width           =   645
         End
         Begin VB.Label LS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   25
            Left            =   4380
            TabIndex        =   244
            Top             =   2070
            Width           =   645
         End
         Begin VB.Label LT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   15
            Left            =   4380
            TabIndex        =   243
            Top             =   1860
            Width           =   645
         End
         Begin VB.Label LS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   15
            Left            =   4380
            TabIndex        =   242
            Top             =   1650
            Width           =   645
         End
         Begin VB.Label LT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   106
            Left            =   5130
            TabIndex        =   241
            Top             =   6090
            Width           =   645
         End
         Begin VB.Label LS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   106
            Left            =   5130
            TabIndex        =   240
            Top             =   5880
            Width           =   645
         End
         Begin VB.Label LT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   96
            Left            =   5130
            TabIndex        =   239
            Top             =   5670
            Width           =   645
         End
         Begin VB.Label LS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   96
            Left            =   5130
            TabIndex        =   238
            Top             =   5460
            Width           =   645
         End
         Begin VB.Label LT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   86
            Left            =   5130
            TabIndex        =   237
            Top             =   4830
            Width           =   645
         End
         Begin VB.Label LS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   86
            Left            =   5130
            TabIndex        =   236
            Top             =   4620
            Width           =   645
         End
         Begin VB.Label LT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   76
            Left            =   5130
            TabIndex        =   235
            Top             =   4410
            Width           =   645
         End
         Begin VB.Label LS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   76
            Left            =   5130
            TabIndex        =   234
            Top             =   4200
            Width           =   645
         End
         Begin VB.Label LT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   66
            Left            =   5130
            TabIndex        =   233
            Top             =   3960
            Width           =   645
         End
         Begin VB.Label LS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   66
            Left            =   5130
            TabIndex        =   232
            Top             =   3750
            Width           =   645
         End
         Begin VB.Label LT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   56
            Left            =   5130
            TabIndex        =   231
            Top             =   3540
            Width           =   645
         End
         Begin VB.Label LS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   56
            Left            =   5130
            TabIndex        =   230
            Top             =   3330
            Width           =   645
         End
         Begin VB.Label LT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   46
            Left            =   5130
            TabIndex        =   229
            Top             =   3120
            Width           =   645
         End
         Begin VB.Label LS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   46
            Left            =   5130
            TabIndex        =   228
            Top             =   2910
            Width           =   645
         End
         Begin VB.Label LT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   36
            Left            =   5130
            TabIndex        =   227
            Top             =   2700
            Width           =   645
         End
         Begin VB.Label LS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   36
            Left            =   5130
            TabIndex        =   226
            Top             =   2490
            Width           =   645
         End
         Begin VB.Label LT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   26
            Left            =   5130
            TabIndex        =   225
            Top             =   2280
            Width           =   645
         End
         Begin VB.Label LS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   26
            Left            =   5130
            TabIndex        =   224
            Top             =   2070
            Width           =   645
         End
         Begin VB.Label LT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   16
            Left            =   5130
            TabIndex        =   223
            Top             =   1860
            Width           =   645
         End
         Begin VB.Label LS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   16
            Left            =   5130
            TabIndex        =   222
            Top             =   1650
            Width           =   645
         End
         Begin VB.Label LT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   107
            Left            =   5850
            TabIndex        =   221
            Top             =   6090
            Width           =   645
         End
         Begin VB.Label LS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   107
            Left            =   5850
            TabIndex        =   220
            Top             =   5880
            Width           =   645
         End
         Begin VB.Label LT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   97
            Left            =   5850
            TabIndex        =   219
            Top             =   5670
            Width           =   645
         End
         Begin VB.Label LS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   97
            Left            =   5850
            TabIndex        =   218
            Top             =   5460
            Width           =   645
         End
         Begin VB.Label LT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   87
            Left            =   5850
            TabIndex        =   217
            Top             =   4830
            Width           =   645
         End
         Begin VB.Label LS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   87
            Left            =   5850
            TabIndex        =   216
            Top             =   4620
            Width           =   645
         End
         Begin VB.Label LT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   77
            Left            =   5850
            TabIndex        =   215
            Top             =   4410
            Width           =   645
         End
         Begin VB.Label LS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   77
            Left            =   5850
            TabIndex        =   214
            Top             =   4200
            Width           =   645
         End
         Begin VB.Label LT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   67
            Left            =   5850
            TabIndex        =   213
            Top             =   3960
            Width           =   645
         End
         Begin VB.Label LS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   67
            Left            =   5850
            TabIndex        =   212
            Top             =   3750
            Width           =   645
         End
         Begin VB.Label LT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   57
            Left            =   5850
            TabIndex        =   211
            Top             =   3540
            Width           =   645
         End
         Begin VB.Label LS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   57
            Left            =   5850
            TabIndex        =   210
            Top             =   3330
            Width           =   645
         End
         Begin VB.Label LT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   47
            Left            =   5850
            TabIndex        =   209
            Top             =   3120
            Width           =   645
         End
         Begin VB.Label LS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   47
            Left            =   5850
            TabIndex        =   208
            Top             =   2910
            Width           =   645
         End
         Begin VB.Label LT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   37
            Left            =   5850
            TabIndex        =   207
            Top             =   2700
            Width           =   645
         End
         Begin VB.Label LS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   37
            Left            =   5850
            TabIndex        =   206
            Top             =   2490
            Width           =   645
         End
         Begin VB.Label LT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   27
            Left            =   5850
            TabIndex        =   205
            Top             =   2280
            Width           =   645
         End
         Begin VB.Label LS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   27
            Left            =   5850
            TabIndex        =   204
            Top             =   2070
            Width           =   645
         End
         Begin VB.Label LT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   17
            Left            =   5850
            TabIndex        =   203
            Top             =   1860
            Width           =   645
         End
         Begin VB.Label LS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   17
            Left            =   5850
            TabIndex        =   202
            Top             =   1650
            Width           =   645
         End
         Begin VB.Label RS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   17
            Left            =   13050
            TabIndex        =   201
            Top             =   1650
            Width           =   645
         End
         Begin VB.Label RT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   17
            Left            =   13050
            TabIndex        =   200
            Top             =   1860
            Width           =   645
         End
         Begin VB.Label RS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   27
            Left            =   13050
            TabIndex        =   199
            Top             =   2070
            Width           =   645
         End
         Begin VB.Label RT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   27
            Left            =   13050
            TabIndex        =   198
            Top             =   2280
            Width           =   645
         End
         Begin VB.Label RS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   37
            Left            =   13050
            TabIndex        =   197
            Top             =   2490
            Width           =   645
         End
         Begin VB.Label RT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   37
            Left            =   13050
            TabIndex        =   196
            Top             =   2700
            Width           =   645
         End
         Begin VB.Label RS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   47
            Left            =   13050
            TabIndex        =   195
            Top             =   2910
            Width           =   645
         End
         Begin VB.Label RT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   47
            Left            =   13050
            TabIndex        =   194
            Top             =   3120
            Width           =   645
         End
         Begin VB.Label RS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   57
            Left            =   13050
            TabIndex        =   193
            Top             =   3330
            Width           =   645
         End
         Begin VB.Label RT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   57
            Left            =   13050
            TabIndex        =   192
            Top             =   3540
            Width           =   645
         End
         Begin VB.Label RS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   67
            Left            =   13050
            TabIndex        =   191
            Top             =   3750
            Width           =   645
         End
         Begin VB.Label RT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   67
            Left            =   13050
            TabIndex        =   190
            Top             =   3960
            Width           =   645
         End
         Begin VB.Label RS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   77
            Left            =   13050
            TabIndex        =   189
            Top             =   4200
            Width           =   645
         End
         Begin VB.Label RT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   77
            Left            =   13050
            TabIndex        =   188
            Top             =   4410
            Width           =   645
         End
         Begin VB.Label RS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   87
            Left            =   13050
            TabIndex        =   187
            Top             =   4620
            Width           =   645
         End
         Begin VB.Label RT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   87
            Left            =   13050
            TabIndex        =   186
            Top             =   4830
            Width           =   645
         End
         Begin VB.Label RS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   97
            Left            =   13050
            TabIndex        =   185
            Top             =   5460
            Width           =   645
         End
         Begin VB.Label RT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   97
            Left            =   13050
            TabIndex        =   184
            Top             =   5670
            Width           =   645
         End
         Begin VB.Label RS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   107
            Left            =   13050
            TabIndex        =   183
            Top             =   5880
            Width           =   645
         End
         Begin VB.Label RT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   107
            Left            =   13050
            TabIndex        =   182
            Top             =   6090
            Width           =   645
         End
         Begin VB.Label RS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   16
            Left            =   12330
            TabIndex        =   181
            Top             =   1650
            Width           =   645
         End
         Begin VB.Label RT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   16
            Left            =   12330
            TabIndex        =   180
            Top             =   1860
            Width           =   645
         End
         Begin VB.Label RS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   26
            Left            =   12330
            TabIndex        =   179
            Top             =   2070
            Width           =   645
         End
         Begin VB.Label RT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   26
            Left            =   12330
            TabIndex        =   178
            Top             =   2280
            Width           =   645
         End
         Begin VB.Label RS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   36
            Left            =   12330
            TabIndex        =   177
            Top             =   2490
            Width           =   645
         End
         Begin VB.Label RT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   36
            Left            =   12330
            TabIndex        =   176
            Top             =   2700
            Width           =   645
         End
         Begin VB.Label RS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   46
            Left            =   12330
            TabIndex        =   175
            Top             =   2910
            Width           =   645
         End
         Begin VB.Label RT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   46
            Left            =   12330
            TabIndex        =   174
            Top             =   3120
            Width           =   645
         End
         Begin VB.Label RS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   56
            Left            =   12330
            TabIndex        =   173
            Top             =   3330
            Width           =   645
         End
         Begin VB.Label RT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   56
            Left            =   12330
            TabIndex        =   172
            Top             =   3540
            Width           =   645
         End
         Begin VB.Label RS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   66
            Left            =   12330
            TabIndex        =   171
            Top             =   3750
            Width           =   645
         End
         Begin VB.Label RT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   66
            Left            =   12330
            TabIndex        =   170
            Top             =   3960
            Width           =   645
         End
         Begin VB.Label RS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   76
            Left            =   12330
            TabIndex        =   169
            Top             =   4200
            Width           =   645
         End
         Begin VB.Label RT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   76
            Left            =   12330
            TabIndex        =   168
            Top             =   4410
            Width           =   645
         End
         Begin VB.Label RS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   86
            Left            =   12330
            TabIndex        =   167
            Top             =   4620
            Width           =   645
         End
         Begin VB.Label RT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   86
            Left            =   12330
            TabIndex        =   166
            Top             =   4830
            Width           =   645
         End
         Begin VB.Label RS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   96
            Left            =   12330
            TabIndex        =   165
            Top             =   5460
            Width           =   645
         End
         Begin VB.Label RT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   96
            Left            =   12330
            TabIndex        =   164
            Top             =   5670
            Width           =   645
         End
         Begin VB.Label RS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   106
            Left            =   12330
            TabIndex        =   163
            Top             =   5880
            Width           =   645
         End
         Begin VB.Label RT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   106
            Left            =   12330
            TabIndex        =   162
            Top             =   6090
            Width           =   645
         End
         Begin VB.Label RS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   15
            Left            =   11580
            TabIndex        =   161
            Top             =   1650
            Width           =   645
         End
         Begin VB.Label RT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   15
            Left            =   11580
            TabIndex        =   160
            Top             =   1860
            Width           =   645
         End
         Begin VB.Label RS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   25
            Left            =   11580
            TabIndex        =   159
            Top             =   2070
            Width           =   645
         End
         Begin VB.Label RT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   25
            Left            =   11580
            TabIndex        =   158
            Top             =   2280
            Width           =   645
         End
         Begin VB.Label RS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   35
            Left            =   11580
            TabIndex        =   157
            Top             =   2490
            Width           =   645
         End
         Begin VB.Label RT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   35
            Left            =   11580
            TabIndex        =   156
            Top             =   2700
            Width           =   645
         End
         Begin VB.Label RS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   45
            Left            =   11580
            TabIndex        =   155
            Top             =   2910
            Width           =   645
         End
         Begin VB.Label RT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   45
            Left            =   11580
            TabIndex        =   154
            Top             =   3120
            Width           =   645
         End
         Begin VB.Label RS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   55
            Left            =   11580
            TabIndex        =   153
            Top             =   3330
            Width           =   645
         End
         Begin VB.Label RT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   55
            Left            =   11580
            TabIndex        =   152
            Top             =   3540
            Width           =   645
         End
         Begin VB.Label RS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   65
            Left            =   11580
            TabIndex        =   151
            Top             =   3750
            Width           =   645
         End
         Begin VB.Label RT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   65
            Left            =   11580
            TabIndex        =   150
            Top             =   3960
            Width           =   645
         End
         Begin VB.Label RS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   75
            Left            =   11580
            TabIndex        =   149
            Top             =   4200
            Width           =   645
         End
         Begin VB.Label RT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   75
            Left            =   11580
            TabIndex        =   148
            Top             =   4410
            Width           =   645
         End
         Begin VB.Label RS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   85
            Left            =   11580
            TabIndex        =   147
            Top             =   4620
            Width           =   645
         End
         Begin VB.Label RT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   85
            Left            =   11580
            TabIndex        =   146
            Top             =   4830
            Width           =   645
         End
         Begin VB.Label RS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   95
            Left            =   11580
            TabIndex        =   145
            Top             =   5460
            Width           =   645
         End
         Begin VB.Label RT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   95
            Left            =   11580
            TabIndex        =   144
            Top             =   5670
            Width           =   645
         End
         Begin VB.Label RS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   105
            Left            =   11580
            TabIndex        =   143
            Top             =   5880
            Width           =   645
         End
         Begin VB.Label RT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   105
            Left            =   11580
            TabIndex        =   142
            Top             =   6090
            Width           =   645
         End
         Begin VB.Label RS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   14
            Left            =   10830
            TabIndex        =   141
            Top             =   1650
            Width           =   645
         End
         Begin VB.Label RT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   14
            Left            =   10830
            TabIndex        =   140
            Top             =   1860
            Width           =   645
         End
         Begin VB.Label RS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   24
            Left            =   10830
            TabIndex        =   139
            Top             =   2070
            Width           =   645
         End
         Begin VB.Label RT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   24
            Left            =   10830
            TabIndex        =   138
            Top             =   2280
            Width           =   645
         End
         Begin VB.Label RS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   34
            Left            =   10830
            TabIndex        =   137
            Top             =   2490
            Width           =   645
         End
         Begin VB.Label RT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   34
            Left            =   10830
            TabIndex        =   136
            Top             =   2700
            Width           =   645
         End
         Begin VB.Label RS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   44
            Left            =   10830
            TabIndex        =   135
            Top             =   2910
            Width           =   645
         End
         Begin VB.Label RT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   44
            Left            =   10830
            TabIndex        =   134
            Top             =   3120
            Width           =   645
         End
         Begin VB.Label RS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   54
            Left            =   10830
            TabIndex        =   133
            Top             =   3330
            Width           =   645
         End
         Begin VB.Label RT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   54
            Left            =   10830
            TabIndex        =   132
            Top             =   3540
            Width           =   645
         End
         Begin VB.Label RS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   64
            Left            =   10830
            TabIndex        =   131
            Top             =   3750
            Width           =   645
         End
         Begin VB.Label RT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   64
            Left            =   10830
            TabIndex        =   130
            Top             =   3960
            Width           =   645
         End
         Begin VB.Label RS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   74
            Left            =   10830
            TabIndex        =   129
            Top             =   4200
            Width           =   645
         End
         Begin VB.Label RT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   74
            Left            =   10830
            TabIndex        =   128
            Top             =   4410
            Width           =   645
         End
         Begin VB.Label RS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   84
            Left            =   10830
            TabIndex        =   127
            Top             =   4620
            Width           =   645
         End
         Begin VB.Label RT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   84
            Left            =   10830
            TabIndex        =   126
            Top             =   4830
            Width           =   645
         End
         Begin VB.Label RS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   94
            Left            =   10830
            TabIndex        =   125
            Top             =   5460
            Width           =   645
         End
         Begin VB.Label RT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   94
            Left            =   10830
            TabIndex        =   124
            Top             =   5670
            Width           =   645
         End
         Begin VB.Label RS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   104
            Left            =   10830
            TabIndex        =   123
            Top             =   5880
            Width           =   645
         End
         Begin VB.Label RT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   104
            Left            =   10830
            TabIndex        =   122
            Top             =   6090
            Width           =   645
         End
         Begin VB.Label RS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   13
            Left            =   10110
            TabIndex        =   121
            Top             =   1650
            Width           =   645
         End
         Begin VB.Label RT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   13
            Left            =   10110
            TabIndex        =   120
            Top             =   1860
            Width           =   645
         End
         Begin VB.Label RS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   23
            Left            =   10110
            TabIndex        =   119
            Top             =   2070
            Width           =   645
         End
         Begin VB.Label RT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   23
            Left            =   10110
            TabIndex        =   118
            Top             =   2280
            Width           =   645
         End
         Begin VB.Label RS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   33
            Left            =   10110
            TabIndex        =   117
            Top             =   2490
            Width           =   645
         End
         Begin VB.Label RT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   33
            Left            =   10110
            TabIndex        =   116
            Top             =   2700
            Width           =   645
         End
         Begin VB.Label RS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   43
            Left            =   10110
            TabIndex        =   115
            Top             =   2910
            Width           =   645
         End
         Begin VB.Label RT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   43
            Left            =   10110
            TabIndex        =   114
            Top             =   3120
            Width           =   645
         End
         Begin VB.Label RS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   53
            Left            =   10110
            TabIndex        =   113
            Top             =   3330
            Width           =   645
         End
         Begin VB.Label RT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   53
            Left            =   10110
            TabIndex        =   112
            Top             =   3540
            Width           =   645
         End
         Begin VB.Label RS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   63
            Left            =   10110
            TabIndex        =   111
            Top             =   3750
            Width           =   645
         End
         Begin VB.Label RT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   63
            Left            =   10110
            TabIndex        =   110
            Top             =   3960
            Width           =   645
         End
         Begin VB.Label RS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   73
            Left            =   10110
            TabIndex        =   109
            Top             =   4200
            Width           =   645
         End
         Begin VB.Label RT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   73
            Left            =   10110
            TabIndex        =   108
            Top             =   4410
            Width           =   645
         End
         Begin VB.Label RS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   83
            Left            =   10110
            TabIndex        =   107
            Top             =   4620
            Width           =   645
         End
         Begin VB.Label RT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   83
            Left            =   10110
            TabIndex        =   106
            Top             =   4830
            Width           =   645
         End
         Begin VB.Label RS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   93
            Left            =   10110
            TabIndex        =   105
            Top             =   5460
            Width           =   645
         End
         Begin VB.Label RT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   93
            Left            =   10110
            TabIndex        =   104
            Top             =   5670
            Width           =   645
         End
         Begin VB.Label RS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   103
            Left            =   10110
            TabIndex        =   103
            Top             =   5880
            Width           =   645
         End
         Begin VB.Label RT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   103
            Left            =   10110
            TabIndex        =   102
            Top             =   6090
            Width           =   645
         End
         Begin VB.Label RC 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "08:00"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   12
            Left            =   8700
            TabIndex        =   101
            Top             =   1890
            Width           =   645
         End
         Begin VB.Label RC 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "08:00"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   11
            Left            =   8700
            TabIndex        =   100
            Top             =   1680
            Width           =   645
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   12
            X1              =   13020
            X2              =   13020
            Y1              =   5430
            Y2              =   6270
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   13
            X1              =   12300
            X2              =   12300
            Y1              =   5430
            Y2              =   6300
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   18
            X1              =   11550
            X2              =   11550
            Y1              =   5430
            Y2              =   6300
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   19
            X1              =   10800
            X2              =   10800
            Y1              =   5430
            Y2              =   6300
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   29
            X1              =   10080
            X2              =   10080
            Y1              =   5430
            Y2              =   6300
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   30
            X1              =   8610
            X2              =   8610
            Y1              =   5430
            Y2              =   6300
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "10±≥Ω√"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   262
            Left            =   7950
            TabIndex        =   99
            Top             =   5940
            Width           =   705
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "9±≥Ω√"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   263
            Left            =   7980
            TabIndex        =   98
            Top             =   5520
            Width           =   585
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "8±≥Ω√"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   264
            Left            =   7980
            TabIndex        =   97
            Top             =   4680
            Width           =   585
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "7±≥Ω√"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   265
            Left            =   7980
            TabIndex        =   96
            Top             =   4260
            Width           =   585
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "6±≥Ω√"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   266
            Left            =   7980
            TabIndex        =   95
            Top             =   3870
            Width           =   585
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "5±≥Ω√"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   267
            Left            =   7980
            TabIndex        =   94
            Top             =   3420
            Width           =   585
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "4±≥Ω√"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   268
            Left            =   7980
            TabIndex        =   93
            Top             =   3000
            Width           =   585
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "3±≥Ω√"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   269
            Left            =   7980
            TabIndex        =   92
            Top             =   2610
            Width           =   585
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "2±≥Ω√"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   270
            Left            =   7980
            TabIndex        =   91
            Top             =   2190
            Width           =   585
         End
         Begin VB.Label RS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   12
            Left            =   9390
            TabIndex        =   90
            Top             =   1650
            Width           =   645
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "1±≥Ω√"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   272
            Left            =   7980
            TabIndex        =   89
            Top             =   1770
            Width           =   585
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "≈‰"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   273
            Left            =   13260
            TabIndex        =   88
            Top             =   1350
            Width           =   315
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "±›"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   274
            Left            =   12540
            TabIndex        =   87
            Top             =   1350
            Width           =   315
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "∏Ò"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   275
            Left            =   11820
            TabIndex        =   86
            Top             =   1350
            Width           =   315
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "ºˆ"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   276
            Left            =   11070
            TabIndex        =   85
            Top             =   1350
            Width           =   315
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "»≠"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   277
            Left            =   10350
            TabIndex        =   84
            Top             =   1350
            Width           =   315
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "ø˘"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   278
            Left            =   9600
            TabIndex        =   83
            Top             =   1350
            Width           =   315
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   31
            X1              =   7920
            X2              =   13770
            Y1              =   5430
            Y2              =   5430
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   32
            X1              =   7920
            X2              =   13770
            Y1              =   5850
            Y2              =   5850
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   33
            X1              =   7920
            X2              =   13770
            Y1              =   5010
            Y2              =   5010
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   34
            X1              =   7920
            X2              =   13770
            Y1              =   4590
            Y2              =   4590
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   35
            X1              =   7920
            X2              =   13770
            Y1              =   4140
            Y2              =   4140
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   36
            X1              =   7920
            X2              =   13770
            Y1              =   3720
            Y2              =   3720
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   37
            X1              =   7920
            X2              =   13770
            Y1              =   3300
            Y2              =   3300
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   38
            X1              =   7920
            X2              =   13770
            Y1              =   2880
            Y2              =   2880
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   39
            X1              =   7920
            X2              =   13770
            Y1              =   2460
            Y2              =   2460
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   40
            X1              =   7920
            X2              =   13770
            Y1              =   2040
            Y2              =   2040
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   41
            X1              =   10080
            X2              =   10080
            Y1              =   1260
            Y2              =   5010
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   42
            X1              =   10800
            X2              =   10800
            Y1              =   1260
            Y2              =   5010
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   43
            X1              =   11550
            X2              =   11550
            Y1              =   1260
            Y2              =   5010
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   44
            X1              =   12300
            X2              =   12300
            Y1              =   1260
            Y2              =   5010
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   45
            X1              =   8610
            X2              =   8610
            Y1              =   1620
            Y2              =   5010
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   46
            X1              =   13020
            X2              =   13020
            Y1              =   1260
            Y2              =   5010
         End
         Begin VB.Label Labels 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "Ω√   ∞£   «•"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   20.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   279
            Left            =   9660
            TabIndex        =   82
            Top             =   300
            Width           =   2235
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            BorderWidth     =   2
            Index           =   48
            X1              =   7920
            X2              =   13770
            Y1              =   1620
            Y2              =   1620
         End
         Begin VB.Shape Boxs 
            BorderColor     =   &H00FF0000&
            BorderWidth     =   2
            Height          =   5055
            Index           =   1
            Left            =   7920
            Top             =   1260
            Width           =   5865
         End
         Begin VB.Label RT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   12
            Left            =   9390
            TabIndex        =   81
            Top             =   1860
            Width           =   645
         End
         Begin VB.Label RC 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "08:00"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   22
            Left            =   8700
            TabIndex        =   80
            Top             =   2310
            Width           =   645
         End
         Begin VB.Label RS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   22
            Left            =   9390
            TabIndex        =   79
            Top             =   2070
            Width           =   645
         End
         Begin VB.Label RT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   22
            Left            =   9390
            TabIndex        =   78
            Top             =   2280
            Width           =   645
         End
         Begin VB.Label RC 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "08:00"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   32
            Left            =   8700
            TabIndex        =   77
            Top             =   2730
            Width           =   645
         End
         Begin VB.Label RC 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "08:00"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   31
            Left            =   8700
            TabIndex        =   76
            Top             =   2520
            Width           =   645
         End
         Begin VB.Label RS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   32
            Left            =   9390
            TabIndex        =   75
            Top             =   2490
            Width           =   645
         End
         Begin VB.Label RT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   32
            Left            =   9390
            TabIndex        =   74
            Top             =   2700
            Width           =   645
         End
         Begin VB.Label RC 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "08:00"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   42
            Left            =   8700
            TabIndex        =   73
            Top             =   3150
            Width           =   645
         End
         Begin VB.Label RC 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "08:00"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   41
            Left            =   8700
            TabIndex        =   72
            Top             =   2940
            Width           =   645
         End
         Begin VB.Label RS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   42
            Left            =   9390
            TabIndex        =   71
            Top             =   2910
            Width           =   645
         End
         Begin VB.Label RT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   42
            Left            =   9390
            TabIndex        =   70
            Top             =   3120
            Width           =   645
         End
         Begin VB.Label RC 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "08:00"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   52
            Left            =   8700
            TabIndex        =   69
            Top             =   3540
            Width           =   645
         End
         Begin VB.Label RC 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "08:00"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   51
            Left            =   8700
            TabIndex        =   68
            Top             =   3330
            Width           =   645
         End
         Begin VB.Label RS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   52
            Left            =   9390
            TabIndex        =   67
            Top             =   3330
            Width           =   645
         End
         Begin VB.Label RT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   52
            Left            =   9390
            TabIndex        =   66
            Top             =   3540
            Width           =   645
         End
         Begin VB.Label RC 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "08:00"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   62
            Left            =   8700
            TabIndex        =   65
            Top             =   3960
            Width           =   645
         End
         Begin VB.Label RC 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "08:00"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   61
            Left            =   8700
            TabIndex        =   64
            Top             =   3750
            Width           =   645
         End
         Begin VB.Label RS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   62
            Left            =   9390
            TabIndex        =   63
            Top             =   3750
            Width           =   645
         End
         Begin VB.Label RT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   62
            Left            =   9390
            TabIndex        =   62
            Top             =   3960
            Width           =   645
         End
         Begin VB.Label RC 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "08:00"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   72
            Left            =   8700
            TabIndex        =   61
            Top             =   4410
            Width           =   645
         End
         Begin VB.Label RC 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "08:00"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   71
            Left            =   8700
            TabIndex        =   60
            Top             =   4200
            Width           =   645
         End
         Begin VB.Label RS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   72
            Left            =   9390
            TabIndex        =   59
            Top             =   4200
            Width           =   645
         End
         Begin VB.Label RT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   72
            Left            =   9390
            TabIndex        =   58
            Top             =   4410
            Width           =   645
         End
         Begin VB.Label RC 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "08:00"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   82
            Left            =   8700
            TabIndex        =   57
            Top             =   4830
            Width           =   645
         End
         Begin VB.Label RC 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "08:00"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   81
            Left            =   8700
            TabIndex        =   56
            Top             =   4620
            Width           =   645
         End
         Begin VB.Label RS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   82
            Left            =   9390
            TabIndex        =   55
            Top             =   4620
            Width           =   645
         End
         Begin VB.Label RT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   82
            Left            =   9390
            TabIndex        =   54
            Top             =   4830
            Width           =   645
         End
         Begin VB.Label RC 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "08:00"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   92
            Left            =   8700
            TabIndex        =   53
            Top             =   5670
            Width           =   645
         End
         Begin VB.Label RC 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "08:00"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   91
            Left            =   8700
            TabIndex        =   52
            Top             =   5460
            Width           =   645
         End
         Begin VB.Label RS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   92
            Left            =   9390
            TabIndex        =   51
            Top             =   5460
            Width           =   645
         End
         Begin VB.Label RT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   92
            Left            =   9390
            TabIndex        =   50
            Top             =   5670
            Width           =   645
         End
         Begin VB.Label RC 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "08:00"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   102
            Left            =   8700
            TabIndex        =   49
            Top             =   6090
            Width           =   645
         End
         Begin VB.Label RC 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "08:00"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   101
            Left            =   8700
            TabIndex        =   48
            Top             =   5880
            Width           =   645
         End
         Begin VB.Label RS 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "æA"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   102
            Left            =   9390
            TabIndex        =   47
            Top             =   5880
            Width           =   645
         End
         Begin VB.Label RT 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¿Ø«œ±’"
            BeginProperty Font 
               Name            =   "±º∏≤"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   102
            Left            =   9390
            TabIndex        =   46
            Top             =   6090
            Width           =   645
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            BorderWidth     =   2
            Index           =   50
            X1              =   9360
            X2              =   9360
            Y1              =   1260
            Y2              =   6300
         End
         Begin VB.Shape FillBOXs2 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  '≈ı∏Ì«œ¡ˆ æ ¿Ω
            BorderStyle     =   0  '≈ı∏Ì
            Height          =   555
            Index           =   0
            Left            =   8640
            Shape           =   4  'µ’±Ÿ ªÁ∞¢«¸
            Top             =   210
            Width           =   4035
         End
         Begin VB.Shape FillBOXs2 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  '≈ı∏Ì«œ¡ˆ æ ¿Ω
            BorderStyle     =   0  '≈ı∏Ì
            Height          =   555
            Index           =   2
            Left            =   1440
            Shape           =   4  'µ’±Ÿ ªÁ∞¢«¸
            Top             =   210
            Width           =   4035
         End
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0082C8E8&
      BorderStyle     =   0  'æ¯¿Ω
      Caption         =   "Frame2"
      Height          =   495
      Left            =   30
      TabIndex        =   36
      Top             =   30
      Width           =   14445
      Begin VB.Frame Frame1 
         BackColor       =   &H00D2EAF5&
         BorderStyle     =   0  'æ¯¿Ω
         Caption         =   "Frame1"
         Height          =   435
         Left            =   30
         TabIndex        =   37
         Top             =   30
         Width           =   14385
         Begin VB.CommandButton cmdShiftRight 
            Caption         =   "¢∫"
            Height          =   375
            Left            =   13920
            TabIndex        =   11
            Top             =   30
            Width           =   465
         End
         Begin VB.CommandButton cmdShiftLeft 
            Caption         =   "¢∏"
            Height          =   375
            Left            =   12720
            TabIndex        =   9
            Top             =   30
            Width           =   465
         End
         Begin VB.CommandButton cmdinFo_in 
            Caption         =   "æ»≥ª ¡∂»∏"
            Height          =   375
            Left            =   8130
            TabIndex        =   6
            Top             =   30
            Width           =   1035
         End
         Begin VB.CommandButton cmdTime_in 
            Caption         =   "Ω√∞£ ¡∂»∏"
            Height          =   375
            Left            =   7020
            TabIndex        =   5
            Top             =   30
            Width           =   1035
         End
         Begin VB.TextBox txtLsn 
            Enabled         =   0   'False
            Height          =   300
            Index           =   1
            Left            =   3150
            TabIndex        =   2
            Text            =   "txtLsn"
            Top             =   67
            Width           =   615
         End
         Begin VB.TextBox txtLsn 
            BackColor       =   &H00C0FFFF&
            Height          =   300
            Index           =   0
            Left            =   1950
            TabIndex        =   1
            Text            =   "txtLsn"
            Top             =   67
            Width           =   1185
         End
         Begin VB.CommandButton cmdFind 
            Caption         =   "Ω√∞£«• ¡∂»∏"
            Height          =   375
            Left            =   5280
            TabIndex        =   4
            Top             =   30
            Width           =   1515
         End
         Begin VB.ComboBox cboKaeyol 
            Height          =   300
            Left            =   450
            Style           =   2  'µÂ∑”¥ŸøÓ ∏Ò∑œ
            TabIndex        =   0
            Top             =   67
            Width           =   1155
         End
         Begin VB.CommandButton cmdPrint 
            Caption         =   "«ˆ¿Á∆‰¿Ã¡ˆ √‚∑¬"
            Height          =   375
            Left            =   9540
            TabIndex        =   7
            Top             =   30
            Width           =   1515
         End
         Begin VB.CommandButton cmdPrintAll 
            Caption         =   "¿¸√º∆‰¿Ã¡ˆ √‚∑¬"
            Height          =   375
            Left            =   11100
            TabIndex        =   8
            Top             =   30
            Width           =   1515
         End
         Begin VB.TextBox txtPage 
            Enabled         =   0   'False
            Height          =   375
            Left            =   13170
            TabIndex        =   10
            Text            =   "txtPage"
            Top             =   30
            Width           =   735
         End
         Begin EditLib.fpMask fpYM 
            Height          =   285
            Left            =   3960
            TabIndex        =   3
            Top             =   60
            Width           =   1005
            _Version        =   196608
            _ExtentX        =   1773
            _ExtentY        =   503
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "±º∏≤"
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
            Mask            =   "######"
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
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "π›"
            BeginProperty Font 
               Name            =   "±º∏≤√º"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   1710
            TabIndex        =   39
            Top             =   120
            Width           =   945
         End
         Begin VB.Label NonPrintLbl 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "∞Ëø≠"
            BeginProperty Font 
               Name            =   "±º∏≤√º"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   3
            Left            =   30
            TabIndex        =   38
            Top             =   120
            Width           =   945
         End
      End
   End
   Begin FPSpread.vaSpread sprLsn 
      Height          =   6255
      Left            =   2790
      TabIndex        =   369
      Top             =   10680
      Width           =   2685
      _Version        =   393216
      _ExtentX        =   4736
      _ExtentY        =   11033
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
      MaxCols         =   2
      ScrollBars      =   2
      SpreadDesigner  =   "PRT050.frx":0000
   End
   Begin FPSpread.vaSpread sprinFo 
      Height          =   4065
      Left            =   15450
      TabIndex        =   13
      Top             =   7440
      Width           =   6045
      _Version        =   393216
      _ExtentX        =   10663
      _ExtentY        =   7170
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
      MaxCols         =   1
      MaxRows         =   11
      ProcessTab      =   -1  'True
      ScrollBars      =   0
      SpreadDesigner  =   "PRT050.frx":1809
   End
   Begin FPSpread.vaSpread sprTime 
      Height          =   5535
      Left            =   15450
      TabIndex        =   12
      Top             =   1650
      Width           =   1425
      _Version        =   393216
      _ExtentX        =   2514
      _ExtentY        =   9763
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
      MaxCols         =   1
      MaxRows         =   22
      ProcessTab      =   -1  'True
      ScrollBars      =   0
      SpreadDesigner  =   "PRT050.frx":1C83
   End
   Begin MSComDlg.CommonDialog dlgPrint 
      Left            =   14640
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "PRT050"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

