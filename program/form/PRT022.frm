VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form PRT022 
   Caption         =   "Ω√∞£«• √‚∑¬ >> ∞≠ªÁ∫∞ Ω√∞£«• √‚∑¬ (≥Î∑Æ¡¯) - CP"
   ClientHeight    =   10440
   ClientLeft      =   6795
   ClientTop       =   3330
   ClientWidth     =   14775
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10440
   ScaleWidth      =   14775
   Begin VB.Frame Frame2 
      BackColor       =   &H0082C8E8&
      BorderStyle     =   0  'æ¯¿Ω
      Caption         =   "Frame2"
      Height          =   495
      Left            =   30
      TabIndex        =   321
      Top             =   30
      Width           =   14445
      Begin VB.Frame Frame1 
         BackColor       =   &H00D2EAF5&
         BorderStyle     =   0  'æ¯¿Ω
         Caption         =   "Frame1"
         Height          =   435
         Left            =   30
         TabIndex        =   322
         Top             =   30
         Width           =   14385
         Begin VB.TextBox txtPage 
            Enabled         =   0   'False
            Height          =   375
            Left            =   13140
            TabIndex        =   332
            Text            =   "txtPage"
            Top             =   30
            Width           =   735
         End
         Begin VB.CommandButton cmdShiftLeft 
            Caption         =   "¢∏"
            Height          =   375
            Left            =   12690
            TabIndex        =   331
            Top             =   30
            Width           =   465
         End
         Begin VB.CommandButton cmdShiftRight 
            Caption         =   "¢∫"
            Height          =   375
            Left            =   13890
            TabIndex        =   330
            Top             =   30
            Width           =   465
         End
         Begin VB.CommandButton cmdPrintAll 
            Caption         =   "¿¸√º∆‰¿Ã¡ˆ √‚∑¬"
            Height          =   375
            Left            =   11130
            TabIndex        =   329
            Top             =   30
            Width           =   1515
         End
         Begin VB.CommandButton cmdPrint 
            Caption         =   "«ˆ¿Á∆‰¿Ã¡ˆ √‚∑¬"
            Height          =   375
            Left            =   9570
            TabIndex        =   328
            Top             =   30
            Width           =   1515
         End
         Begin VB.CommandButton cmdFind 
            Caption         =   "Ω√∞£«• ¡∂»∏"
            Height          =   375
            Left            =   3960
            TabIndex        =   327
            Top             =   30
            Width           =   1515
         End
         Begin VB.TextBox txtTeacher 
            BackColor       =   &H00C0FFFF&
            Height          =   300
            Index           =   0
            Left            =   690
            TabIndex        =   326
            Text            =   "txtTeacher"
            Top             =   67
            Width           =   1185
         End
         Begin VB.TextBox txtTeacher 
            Enabled         =   0   'False
            Height          =   300
            Index           =   1
            Left            =   1890
            TabIndex        =   325
            Text            =   "txtTeacher"
            Top             =   67
            Width           =   615
         End
         Begin VB.CommandButton cmdTime_in 
            Caption         =   "Ω√∞£ ¡∂»∏"
            Height          =   375
            Left            =   6330
            TabIndex        =   324
            Top             =   30
            Width           =   1035
         End
         Begin VB.CommandButton cmdinFo_in 
            Caption         =   "æ»≥ª ¡∂»∏"
            Height          =   375
            Left            =   7950
            TabIndex        =   323
            Top             =   30
            Width           =   1035
         End
         Begin EditLib.fpMask fpYM 
            Height          =   285
            Left            =   2910
            TabIndex        =   333
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
            Caption         =   "∞≠ªÁ"
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
            Left            =   270
            TabIndex        =   334
            Top             =   120
            Width           =   945
         End
      End
   End
   Begin VB.PictureBox pReportControl 
      BorderStyle     =   0  'æ¯¿Ω
      Height          =   9765
      Left            =   30
      ScaleHeight     =   9765
      ScaleWidth      =   14445
      TabIndex        =   0
      Top             =   540
      Width           =   14445
      Begin VB.PictureBox pReportViewer 
         BackColor       =   &H00FFFFFF&
         Height          =   9765
         Left            =   0
         ScaleHeight     =   9705
         ScaleWidth      =   14175
         TabIndex        =   2
         Top             =   0
         Width           =   14235
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
            Index           =   3
            Left            =   1500
            TabIndex        =   341
            Text            =   "LTB"
            Top             =   5070
            Width           =   645
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
            Index           =   4
            Left            =   1500
            TabIndex        =   340
            Text            =   "LTB"
            Top             =   5280
            Width           =   615
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
            Index           =   5
            Left            =   2580
            TabIndex        =   339
            Text            =   "LTB"
            Top             =   5190
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
            Index           =   3
            Left            =   8700
            TabIndex        =   338
            Text            =   "RTB"
            Top             =   5070
            Width           =   645
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
            Index           =   4
            Left            =   8700
            TabIndex        =   337
            Text            =   "RTB"
            Top             =   5280
            Width           =   615
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
            Index           =   5
            Left            =   9810
            TabIndex        =   336
            Text            =   "RTB"
            Top             =   5190
            Width           =   3225
         End
         Begin VB.TextBox MR 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'æ¯¿Ω
            Height          =   270
            Index           =   9
            Left            =   7920
            TabIndex        =   28
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
            TabIndex        =   27
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
            TabIndex        =   26
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
            TabIndex        =   25
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
            TabIndex        =   24
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
            TabIndex        =   23
            Text            =   "ML"
            Top             =   8520
            Width           =   5895
         End
         Begin VB.TextBox ML 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'æ¯¿Ω
            Height          =   270
            Index           =   0
            Left            =   720
            TabIndex        =   22
            Text            =   "ML"
            Top             =   6630
            Width           =   5895
         End
         Begin VB.TextBox ML 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'æ¯¿Ω
            Height          =   270
            Index           =   1
            Left            =   720
            TabIndex        =   21
            Text            =   "ML"
            Top             =   6900
            Width           =   5895
         End
         Begin VB.TextBox ML 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'æ¯¿Ω
            Height          =   270
            Index           =   2
            Left            =   720
            TabIndex        =   20
            Text            =   "ML"
            Top             =   7170
            Width           =   5895
         End
         Begin VB.TextBox ML 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'æ¯¿Ω
            Height          =   270
            Index           =   3
            Left            =   720
            TabIndex        =   19
            Text            =   "ML"
            Top             =   7440
            Width           =   5895
         End
         Begin VB.TextBox ML 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'æ¯¿Ω
            Height          =   270
            Index           =   4
            Left            =   720
            TabIndex        =   18
            Text            =   "ML"
            Top             =   7710
            Width           =   5895
         End
         Begin VB.TextBox ML 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'æ¯¿Ω
            Height          =   270
            Index           =   5
            Left            =   720
            TabIndex        =   17
            Text            =   "ML"
            Top             =   7980
            Width           =   5895
         End
         Begin VB.TextBox ML 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'æ¯¿Ω
            Height          =   270
            Index           =   6
            Left            =   720
            TabIndex        =   16
            Text            =   "ML"
            Top             =   8250
            Width           =   5895
         End
         Begin VB.TextBox MR 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'æ¯¿Ω
            Height          =   270
            Index           =   0
            Left            =   7920
            TabIndex        =   15
            Text            =   "MR"
            Top             =   6630
            Width           =   5895
         End
         Begin VB.TextBox MR 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'æ¯¿Ω
            Height          =   270
            Index           =   1
            Left            =   7920
            TabIndex        =   14
            Text            =   "MR"
            Top             =   6900
            Width           =   5895
         End
         Begin VB.TextBox MR 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'æ¯¿Ω
            Height          =   270
            Index           =   2
            Left            =   7920
            TabIndex        =   13
            Text            =   "MR"
            Top             =   7170
            Width           =   5895
         End
         Begin VB.TextBox MR 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'æ¯¿Ω
            Height          =   270
            Index           =   3
            Left            =   7920
            TabIndex        =   12
            Text            =   "MR"
            Top             =   7440
            Width           =   5895
         End
         Begin VB.TextBox MR 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'æ¯¿Ω
            Height          =   270
            Index           =   4
            Left            =   7920
            TabIndex        =   11
            Text            =   "MR"
            Top             =   7710
            Width           =   5895
         End
         Begin VB.TextBox MR 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'æ¯¿Ω
            Height          =   270
            Index           =   5
            Left            =   7920
            TabIndex        =   10
            Text            =   "MR"
            Top             =   7980
            Width           =   5895
         End
         Begin VB.TextBox MR 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'æ¯¿Ω
            Height          =   270
            Index           =   6
            Left            =   7920
            TabIndex        =   9
            Text            =   "MR"
            Top             =   8250
            Width           =   5895
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
            TabIndex        =   8
            Text            =   "LTB"
            Top             =   3360
            Width           =   645
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
            TabIndex        =   7
            Text            =   "LTB"
            Top             =   3570
            Width           =   615
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
            TabIndex        =   6
            Text            =   "LTB"
            Top             =   3480
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
            Index           =   1
            Left            =   8700
            TabIndex        =   5
            Text            =   "RTB"
            Top             =   3360
            Width           =   645
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
            TabIndex        =   4
            Text            =   "RTB"
            Top             =   3570
            Width           =   615
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
            TabIndex        =   3
            Text            =   "RTB"
            Top             =   3480
            Width           =   3225
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   50
            X1              =   2880
            X2              =   2880
            Y1              =   1290
            Y2              =   3330
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   6
            X1              =   3600
            X2              =   3600
            Y1              =   1290
            Y2              =   3330
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   5
            X1              =   4350
            X2              =   4350
            Y1              =   1290
            Y2              =   3330
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   2
            X1              =   5100
            X2              =   5100
            Y1              =   1290
            Y2              =   3330
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   4
            X1              =   1410
            X2              =   1410
            Y1              =   1650
            Y2              =   3330
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   3
            X1              =   5820
            X2              =   5820
            Y1              =   1290
            Y2              =   3330
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   41
            X1              =   5820
            X2              =   5820
            Y1              =   3750
            Y2              =   5040
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   42
            X1              =   5100
            X2              =   5100
            Y1              =   3750
            Y2              =   5040
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   43
            X1              =   4350
            X2              =   4350
            Y1              =   3750
            Y2              =   5040
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   44
            X1              =   3600
            X2              =   3600
            Y1              =   3750
            Y2              =   5040
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   46
            X1              =   2880
            X2              =   2880
            Y1              =   3750
            Y2              =   5040
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   45
            X1              =   1410
            X2              =   1410
            Y1              =   3750
            Y2              =   5040
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   1
            X1              =   10080
            X2              =   10080
            Y1              =   1290
            Y2              =   3330
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   49
            X1              =   10800
            X2              =   10800
            Y1              =   1290
            Y2              =   3330
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   51
            X1              =   11550
            X2              =   11550
            Y1              =   1290
            Y2              =   3330
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   52
            X1              =   12300
            X2              =   12300
            Y1              =   1290
            Y2              =   3330
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   53
            X1              =   8610
            X2              =   8610
            Y1              =   1650
            Y2              =   3330
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   54
            X1              =   13020
            X2              =   13020
            Y1              =   1290
            Y2              =   3330
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   55
            X1              =   13020
            X2              =   13020
            Y1              =   3750
            Y2              =   5040
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   56
            X1              =   12300
            X2              =   12300
            Y1              =   3750
            Y2              =   5040
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   57
            X1              =   11550
            X2              =   11550
            Y1              =   3750
            Y2              =   5040
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   58
            X1              =   10800
            X2              =   10800
            Y1              =   3750
            Y2              =   5040
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   59
            X1              =   10080
            X2              =   10080
            Y1              =   3750
            Y2              =   5040
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   60
            X1              =   8610
            X2              =   8610
            Y1              =   3750
            Y2              =   5040
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
            TabIndex        =   320
            Top             =   6120
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
            TabIndex        =   319
            Top             =   5910
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
            TabIndex        =   318
            Top             =   5910
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
            TabIndex        =   317
            Top             =   6120
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
            TabIndex        =   316
            Top             =   5700
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
            TabIndex        =   315
            Top             =   5490
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
            TabIndex        =   314
            Top             =   5490
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
            TabIndex        =   313
            Top             =   5700
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
            TabIndex        =   312
            Top             =   4860
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
            TabIndex        =   311
            Top             =   4650
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
            TabIndex        =   310
            Top             =   4650
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
            TabIndex        =   309
            Top             =   4860
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
            TabIndex        =   308
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
            Index           =   62
            Left            =   9390
            TabIndex        =   307
            Top             =   4200
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
            TabIndex        =   306
            Top             =   4200
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
            TabIndex        =   305
            Top             =   4410
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
            TabIndex        =   304
            Top             =   3990
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
            TabIndex        =   303
            Top             =   3780
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
            TabIndex        =   302
            Top             =   3780
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
            TabIndex        =   301
            Top             =   3990
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
            TabIndex        =   300
            Top             =   3150
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
            TabIndex        =   299
            Top             =   2940
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
            TabIndex        =   298
            Top             =   2970
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
            TabIndex        =   297
            Top             =   3180
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
            TabIndex        =   296
            Top             =   2730
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
            TabIndex        =   295
            Top             =   2520
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
            TabIndex        =   294
            Top             =   2550
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
            TabIndex        =   293
            Top             =   2760
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
            TabIndex        =   292
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
            TabIndex        =   291
            Top             =   2100
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
            TabIndex        =   290
            Top             =   2340
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
            Index           =   12
            Left            =   9390
            TabIndex        =   289
            Top             =   1890
            Width           =   645
         End
         Begin VB.Shape Boxs 
            BorderColor     =   &H00FF0000&
            BorderWidth     =   2
            Height          =   5055
            Index           =   1
            Left            =   7920
            Top             =   1290
            Width           =   5865
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            BorderWidth     =   2
            Index           =   48
            X1              =   7920
            X2              =   13770
            Y1              =   1650
            Y2              =   1650
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
            TabIndex        =   288
            Top             =   330
            Width           =   2235
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            BorderWidth     =   2
            Index           =   47
            X1              =   9360
            X2              =   9360
            Y1              =   1290
            Y2              =   6330
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   40
            X1              =   7920
            X2              =   13770
            Y1              =   2070
            Y2              =   2070
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   39
            X1              =   7920
            X2              =   13770
            Y1              =   2490
            Y2              =   2490
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   38
            X1              =   7920
            X2              =   13770
            Y1              =   2910
            Y2              =   2910
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   37
            X1              =   7920
            X2              =   13770
            Y1              =   3330
            Y2              =   3330
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   36
            X1              =   7920
            X2              =   13770
            Y1              =   3750
            Y2              =   3750
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   35
            X1              =   7920
            X2              =   13770
            Y1              =   4170
            Y2              =   4170
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   34
            X1              =   7920
            X2              =   13770
            Y1              =   4620
            Y2              =   4620
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   33
            X1              =   7920
            X2              =   13770
            Y1              =   5040
            Y2              =   5040
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   32
            X1              =   7920
            X2              =   13770
            Y1              =   5880
            Y2              =   5880
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   31
            X1              =   7920
            X2              =   13770
            Y1              =   5460
            Y2              =   5460
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
            TabIndex        =   287
            Top             =   1380
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
            TabIndex        =   286
            Top             =   1380
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
            TabIndex        =   285
            Top             =   1380
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
            TabIndex        =   284
            Top             =   1380
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
            TabIndex        =   283
            Top             =   1380
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
            Index           =   273
            Left            =   13260
            TabIndex        =   282
            Top             =   1380
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
            Index           =   272
            Left            =   7980
            TabIndex        =   281
            Top             =   1800
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
            TabIndex        =   280
            Top             =   1680
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
            Index           =   270
            Left            =   7980
            TabIndex        =   279
            Top             =   2220
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
            TabIndex        =   278
            Top             =   2640
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
            TabIndex        =   277
            Top             =   3030
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
            TabIndex        =   276
            Top             =   3870
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
            TabIndex        =   275
            Top             =   4290
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
            TabIndex        =   274
            Top             =   4710
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
            TabIndex        =   273
            Top             =   5550
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
            Index           =   263
            Left            =   7980
            TabIndex        =   272
            Top             =   6000
            Width           =   585
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   30
            X1              =   8610
            X2              =   8610
            Y1              =   5460
            Y2              =   6330
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   29
            X1              =   10080
            X2              =   10080
            Y1              =   5460
            Y2              =   6330
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   19
            X1              =   10800
            X2              =   10800
            Y1              =   5460
            Y2              =   6330
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   18
            X1              =   11550
            X2              =   11550
            Y1              =   5460
            Y2              =   6330
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   13
            X1              =   12300
            X2              =   12300
            Y1              =   5460
            Y2              =   6330
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   12
            X1              =   13020
            X2              =   13020
            Y1              =   5460
            Y2              =   6330
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
            TabIndex        =   271
            Top             =   1710
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
            TabIndex        =   270
            Top             =   1920
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
            TabIndex        =   269
            Top             =   6120
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
            TabIndex        =   268
            Top             =   5910
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
            TabIndex        =   267
            Top             =   5700
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
            TabIndex        =   266
            Top             =   5490
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
            TabIndex        =   265
            Top             =   4860
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
            TabIndex        =   264
            Top             =   4650
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
            TabIndex        =   263
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
            Index           =   63
            Left            =   10110
            TabIndex        =   262
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
            Index           =   53
            Left            =   10110
            TabIndex        =   261
            Top             =   3990
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
            TabIndex        =   260
            Top             =   3780
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
            TabIndex        =   259
            Top             =   3150
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
            TabIndex        =   258
            Top             =   2940
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
            TabIndex        =   257
            Top             =   2730
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
            TabIndex        =   256
            Top             =   2520
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
            TabIndex        =   255
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
            Index           =   23
            Left            =   10110
            TabIndex        =   254
            Top             =   2100
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
            TabIndex        =   253
            Top             =   1890
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
            TabIndex        =   252
            Top             =   1680
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
            TabIndex        =   251
            Top             =   6120
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
            TabIndex        =   250
            Top             =   5910
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
            TabIndex        =   249
            Top             =   5700
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
            TabIndex        =   248
            Top             =   5490
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
            TabIndex        =   247
            Top             =   4860
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
            TabIndex        =   246
            Top             =   4650
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
            TabIndex        =   245
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
            Index           =   64
            Left            =   10830
            TabIndex        =   244
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
            Index           =   54
            Left            =   10830
            TabIndex        =   243
            Top             =   3990
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
            TabIndex        =   242
            Top             =   3780
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
            TabIndex        =   241
            Top             =   3150
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
            TabIndex        =   240
            Top             =   2940
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
            TabIndex        =   239
            Top             =   2730
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
            TabIndex        =   238
            Top             =   2520
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
            TabIndex        =   237
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
            Index           =   24
            Left            =   10830
            TabIndex        =   236
            Top             =   2100
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
            TabIndex        =   235
            Top             =   1890
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
            TabIndex        =   234
            Top             =   1680
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
            TabIndex        =   233
            Top             =   6120
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
            TabIndex        =   232
            Top             =   5910
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
            TabIndex        =   231
            Top             =   5700
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
            TabIndex        =   230
            Top             =   5490
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
            TabIndex        =   229
            Top             =   4860
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
            TabIndex        =   228
            Top             =   4650
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
            TabIndex        =   227
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
            Index           =   65
            Left            =   11580
            TabIndex        =   226
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
            Index           =   55
            Left            =   11580
            TabIndex        =   225
            Top             =   3990
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
            TabIndex        =   224
            Top             =   3780
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
            TabIndex        =   223
            Top             =   3150
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
            TabIndex        =   222
            Top             =   2940
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
            TabIndex        =   221
            Top             =   2730
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
            TabIndex        =   220
            Top             =   2520
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
            TabIndex        =   219
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
            Index           =   25
            Left            =   11580
            TabIndex        =   218
            Top             =   2100
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
            TabIndex        =   217
            Top             =   1890
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
            TabIndex        =   216
            Top             =   1680
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
            TabIndex        =   215
            Top             =   6120
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
            TabIndex        =   214
            Top             =   5910
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
            TabIndex        =   213
            Top             =   5700
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
            TabIndex        =   212
            Top             =   5490
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
            TabIndex        =   211
            Top             =   4860
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
            TabIndex        =   210
            Top             =   4650
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
            TabIndex        =   209
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
            Index           =   66
            Left            =   12330
            TabIndex        =   208
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
            Index           =   56
            Left            =   12330
            TabIndex        =   207
            Top             =   3990
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
            TabIndex        =   206
            Top             =   3780
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
            TabIndex        =   205
            Top             =   3150
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
            TabIndex        =   204
            Top             =   2940
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
            TabIndex        =   203
            Top             =   2730
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
            TabIndex        =   202
            Top             =   2520
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
            TabIndex        =   201
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
            Index           =   26
            Left            =   12330
            TabIndex        =   200
            Top             =   2100
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
            TabIndex        =   199
            Top             =   1890
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
            TabIndex        =   198
            Top             =   1680
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
            TabIndex        =   197
            Top             =   6120
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
            TabIndex        =   196
            Top             =   5910
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
            TabIndex        =   195
            Top             =   5700
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
            TabIndex        =   194
            Top             =   5490
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
            TabIndex        =   193
            Top             =   4860
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
            TabIndex        =   192
            Top             =   4650
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
            TabIndex        =   191
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
            Index           =   67
            Left            =   13050
            TabIndex        =   190
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
            Index           =   57
            Left            =   13050
            TabIndex        =   189
            Top             =   3990
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
            TabIndex        =   188
            Top             =   3780
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
            TabIndex        =   187
            Top             =   3150
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
            TabIndex        =   186
            Top             =   2940
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
            TabIndex        =   185
            Top             =   2730
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
            TabIndex        =   184
            Top             =   2520
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
            TabIndex        =   183
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
            Index           =   27
            Left            =   13050
            TabIndex        =   182
            Top             =   2100
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
            TabIndex        =   181
            Top             =   1890
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
            TabIndex        =   180
            Top             =   1680
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
            TabIndex        =   179
            Top             =   1680
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
            TabIndex        =   178
            Top             =   1890
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
            TabIndex        =   177
            Top             =   2100
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
            TabIndex        =   176
            Top             =   2310
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
            TabIndex        =   175
            Top             =   2520
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
            TabIndex        =   174
            Top             =   2730
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
            TabIndex        =   173
            Top             =   2940
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
            TabIndex        =   172
            Top             =   3150
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
            TabIndex        =   171
            Top             =   3780
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
            TabIndex        =   170
            Top             =   3990
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
            TabIndex        =   169
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
            TabIndex        =   168
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
            TabIndex        =   167
            Top             =   4650
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
            TabIndex        =   166
            Top             =   4860
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
            TabIndex        =   165
            Top             =   5490
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
            TabIndex        =   164
            Top             =   5700
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
            TabIndex        =   163
            Top             =   5910
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
            TabIndex        =   162
            Top             =   6120
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
            TabIndex        =   161
            Top             =   1680
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
            TabIndex        =   160
            Top             =   1890
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
            TabIndex        =   159
            Top             =   2100
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
            TabIndex        =   158
            Top             =   2310
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
            TabIndex        =   157
            Top             =   2520
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
            TabIndex        =   156
            Top             =   2730
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
            TabIndex        =   155
            Top             =   2940
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
            TabIndex        =   154
            Top             =   3150
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
            TabIndex        =   153
            Top             =   3780
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
            TabIndex        =   152
            Top             =   3990
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
            TabIndex        =   151
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
            TabIndex        =   150
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
            TabIndex        =   149
            Top             =   4650
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
            TabIndex        =   148
            Top             =   4860
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
            TabIndex        =   147
            Top             =   5490
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
            TabIndex        =   146
            Top             =   5700
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
            TabIndex        =   145
            Top             =   5910
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
            TabIndex        =   144
            Top             =   6120
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
            TabIndex        =   143
            Top             =   1680
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
            TabIndex        =   142
            Top             =   1890
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
            TabIndex        =   141
            Top             =   2100
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
            TabIndex        =   140
            Top             =   2310
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
            TabIndex        =   139
            Top             =   2520
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
            TabIndex        =   138
            Top             =   2730
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
            TabIndex        =   137
            Top             =   2940
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
            TabIndex        =   136
            Top             =   3150
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
            TabIndex        =   135
            Top             =   3780
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
            TabIndex        =   134
            Top             =   3990
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
            TabIndex        =   133
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
            TabIndex        =   132
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
            TabIndex        =   131
            Top             =   4650
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
            TabIndex        =   130
            Top             =   4860
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
            TabIndex        =   129
            Top             =   5490
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
            TabIndex        =   128
            Top             =   5700
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
            TabIndex        =   127
            Top             =   5910
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
            TabIndex        =   126
            Top             =   6120
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
            TabIndex        =   125
            Top             =   1680
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
            TabIndex        =   124
            Top             =   1890
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
            TabIndex        =   123
            Top             =   2100
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
            TabIndex        =   122
            Top             =   2310
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
            TabIndex        =   121
            Top             =   2520
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
            TabIndex        =   120
            Top             =   2730
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
            TabIndex        =   119
            Top             =   2940
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
            TabIndex        =   118
            Top             =   3150
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
            TabIndex        =   117
            Top             =   3780
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
            TabIndex        =   116
            Top             =   3990
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
            TabIndex        =   115
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
            TabIndex        =   114
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
            TabIndex        =   113
            Top             =   4650
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
            TabIndex        =   112
            Top             =   4860
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
            TabIndex        =   111
            Top             =   5490
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
            TabIndex        =   110
            Top             =   5700
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
            TabIndex        =   109
            Top             =   5910
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
            TabIndex        =   108
            Top             =   6120
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
            TabIndex        =   107
            Top             =   1680
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
            TabIndex        =   106
            Top             =   1890
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
            TabIndex        =   105
            Top             =   2100
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
            TabIndex        =   104
            Top             =   2310
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
            TabIndex        =   103
            Top             =   2520
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
            TabIndex        =   102
            Top             =   2730
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
            TabIndex        =   101
            Top             =   2940
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
            TabIndex        =   100
            Top             =   3150
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
            TabIndex        =   99
            Top             =   3780
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
            TabIndex        =   98
            Top             =   3990
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
            TabIndex        =   97
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
            TabIndex        =   96
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
            TabIndex        =   95
            Top             =   4650
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
            TabIndex        =   94
            Top             =   4860
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
            TabIndex        =   93
            Top             =   5490
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
            TabIndex        =   92
            Top             =   5700
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
            TabIndex        =   91
            Top             =   5910
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
            TabIndex        =   90
            Top             =   6120
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
            TabIndex        =   89
            Top             =   1920
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
            Index           =   11
            Left            =   1500
            TabIndex        =   88
            Top             =   1710
            Width           =   645
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   27
            X1              =   5820
            X2              =   5820
            Y1              =   5460
            Y2              =   6330
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   26
            X1              =   5100
            X2              =   5100
            Y1              =   5460
            Y2              =   6330
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   25
            X1              =   4350
            X2              =   4350
            Y1              =   5460
            Y2              =   6330
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   24
            X1              =   3600
            X2              =   3600
            Y1              =   5460
            Y2              =   6330
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   23
            X1              =   2880
            X2              =   2880
            Y1              =   5460
            Y2              =   6330
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   22
            X1              =   1410
            X2              =   1410
            Y1              =   5460
            Y2              =   6330
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
            TabIndex        =   87
            Top             =   6000
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
            TabIndex        =   86
            Top             =   5550
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
            TabIndex        =   85
            Top             =   4710
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
            TabIndex        =   84
            Top             =   4290
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
            TabIndex        =   83
            Top             =   3870
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
            TabIndex        =   82
            Top             =   3030
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
            TabIndex        =   81
            Top             =   2640
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
            Index           =   8
            Left            =   780
            TabIndex        =   80
            Top             =   2220
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
            TabIndex        =   79
            Top             =   1680
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
            Index           =   6
            Left            =   780
            TabIndex        =   78
            Top             =   1800
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
            Index           =   5
            Left            =   6060
            TabIndex        =   77
            Top             =   1380
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
            TabIndex        =   76
            Top             =   1380
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
            TabIndex        =   75
            Top             =   1380
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
            TabIndex        =   74
            Top             =   1380
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
            TabIndex        =   73
            Top             =   1380
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
            Index           =   0
            Left            =   2400
            TabIndex        =   72
            Top             =   1380
            Width           =   315
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   8
            X1              =   720
            X2              =   6570
            Y1              =   5460
            Y2              =   5460
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   21
            X1              =   720
            X2              =   6570
            Y1              =   5880
            Y2              =   5880
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   20
            X1              =   720
            X2              =   6570
            Y1              =   5040
            Y2              =   5040
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   17
            X1              =   720
            X2              =   6570
            Y1              =   4170
            Y2              =   4170
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   16
            X1              =   720
            X2              =   6570
            Y1              =   4620
            Y2              =   4620
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   14
            X1              =   720
            X2              =   6570
            Y1              =   3750
            Y2              =   3750
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   11
            X1              =   720
            X2              =   6570
            Y1              =   3330
            Y2              =   3330
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   10
            X1              =   720
            X2              =   6570
            Y1              =   2910
            Y2              =   2910
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   9
            X1              =   720
            X2              =   6570
            Y1              =   2490
            Y2              =   2490
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            Index           =   7
            X1              =   720
            X2              =   6570
            Y1              =   2070
            Y2              =   2070
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            BorderWidth     =   2
            Index           =   0
            X1              =   2160
            X2              =   2160
            Y1              =   1290
            Y2              =   6330
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
            TabIndex        =   71
            Top             =   330
            Width           =   2235
         End
         Begin VB.Line Lines 
            BorderColor     =   &H00FF0000&
            BorderWidth     =   2
            Index           =   28
            X1              =   720
            X2              =   6570
            Y1              =   1650
            Y2              =   1650
         End
         Begin VB.Shape Boxs 
            BorderColor     =   &H00FF0000&
            BorderWidth     =   2
            Height          =   5055
            Index           =   0
            Left            =   720
            Top             =   1290
            Width           =   5865
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
            TabIndex        =   70
            Top             =   1890
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
            TabIndex        =   69
            Top             =   2340
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
            TabIndex        =   68
            Top             =   2130
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
            TabIndex        =   67
            Top             =   2100
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
            TabIndex        =   66
            Top             =   2310
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
            TabIndex        =   65
            Top             =   2760
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
            TabIndex        =   64
            Top             =   2550
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
            TabIndex        =   63
            Top             =   2520
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
            TabIndex        =   62
            Top             =   2730
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
            TabIndex        =   61
            Top             =   3180
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
            TabIndex        =   60
            Top             =   2970
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
            TabIndex        =   59
            Top             =   2940
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
            TabIndex        =   58
            Top             =   3150
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
            TabIndex        =   57
            Top             =   3990
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
            TabIndex        =   56
            Top             =   3780
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
            TabIndex        =   55
            Top             =   3780
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
            TabIndex        =   54
            Top             =   3990
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
            TabIndex        =   53
            Top             =   4410
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
            TabIndex        =   52
            Top             =   4200
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
            TabIndex        =   51
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
            Index           =   62
            Left            =   2190
            TabIndex        =   50
            Top             =   4410
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
            TabIndex        =   49
            Top             =   4860
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
            TabIndex        =   48
            Top             =   4650
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
            TabIndex        =   47
            Top             =   4650
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
            TabIndex        =   46
            Top             =   4860
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
            TabIndex        =   45
            Top             =   5700
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
            TabIndex        =   44
            Top             =   5490
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
            TabIndex        =   43
            Top             =   5490
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
            TabIndex        =   42
            Top             =   5700
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
            TabIndex        =   41
            Top             =   6120
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
            TabIndex        =   40
            Top             =   5910
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
            TabIndex        =   39
            Top             =   5910
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
            TabIndex        =   38
            Top             =   6120
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
            TabIndex        =   37
            Top             =   2130
            Width           =   645
         End
         Begin VB.Label LHD 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "∞≠ªÁ : ¿Ø«œ±’"
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
            Left            =   810
            TabIndex        =   36
            Top             =   1050
            Width           =   1245
         End
         Begin VB.Label LHD 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¥„¿” : "
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
            TabIndex        =   35
            Top             =   1050
            Width           =   1395
         End
         Begin VB.Label LHD 
            BackStyle       =   0  '≈ı∏Ì
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
            Left            =   3600
            TabIndex        =   34
            Top             =   1680
            Width           =   1125
         End
         Begin VB.Label LHD 
            BackStyle       =   0  '≈ı∏Ì
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
            Left            =   5430
            TabIndex        =   33
            Top             =   1680
            Width           =   1215
         End
         Begin VB.Label RHD 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "∞≠ªÁ : ¿Ø«œ±’"
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
            Left            =   8010
            TabIndex        =   32
            Top             =   1050
            Width           =   1245
         End
         Begin VB.Label RHD 
            BackStyle       =   0  '≈ı∏Ì
            Caption         =   "¥„¿” : "
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
            Left            =   9360
            TabIndex        =   31
            Top             =   1050
            Width           =   1395
         End
         Begin VB.Label RHD 
            BackStyle       =   0  '≈ı∏Ì
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
            Left            =   10800
            TabIndex        =   30
            Top             =   1680
            Width           =   1125
         End
         Begin VB.Label RHD 
            BackStyle       =   0  '≈ı∏Ì
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
            Left            =   12630
            TabIndex        =   29
            Top             =   1680
            Width           =   1215
         End
         Begin VB.Shape FillBOXs2 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  '≈ı∏Ì«œ¡ˆ æ ¿Ω
            BorderStyle     =   0  '≈ı∏Ì
            Height          =   555
            Index           =   0
            Left            =   8640
            Shape           =   4  'µ’±Ÿ ªÁ∞¢«¸
            Top             =   240
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
            Top             =   240
            Width           =   4035
         End
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   9765
         Left            =   14220
         Max             =   1
         TabIndex        =   1
         Top             =   0
         Width           =   225
      End
   End
   Begin FPSpread.vaSpread sprTeacher 
      Height          =   6255
      Left            =   1860
      TabIndex        =   335
      Top             =   10590
      Width           =   1905
      _Version        =   393216
      _ExtentX        =   3360
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
      SpreadDesigner  =   "PRT022.frx":0000
   End
   Begin MSComDlg.CommonDialog dlgPrint 
      Left            =   14640
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin FPSpread.vaSpread sprinFo 
      Height          =   4395
      Left            =   15150
      TabIndex        =   342
      Top             =   5880
      Width           =   6045
      _Version        =   393216
      _ExtentX        =   10663
      _ExtentY        =   7752
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
      MaxRows         =   12
      ProcessTab      =   -1  'True
      ScrollBars      =   0
      SpreadDesigner  =   "PRT022.frx":1853
   End
   Begin FPSpread.vaSpread sprTime 
      Height          =   5535
      Left            =   15150
      TabIndex        =   343
      Top             =   90
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
      SpreadDesigner  =   "PRT022.frx":1D44
   End
End
Attribute VB_Name = "PRT022"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################################################
'   Ω√ Ω∫ ≈€  ∏Ì : ¥Îº∫«–ø¯ ¿‘«–ªÁ¡§, π›πË¡§ & Ω√∞£«• «¡∑Œ±◊∑•
'   º≠∫ÍΩ√Ω∫≈€∏Ì :
'   ∏   µ‚   ∏Ì : PRT021
'   ∏ µ‚  ∏Ò ¿˚ : π›∫∞ Ω√∞£«• √‚∑¬
'
'   ¿€   º∫   ¿œ : 2007/11/22
'   ¿€   º∫   ¿⁄ : ¿Ø«œ±’
' --------------------------------------------------------------------------------------------------------------
' --------------------------------------------------------------------------------------------------------------
'                 ºˆ     ¡§     ≥ª     øÎ
' --------------------------------------------------------------------------------------------------------------
'   1. ºˆ¡§¿œ :
'   2. ≥ª  øÎ :
'################################################################################################################

Option Explicit

Private Type tTimeTable
    '<< ∫Ò±≥ KEY VALUE >>
    TCRNM           As String
    
    '< DATA >
    DAMIM           As String
    DATA(110, 2)    As String
End Type
Private uTimeTable()    As tTimeTable


Private sini_Path   As String


Private Sub Form_Click()
    sprTeacher.Visible = False
    sprTime.Visible = False
    sprinFo.Visible = False
End Sub

Private Sub Frame1_Click()
    sprTeacher.Visible = False
    sprTime.Visible = False
    sprinFo.Visible = False
    
End Sub

Private Sub Frame2_Click()
    sprTeacher.Visible = False
    sprTime.Visible = False
    sprinFo.Visible = False

End Sub



Private Sub Form_Load()

    Dim nRow        As Long

    Me.Top = 0
    Me.Left = 0
    Me.Width = 14550
    Me.Height = 10900
    
    basFunction.RemoveContextMenu txtTeacher(0)
    
    fpYM.Text = Format(Now, "YYYYMM")
    
    Me.Tag = "LOAD"
        
        Me.Width = 14600
        Me.Height = 10755
        
        sini_Path = App.Path & "\DAESUNG.INI"       '<< ini file
        cmdTime_in.Caption = "Ω√∞£ ¡∂»∏"
        cmdinFo_in.Caption = "æ»≥ª ¡∂»∏"
        
        '>> sprTime
        cmdTime_in.Tag = ""
        With sprTime
            .Top = 480
            .Left = 5250
        
            .ShadowColor = basModule.ShadowColor2
            .ShadowDark = basModule.ShadowDark2
            .ShadowText = basModule.ShadowText2
            .GridColor = basModule.GridColor2
            .GrayAreaBackColor = basModule.GrayAreaBackColor2
            
            For nRow = 1 To .MaxRows Step 1
                .Row = nRow
                .Col = 1
                    .Text = ""
                    
                If (nRow Mod 2) = 0 Then
                    Call .SetCellBorder(.Col, .Row, .Col, .Row, 8, basModule.SectionColor1, CellBorderStyleSolid)
                End If
                
            Next nRow
            
            .ZOrder 0
            .Visible = False
        End With
        
        '>> sprinFo
        cmdinFo_in.Tag = ""
        With sprinFo
            .Top = 480
            .Left = 6870
        
            .ShadowColor = basModule.ShadowColor2
            .ShadowDark = basModule.ShadowDark2
            .ShadowText = basModule.ShadowText2
            .GridColor = basModule.GridColor2
            .GrayAreaBackColor = basModule.GrayAreaBackColor2
            
            For nRow = 1 To .MaxRows Step 1
                .Row = nRow
                .Col = 1
                    .Text = ""
                    
                Call .SetCellBorder(.Col, .Row, .Col, .Row, 8, basModule.SectionColor1, CellBorderStyleSolid)
            Next nRow
            
            .ZOrder 0
            .Visible = False
        End With
        
        
        txtTeacher(0).Text = ""
        txtTeacher(1).Text = ""
        
        txtTeacher(0).Tag = ""
        With sprTeacher
            .Top = 480
            .Left = 750
        
            .ShadowColor = basModule.ShadowColor1
            .ShadowDark = basModule.ShadowDark1
            .ShadowText = basModule.ShadowText1
            .GridColor = basModule.GridColor1
            .GrayAreaBackColor = basModule.GrayAreaBackColor1
            
            .ZOrder 0
            .MaxRows = 0
            .Visible = False
        End With
        
        VScroll1.Min = 1
        VScroll1.Max = 100
        VScroll1.SmallChange = 1
        VScroll1.LargeChange = 1
        VScroll1.Enabled = False
        
        ReDim uTimeTable(0) As tTimeTable
        
        
        
        Call Clear_Form_Control                 '< CONTROL √ ±‚»≠
        'Call Test_Print                     '< TEST

        Call init_Display_Time_and_inFo         '< Ω√∞£ π◊ æ»≥ª≥ªø™ => Ω√∞£«•∑Œ
        
        
    Me.Tag = ""
    
End Sub

'## ≈◊Ω∫∆Æ √‚∑¬
Private Sub Test_Print()

    Dim nRow        As Integer
    Dim nCol        As Integer
    
    Dim sinDex      As String
    
    On Error Resume Next
    
    For nRow = 1 To 10 Step 1
        '< Ω√∞£ >
        For nCol = 1 To 2 Step 1
            sinDex = Trim(CStr(nRow)) & Trim(CStr(nCol))
            
            LC(CInt(sinDex)).Caption = "LC" & Trim(CStr(nRow)) & Trim(CStr(nCol))
            RC(CInt(sinDex)).Caption = "RC" & Trim(CStr(nRow)) & Trim(CStr(nCol))
        Next nCol
        
        '< ∞˙∏Ò/ ∞≠ªÁ≥ªø™ test >
        For nCol = 2 To 7 Step 1
            sinDex = Trim(CStr(nRow)) & Trim(CStr(nCol))
            
            LS(CInt(sinDex)).Caption = "LS" & Trim(CStr(nRow)) & Trim(CStr(nCol))
            LT(CInt(sinDex)).Caption = "LT" & Trim(CStr(nRow)) & Trim(CStr(nCol))
            
            RS(CInt(sinDex)).Caption = "RS" & Trim(CStr(nRow)) & Trim(CStr(nCol))
            RT(CInt(sinDex)).Caption = "RT" & Trim(CStr(nRow)) & Trim(CStr(nCol))
        Next nCol
    Next nRow

End Sub


'## control √ ±‚»≠
Private Sub Clear_Form_Control()
    Dim UsrCtl      As Control
    
    '>> √ ±‚»≠
    For Each UsrCtl In Me
        With UsrCtl
            If UCase(TypeName(UsrCtl)) = "TEXTBOX" And UCase(UsrCtl.Name) <> "TXTTEACHER" Then .Text = ""
            If UCase(UsrCtl.Name) = "LC" Or _
               UCase(UsrCtl.Name) = "LS" Or _
               UCase(UsrCtl.Name) = "LT" Or _
               UCase(UsrCtl.Name) = "RC" Or _
               UCase(UsrCtl.Name) = "RS" Or _
               UCase(UsrCtl.Name) = "RT" Or _
               UCase(UsrCtl.Name) = "LHD" Or _
               UCase(UsrCtl.Name) = "RHD" Then
                .Caption = ""
            End If
            
            If UCase(TypeName(UsrCtl)) = "LINE" Then .BorderColor = &H0
            If UCase(TypeName(UsrCtl)) = "SHAPE" Then .BorderColor = &H0
        End With
    Next
End Sub


'## Ω√∞£ π◊ æ»≥ª≥ªø™ => Ω√∞£«•∑Œ
Private Sub init_Display_Time_and_inFo()
    
    Dim sGbn        As String
    Dim nRtn        As Long
    
    Dim sTmp        As String
    Dim sData       As String * 255
    
    '## Ω√∞£≥ªø™
    sGbn = "TIME"
        sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "11", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
            LC(11).Caption = sTmp:  RC(11).Caption = sTmp
        sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "12", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
            LC(12).Caption = sTmp:  RC(12).Caption = sTmp
                                                                                                                                                                      
        sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "21", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
            LC(21).Caption = sTmp:  RC(21).Caption = sTmp
        sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "22", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
            LC(22).Caption = sTmp:  RC(22).Caption = sTmp
                                                                                                                                                                      
        sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "31", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
            LC(31).Caption = sTmp:  RC(31).Caption = sTmp
        sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "32", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
            LC(32).Caption = sTmp:  RC(32).Caption = sTmp
                                                                                                                                                                      
        sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "41", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
            LC(41).Caption = sTmp:  RC(41).Caption = sTmp
        sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "42", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
            LC(42).Caption = sTmp:  RC(42).Caption = sTmp
                                                                                                                                                                      
        sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "51", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
            LC(51).Caption = sTmp:  RC(51).Caption = sTmp
        sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "52", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
            LC(52).Caption = sTmp:  RC(52).Caption = sTmp
            
        
        sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "B1", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
            LTB(1).Text = sTmp:     RTB(1).Text = sTmp
        sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "B2", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
            LTB(2).Text = sTmp:     RTB(2).Text = sTmp
            
                                                                                                                                                                      
        sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "61", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
            LC(61).Caption = sTmp:  RC(61).Caption = sTmp
        sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "62", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
            LC(62).Caption = sTmp:  RC(62).Caption = sTmp
                                                                                                                                                                      
        sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "71", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
            LC(71).Caption = sTmp:  RC(71).Caption = sTmp
        sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "72", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
            LC(72).Caption = sTmp:  RC(72).Caption = sTmp
                                                                                                                                                                      
                                                                                                                                                                      
        '>> 2008.02.25 : √ﬂ∞°
        sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "B3", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
            LTB(3).Text = sTmp:     RTB(3).Text = sTmp
        sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "B4", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
            LTB(4).Text = sTmp:     RTB(4).Text = sTmp
                                                                                                                                                                      
                                                                                                                                                                      
        sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "81", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
            LC(81).Caption = sTmp:  RC(81).Caption = sTmp
        sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "82", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
            LC(82).Caption = sTmp:  RC(82).Caption = sTmp
                                                                                                                                                                      
        sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "91", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
            LC(91).Caption = sTmp:  RC(91).Caption = sTmp
        sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "92", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
            LC(92).Caption = sTmp:  RC(92).Caption = sTmp
        
'        sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "101", "", sData, 255, sini_Path):    If nRtn > 0 Then sTmp = Left(sData, nRtn)
'            LC(101).Caption = sTmp:  RC(101).Caption = sTmp
'        sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "102", "", sData, 255, sini_Path):    If nRtn > 0 Then sTmp = Left(sData, nRtn)
'            LC(102).Caption = sTmp:  RC(102).Caption = sTmp
                        
    
    '## æ»≥ª≥ªø™
    sGbn = "INFOTCH"
        sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "LRTB1", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
            LTB(0).Text = sTmp:     RTB(0).Text = sTmp
            
        sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "LRTB2", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
            LTB(5).Text = sTmp:     RTB(5).Text = sTmp
            
        sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "INFO1", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
            ML(0).Text = sTmp:     MR(0).Text = sTmp
        sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "INFO2", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
            ML(1).Text = sTmp:     MR(1).Text = sTmp
        sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "INFO3", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
            ML(2).Text = sTmp:     MR(2).Text = sTmp
        sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "INFO4", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
            ML(3).Text = sTmp:     MR(3).Text = sTmp
        sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "INFO5", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
            ML(4).Text = sTmp:     MR(4).Text = sTmp
        sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "INFO6", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
            ML(5).Text = sTmp:     MR(5).Text = sTmp
        sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "INFO7", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
            ML(6).Text = sTmp:     MR(6).Text = sTmp
            
        sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "INFO8", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
            ML(7).Text = sTmp:     MR(7).Text = sTmp
        sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "INFO9", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
            ML(8).Text = sTmp:     MR(8).Text = sTmp
        sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "INF10", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
            ML(9).Text = sTmp:     MR(9).Text = sTmp
    
End Sub








'## Ω√∞£«• Ω√∞£ µÓ∑œ >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Private Sub cmdTime_in_Click()
    
    Dim sGbn        As String
    Dim nRtn        As Long
    
    Dim sTmp        As String
    Dim sData       As String * 255
    
    If cmdTime_in.Tag = "" Then
        cmdTime_in.Caption = "Ω√∞£ µÓ∑œ"
        
        '## µ•¿Ã≈Õ ∫“∑Øø¿±‚
        sprTime.Col = 1
        sGbn = "TIME"
            sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "11", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
                LC(11).Caption = sTmp:  RC(11).Caption = sTmp:      sprTime.Row = 1:        sprTime.Value = Replace(Trim(sTmp), ":", "", 1, -1, vbTextCompare)
            sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "12", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
                LC(12).Caption = sTmp:  RC(12).Caption = sTmp:      sprTime.Row = 2:        sprTime.Value = Replace(Trim(sTmp), ":", "", 1, -1, vbTextCompare)
                                                                                                                                                                          
            sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "21", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
                LC(21).Caption = sTmp:  RC(21).Caption = sTmp:      sprTime.Row = 3:        sprTime.Value = Replace(Trim(sTmp), ":", "", 1, -1, vbTextCompare)
            sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "22", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
                LC(22).Caption = sTmp:  RC(22).Caption = sTmp:      sprTime.Row = 4:        sprTime.Value = Replace(Trim(sTmp), ":", "", 1, -1, vbTextCompare)
                                                                                                                                                                          
            sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "31", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
                LC(31).Caption = sTmp:  RC(31).Caption = sTmp:      sprTime.Row = 5:        sprTime.Value = Replace(Trim(sTmp), ":", "", 1, -1, vbTextCompare)
            sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "32", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
                LC(32).Caption = sTmp:  RC(32).Caption = sTmp:      sprTime.Row = 6:        sprTime.Value = Replace(Trim(sTmp), ":", "", 1, -1, vbTextCompare)
                                                                                                                                                                          
            sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "41", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
                LC(41).Caption = sTmp:  RC(41).Caption = sTmp:      sprTime.Row = 7:        sprTime.Value = Replace(Trim(sTmp), ":", "", 1, -1, vbTextCompare)
            sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "42", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
                LC(42).Caption = sTmp:  RC(42).Caption = sTmp:      sprTime.Row = 8:        sprTime.Value = Replace(Trim(sTmp), ":", "", 1, -1, vbTextCompare)
                                                                                                                                                                          
            sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "51", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
                LC(51).Caption = sTmp:  RC(51).Caption = sTmp:      sprTime.Row = 11:        sprTime.Value = Replace(Trim(sTmp), ":", "", 1, -1, vbTextCompare)
            sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "52", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
                LC(52).Caption = sTmp:  RC(52).Caption = sTmp:      sprTime.Row = 12:        sprTime.Value = Replace(Trim(sTmp), ":", "", 1, -1, vbTextCompare)
                
            
            sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "B1", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
                LTB(1).Text = sTmp:     RTB(1).Text = sTmp:      sprTime.Row = 9:        sprTime.Value = Replace(Trim(sTmp), ":", "", 1, -1, vbTextCompare)
            sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "B2", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
                LTB(2).Text = sTmp:     RTB(2).Text = sTmp:      sprTime.Row = 10:        sprTime.Value = Replace(Trim(sTmp), ":", "", 1, -1, vbTextCompare)
                
                                                                                                                                                                          
            sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "61", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
                LC(61).Caption = sTmp:  RC(61).Caption = sTmp:      sprTime.Row = 13:        sprTime.Value = Replace(Trim(sTmp), ":", "", 1, -1, vbTextCompare)
            sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "62", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
                LC(62).Caption = sTmp:  RC(62).Caption = sTmp:      sprTime.Row = 14:        sprTime.Value = Replace(Trim(sTmp), ":", "", 1, -1, vbTextCompare)
                                                                                                                                                                          
            sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "71", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
                LC(71).Caption = sTmp:  RC(71).Caption = sTmp:      sprTime.Row = 15:        sprTime.Value = Replace(Trim(sTmp), ":", "", 1, -1, vbTextCompare)
            sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "72", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
                LC(72).Caption = sTmp:  RC(72).Caption = sTmp:      sprTime.Row = 16:        sprTime.Value = Replace(Trim(sTmp), ":", "", 1, -1, vbTextCompare)
                                                                                                                                                                          
            
            sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "B3", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
                LTB(3).Text = sTmp:     RTB(3).Text = sTmp:      sprTime.Row = 17:        sprTime.Value = Replace(Trim(sTmp), ":", "", 1, -1, vbTextCompare)
            sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "B4", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
                LTB(4).Text = sTmp:     RTB(4).Text = sTmp:      sprTime.Row = 18:        sprTime.Value = Replace(Trim(sTmp), ":", "", 1, -1, vbTextCompare)
            
            
            sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "81", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
                LC(81).Caption = sTmp:  RC(81).Caption = sTmp:      sprTime.Row = 19:        sprTime.Value = Replace(Trim(sTmp), ":", "", 1, -1, vbTextCompare)
            sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "82", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
                LC(82).Caption = sTmp:  RC(82).Caption = sTmp:      sprTime.Row = 20:        sprTime.Value = Replace(Trim(sTmp), ":", "", 1, -1, vbTextCompare)
                                                                                                                                                                          
            sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "91", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
                LC(91).Caption = sTmp:  RC(91).Caption = sTmp:      sprTime.Row = 21:        sprTime.Value = Replace(Trim(sTmp), ":", "", 1, -1, vbTextCompare)
            sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "92", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
                LC(92).Caption = sTmp:  RC(92).Caption = sTmp:      sprTime.Row = 22:        sprTime.Value = Replace(Trim(sTmp), ":", "", 1, -1, vbTextCompare)
            
'            sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "101", "", sData, 255, sini_Path):    If nRtn > 0 Then sTmp = Left(sData, nRtn)
'                LC(101).Caption = sTmp:  RC(101).Caption = sTmp:      sprTime.Row = 21:        sprTime.Value = Replace(Trim(sTmp), ":", "", 1, -1, vbTextCompare)
'            sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "102", "", sData, 255, sini_Path):    If nRtn > 0 Then sTmp = Left(sData, nRtn)
'                LC(102).Caption = sTmp:  RC(102).Caption = sTmp:      sprTime.Row = 22:        sprTime.Value = Replace(Trim(sTmp), ":", "", 1, -1, vbTextCompare)
                            
        sprTime.Visible = True
        cmdTime_in.Tag = "SAVE"
        
        sprTime.SetActiveCell 1, 1
        
        Exit Sub
    End If
    
    If MsgBox("Ω√∞£¿ª µÓ∑œ«œΩ√∞⁄Ω¿¥œ±Ó?", vbQuestion + vbYesNo, "Ω√∞£«• Ω√∞£µÓ∑œ") = vbNo Then
        cmdTime_in.Caption = "Ω√∞£ ¡∂»∏"
        sprTime.Visible = False
        cmdTime_in.Tag = ""
        Exit Sub
    End If
    
    If cmdTime_in.Tag = "SAVE" Then
        With sprTime
            sGbn = "TIME"
            
            .Col = 1
            '< 1±≥Ω√
                .Row = 1:   sTmp = Left(Trim(.Text), 5)
                    nRtn = basModule.WritePrivateProfileString(sGbn, "11", sTmp, sini_Path): LC(11).Caption = sTmp:   RC(11).Caption = sTmp
                .Row = 2:   sTmp = Left(Trim(.Text), 5)
                    nRtn = basModule.WritePrivateProfileString(sGbn, "12", sTmp, sini_Path): LC(12).Caption = sTmp:   RC(12).Caption = sTmp
            '< 2±≥Ω√
                .Row = 3:   sTmp = Left(Trim(.Text), 5)
                    nRtn = basModule.WritePrivateProfileString(sGbn, "21", sTmp, sini_Path): LC(21).Caption = sTmp:   RC(21).Caption = sTmp
                .Row = 4:   sTmp = Left(Trim(.Text), 5)
                    nRtn = basModule.WritePrivateProfileString(sGbn, "22", sTmp, sini_Path): LC(22).Caption = sTmp:   RC(22).Caption = sTmp
            '< 3±≥Ω√
                .Row = 5:   sTmp = Left(Trim(.Text), 5)
                    nRtn = basModule.WritePrivateProfileString(sGbn, "31", sTmp, sini_Path): LC(31).Caption = sTmp:   RC(31).Caption = sTmp
                .Row = 6:   sTmp = Left(Trim(.Text), 5)
                    nRtn = basModule.WritePrivateProfileString(sGbn, "32", sTmp, sini_Path): LC(32).Caption = sTmp:   RC(32).Caption = sTmp
            '< 4±≥Ω√
                .Row = 7:   sTmp = Left(Trim(.Text), 5)
                    nRtn = basModule.WritePrivateProfileString(sGbn, "41", sTmp, sini_Path): LC(41).Caption = sTmp:   RC(41).Caption = sTmp
                .Row = 8:   sTmp = Left(Trim(.Text), 5)
                    nRtn = basModule.WritePrivateProfileString(sGbn, "42", sTmp, sini_Path): LC(42).Caption = sTmp:   RC(42).Caption = sTmp
            '< 5±≥Ω√
                .Row = 9:   sTmp = Left(Trim(.Text), 5)
                    nRtn = basModule.WritePrivateProfileString(sGbn, "51", sTmp, sini_Path): LC(51).Caption = sTmp:   RC(51).Caption = sTmp
                .Row = 10:  sTmp = Left(Trim(.Text), 5)
                    nRtn = basModule.WritePrivateProfileString(sGbn, "52", sTmp, sini_Path): LC(52).Caption = sTmp:   RC(52).Caption = sTmp
                                                                                                                                                     
            '< break
                .Row = 11:  sTmp = Left(Trim(.Text), 5)
                    nRtn = basModule.WritePrivateProfileString(sGbn, "B1", sTmp, sini_Path): LTB(1).Text = sTmp:      RTB(1).Text = sTmp
                .Row = 12:  sTmp = Left(Trim(.Text), 5)
                    nRtn = basModule.WritePrivateProfileString(sGbn, "B2", sTmp, sini_Path): LTB(2).Text = sTmp:      RTB(2).Text = sTmp
                                                                                                                                                     
            '< 6±≥Ω√
                .Row = 13:  sTmp = Left(Trim(.Text), 5)
                    nRtn = basModule.WritePrivateProfileString(sGbn, "61", sTmp, sini_Path): LC(61).Caption = sTmp:   RC(61).Caption = sTmp
                .Row = 14:  sTmp = Left(Trim(.Text), 5)
                    nRtn = basModule.WritePrivateProfileString(sGbn, "62", sTmp, sini_Path): LC(62).Caption = sTmp:   RC(62).Caption = sTmp
            '< 7±≥Ω√
                .Row = 15:  sTmp = Left(Trim(.Text), 5)
                    nRtn = basModule.WritePrivateProfileString(sGbn, "71", sTmp, sini_Path): LC(71).Caption = sTmp:   RC(71).Caption = sTmp
                .Row = 16:  sTmp = Left(Trim(.Text), 5)
                    nRtn = basModule.WritePrivateProfileString(sGbn, "72", sTmp, sini_Path): LC(72).Caption = sTmp:   RC(72).Caption = sTmp
                    
            
            '< break
                .Row = 17:  sTmp = Left(Trim(.Text), 5)
                    nRtn = basModule.WritePrivateProfileString(sGbn, "B3", sTmp, sini_Path): LTB(5).Text = sTmp:      RTB(5).Text = sTmp
                .Row = 18:  sTmp = Left(Trim(.Text), 5)
                    nRtn = basModule.WritePrivateProfileString(sGbn, "B4", sTmp, sini_Path): LTB(5).Text = sTmp:      RTB(5).Text = sTmp
                    
            '< 8±≥Ω√
                .Row = 19:  sTmp = Left(Trim(.Text), 5)
                    nRtn = basModule.WritePrivateProfileString(sGbn, "81", sTmp, sini_Path): LC(81).Caption = sTmp:   RC(81).Caption = sTmp
                .Row = 20:  sTmp = Left(Trim(.Text), 5)
                    nRtn = basModule.WritePrivateProfileString(sGbn, "82", sTmp, sini_Path): LC(82).Caption = sTmp:   RC(82).Caption = sTmp
            '< 9±≥Ω√
                .Row = 21:  sTmp = Left(Trim(.Text), 5)
                    nRtn = basModule.WritePrivateProfileString(sGbn, "91", sTmp, sini_Path): LC(91).Caption = sTmp:   RC(91).Caption = sTmp
                .Row = 22:  sTmp = Left(Trim(.Text), 5)
                    nRtn = basModule.WritePrivateProfileString(sGbn, "92", sTmp, sini_Path): LC(92).Caption = sTmp:   RC(92).Caption = sTmp
'            '< 10±≥Ω√
'                .Row = 21:  sTmp = Left(Trim(.Text), 5)
'                    nRtn = basModule.WritePrivateProfileString(sGbn, "101", sTmp, sini_Path): LC(101).Caption = sTmp: RC(101).Caption = sTmp
'                .Row = 22:  sTmp = Left(Trim(.Text), 5)
'                    nRtn = basModule.WritePrivateProfileString(sGbn, "102", sTmp, sini_Path): LC(102).Caption = sTmp: RC(102).Caption = sTmp
        End With
        
        cmdTime_in.Tag = ""
        cmdTime_in.Caption = "Ω√∞£ ¡∂»∏"
        sprTime.Visible = False
    End If
    
End Sub

Private Sub sprTime_KeyUp(KeyCode As Integer, Shift As Integer)
    With sprTime
        Select Case KeyCode
            Case vbKeyDelete
                .Row = .ActiveRow
                .Col = 1
                    .Text = ""
        End Select
    End With
End Sub




'## Ω√∞£«• æ»≥ªµÓ∑œ  >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Private Sub cmdinFo_in_Click()
    
    Dim sGbn        As String
    Dim nRtn        As Long
    
    Dim sTmp        As String
    Dim sData       As String * 255
    
    If cmdinFo_in.Tag = "" Then
        cmdinFo_in.Caption = "≥ªøÎ µÓ∑œ"
        
        '## µ•¿Ã≈Õ ∫“∑Øø¿±‚
        sprinFo.Col = 1
        sGbn = "INFOTCH"
            sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "LRTB1", "", sData, 255, sini_Path):      If nRtn > 0 Then sTmp = Left(sData, nRtn)
                LTB(0).Text = sTmp:     RTB(0).Text = sTmp:     sprinFo.Row = 1:        sprinFo.Text = Trim(sTmp)
            
            sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "LRTB2", "", sData, 255, sini_Path):      If nRtn > 0 Then sTmp = Left(sData, nRtn)
                LTB(5).Text = sTmp:     RTB(5).Text = sTmp:     sprinFo.Row = 2:        sprinFo.Text = Trim(sTmp)
                
            
            sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "INFO1", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
                ML(0).Text = sTmp:     MR(0).Text = sTmp:     sprinFo.Row = 3:        sprinFo.Text = Trim(sTmp)
            sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "INFO2", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
                ML(1).Text = sTmp:     MR(1).Text = sTmp:     sprinFo.Row = 4:        sprinFo.Text = Trim(sTmp)
            sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "INFO3", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
                ML(2).Text = sTmp:     MR(2).Text = sTmp:     sprinFo.Row = 5:        sprinFo.Text = Trim(sTmp)
            sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "INFO4", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
                ML(3).Text = sTmp:     MR(3).Text = sTmp:     sprinFo.Row = 6:        sprinFo.Text = Trim(sTmp)
            sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "INFO5", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
                ML(4).Text = sTmp:     MR(4).Text = sTmp:     sprinFo.Row = 7:        sprinFo.Text = Trim(sTmp)
            sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "INFO6", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
                ML(5).Text = sTmp:     MR(5).Text = sTmp:     sprinFo.Row = 8:        sprinFo.Text = Trim(sTmp)
            sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "INFO7", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
                ML(6).Text = sTmp:     MR(6).Text = sTmp:     sprinFo.Row = 9:        sprinFo.Text = Trim(sTmp)
                
            sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "INFO8", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
                ML(7).Text = sTmp:     MR(7).Text = sTmp:     sprinFo.Row = 10:        sprinFo.Text = Trim(sTmp)
            sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "INFO9", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
                ML(8).Text = sTmp:     MR(8).Text = sTmp:     sprinFo.Row = 11:       sprinFo.Text = Trim(sTmp)
            sTmp = "":  nRtn = basModule.GetPrivateProfileString(sGbn, "INF10", "", sData, 255, sini_Path):     If nRtn > 0 Then sTmp = Left(sData, nRtn)
                ML(9).Text = sTmp:     MR(9).Text = sTmp:     sprinFo.Row = 12:       sprinFo.Text = Trim(sTmp)
            
        sprinFo.Visible = True
        cmdinFo_in.Tag = "SAVE"
        
        sprinFo.SetActiveCell 1, 1
        
        Exit Sub
    End If
    
    If MsgBox("æ»≥ª∏¶ µÓ∑œ«œΩ√∞⁄Ω¿¥œ±Ó?", vbQuestion + vbYesNo, "Ω√∞£«• æ»≥ªµÓ∑œ") = vbNo Then
        cmdinFo_in.Caption = "æ»≥ª ¡∂»∏"
        sprinFo.Visible = False
        cmdinFo_in.Tag = ""
        Exit Sub
    End If
    
    If cmdinFo_in.Tag = "SAVE" Then
        With sprinFo
            sGbn = "INFOTCH"
            
            .Col = 1
            '< BREAK
                .Row = 1:   sTmp = Trim(.Text)
                    nRtn = basModule.WritePrivateProfileString(sGbn, "LRTB1", sTmp, sini_Path):  LTB(5).Text = sTmp: RTB(5).Text = sTmp
                    
                .Row = 2:   sTmp = Trim(.Text)
                    nRtn = basModule.WritePrivateProfileString(sGbn, "LRTB2", sTmp, sini_Path):  LTB(5).Text = sTmp: RTB(5).Text = sTmp
                
                .Row = 3:   sTmp = Trim(.Text)
                    nRtn = basModule.WritePrivateProfileString(sGbn, "INFO1", sTmp, sini_Path): ML(0).Text = sTmp:  MR(0).Text = sTmp
                .Row = 4:   sTmp = Trim(.Text)
                    nRtn = basModule.WritePrivateProfileString(sGbn, "INFO2", sTmp, sini_Path): ML(1).Text = sTmp:  MR(1).Text = sTmp
                .Row = 5:   sTmp = Trim(.Text)
                    nRtn = basModule.WritePrivateProfileString(sGbn, "INFO3", sTmp, sini_Path): ML(2).Text = sTmp:  MR(2).Text = sTmp
                .Row = 6:   sTmp = Trim(.Text)
                    nRtn = basModule.WritePrivateProfileString(sGbn, "INFO4", sTmp, sini_Path): ML(3).Text = sTmp:  MR(3).Text = sTmp
                .Row = 7:   sTmp = Trim(.Text)
                    nRtn = basModule.WritePrivateProfileString(sGbn, "INFO5", sTmp, sini_Path): ML(4).Text = sTmp:  MR(4).Text = sTmp
                .Row = 8:   sTmp = Trim(.Text)
                    nRtn = basModule.WritePrivateProfileString(sGbn, "INFO6", sTmp, sini_Path): ML(5).Text = sTmp:  MR(5).Text = sTmp
                .Row = 9:   sTmp = Trim(.Text)
                    nRtn = basModule.WritePrivateProfileString(sGbn, "INFO7", sTmp, sini_Path): ML(6).Text = sTmp:  MR(6).Text = sTmp
                    
                .Row = 10:  sTmp = Trim(.Text)
                    nRtn = basModule.WritePrivateProfileString(sGbn, "INFO8", sTmp, sini_Path): ML(7).Text = sTmp:  MR(7).Text = sTmp
                .Row = 11:  sTmp = Trim(.Text)
                    nRtn = basModule.WritePrivateProfileString(sGbn, "INFO9", sTmp, sini_Path): ML(8).Text = sTmp:  MR(8).Text = sTmp
                .Row = 12:  sTmp = Trim(.Text)
                    nRtn = basModule.WritePrivateProfileString(sGbn, "INF10", sTmp, sini_Path): ML(9).Text = sTmp:  MR(9).Text = sTmp

        End With
        
        cmdinFo_in.Tag = ""
        cmdinFo_in.Caption = "æ»≥ª ¡∂»∏"
        sprinFo.Visible = False
    End If
    
End Sub

Private Sub sprinFo_KeyUp(KeyCode As Integer, Shift As Integer)
    With sprinFo
        Select Case KeyCode
            Case vbKeyDelete
                .Row = .ActiveRow
                .Col = 1
                    '.Text = ""
        End Select
    End With
End Sub


'#############################################################################################################################################################




'>> Ω√∞£«• ¡∂»∏
Private Sub cmdFind_Click()
    
    On Error GoTo ErrStmt
    
    ReDim uTimeTable(0) As tTimeTable
    
    cmdFind.Enabled = False
        Call Get_TimeTable_Data
        Call Disp_TimeTable_All_Data(1)
        
    cmdFind.Enabled = True
    
    MsgBox "Ω√∞£«• ¡∂»∏«œø¥Ω¿¥œ¥Ÿ.", vbInformation + vbOKOnly, "Ω√∞£«• ¡∂»∏"
    
    Exit Sub
ErrStmt:
    MsgBox "Ω√∞£«• ¡∂»∏Ω√ ø¿∑˘∞° πﬂª˝«œø¥Ω¿¥œ¥Ÿ.", vbCritical + vbOKOnly, "Ω√∞£«• ¡∂»∏"
    On Error GoTo 0

End Sub

Private Sub Get_TimeTable_Data()
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    
    Dim nLength     As Long
    Dim sStr        As String
    Dim ni          As Long
    Dim nRec        As Long
    Dim sTmp        As String
    
    Dim ninDex      As Long
    
    Dim sTeacher    As String
    Dim nArray      As Long
    
    On Error GoTo ErrStmt
    
    '>> √ ±‚»≠ -------------------------------------------------------------------
    Call Clear_Form_Control                 '< CONTROL √ ±‚»≠
    Call init_Display_Time_and_inFo         '< Ω√∞£ π◊ æ»≥ª≥ªø™ => Ω√∞£«•∑Œ
    '-----------------------------------------------------------------------------
    
    sStr = ""
    sStr = sStr & " SELECT TCRNM, SUBJNM, CLASSNM, DAMIM, IDX"
    sStr = sStr & "  FROM ( "
    sStr = sStr & "        SELECT B.TCRCD, TCRNM, "
    sStr = sStr & "               A.LSNNM AS SUBJNM,"
    sStr = sStr & "               DECODE(A.LSNNM,'πÊº€ºˆæ˜','',DECODE(A.KAEYOL,'01','¿Œ','02','¿⁄','X')||'-'||A.LSNCDNM) AS CLASSNM, "
    sStr = sStr & "               B.DAMIM , A.IDX"
    sStr = sStr & "          FROM (SELECT A.ACID, A.LSNNM, NUM AS LSNCDNM, A.KAEYOL, B.WEEKS, B.LESSON, TRIM(TO_CHAR(LESSON))||TRIM(TO_CHAR(WEEKS)) AS IDX"
    sStr = sStr & "                  FROM (SELECT ACID, TRXCD, LSNNM,"
    sStr = sStr & "                               KAEYOL, B.NUM"
    sStr = sStr & "                          FROM (SELECT ACID, TRXCD, TRXNM AS LSNNM,"
    sStr = sStr & "                                       KAEYOL"
    sStr = sStr & "                                  FROM (SELECT ACID, TRXCD, KAEYOL, TRXNM,"
    sStr = sStr & "                                               SUBSTR(SUBSTR(TRXNM,LENGTH(TRXNM)-5+1, LENGTH(TRXNM)),1,2) AS CUTA,"
    sStr = sStr & "                                               NVL(SUBSTR(SUBSTR(TRXNM,LENGTH(TRXNM)-5+1, LENGTH(TRXNM)),4,2),'AA') AS CUTB"
    sStr = sStr & "                                          FROM SDTRX01TB"
    sStr = sStr & "                                         WHERE ACID = '" & basModule.SchCD & "'"
    sStr = sStr & "                                           AND TRXCD LIKE 'P%'"
    sStr = sStr & "                                       )"
    sStr = sStr & "                                 WHERE LTRIM(CUTA,'0123456789') IS NOT NULL"
    sStr = sStr & "                                   AND LTRIM(CUTB,'0123456789') IS NOT NULL"
    sStr = sStr & "                                 ) A,"
    sStr = sStr & "                                SDTRX90TB B"
    sStr = sStr & "                          WHERE B.NO < 40"
    sStr = sStr & "                        UNION ALL"
    sStr = sStr & "                        SELECT ACID, TRXCD, SUBSTR(TRXNM,1,LENGTH(TRXNM)-5) AS LSNNM,"
    sStr = sStr & "                               KAEYOL, B.NUM"
    sStr = sStr & "                          FROM (SELECT ACID, TRXCD, KAEYOL, TRXNM, CUTA, CUTB"
    sStr = sStr & "                                  FROM (SELECT ACID, TRXCD, KAEYOL, TRXNM,"
    sStr = sStr & "                                               SUBSTR(SUBSTR(TRXNM,LENGTH(TRXNM)-5+1, LENGTH(TRXNM)),1,2) AS CUTA,"
    sStr = sStr & "                                               SUBSTR(SUBSTR(TRXNM,LENGTH(TRXNM)-5+1, LENGTH(TRXNM)),4,2) AS CUTB"
    sStr = sStr & "                                          FROM SDTRX01TB"
    sStr = sStr & "                                         WHERE ACID = '" & basModule.SchCD & "'"
    sStr = sStr & "                                           AND TRXCD LIKE 'P%'"
    sStr = sStr & "                                       )"
    sStr = sStr & "                                 WHERE LTRIM(CUTA,'0123456789') IS NULL"
    sStr = sStr & "                                   AND LTRIM(CUTB,'0123456789') IS NULL"
    sStr = sStr & "                                ) A,"
    sStr = sStr & "                               SDTRX90TB B"
    sStr = sStr & "                         WHERE B.NUM BETWEEN CUTA AND CUTB"
    sStr = sStr & "                        ) A,"
    sStr = sStr & "                       (SELECT ACID, TRXCD, KAEYOL, LESSON, WEEKS"
    sStr = sStr & "                          FROM SDTRX11TB"
    sStr = sStr & "                         WHERE ACID  = '" & basModule.SchCD & "'"
    sStr = sStr & "                           AND TRXCD LIKE 'P%'"
    sStr = sStr & "                        ) B"
    sStr = sStr & "                 WHERE A.ACID   = B.ACID"
    sStr = sStr & "                   AND A.TRXCD  = B.TRXCD"
    sStr = sStr & "                   AND A.KAEYOL = B.KAEYOL"
    sStr = sStr & "                ) A,"
    sStr = sStr & "               (SELECT A.TCRCD, A.TCRNM, A.SUBJNM,"
    sStr = sStr & "                       B.BASE_CLASS AS CLASSNM,"
    sStr = sStr & "                       B.DAMIM,"
    sStr = sStr & "                       TRIM(TO_CHAR(LESSON))||TRIM(TO_CHAR(WEEKS)) AS IDX,"
    sStr = sStr & "                       B.KAEYOL,"
    sStr = sStr & "                       B.LSNCD, B.LSNCDNM"
    sStr = sStr & "                  FROM (SELECT A.ACID, B.TCRCD, A.LSNCD, GET_LSNNM(A.ACID, A.LSNCD) AS LSNNM, A.LESSON, A.WEEKS,"
    sStr = sStr & "                               B.TCRNM, B.SUBJNM"
    sStr = sStr & "                          FROM SDTRX50TB A, SDTCR01TB B"
    sStr = sStr & "                         WHERE A.ACID   = B.ACID"
    sStr = sStr & "                           AND A.TCRCD  = B.TCRCD"
    sStr = sStr & "                           AND A.SUBJCD = B.SUBJCD"
    sStr = sStr & "                           AND A.YM     = '" & Trim(fpYM.UnFmtText) & "'"
    sStr = sStr & "                           AND A.ACID   = '" & basModule.SchCD & "'"
    sStr = sStr & "                        ) A,"
    sStr = sStr & "                       SDLSN01TB B"
    sStr = sStr & "                 WHERE A.ACID  = B.ACID"
    sStr = sStr & "                   AND A.LSNCD = B.LSNCD"
    sStr = sStr & "                   AND A.ACID  = '" & basModule.SchCD & "'"
    sStr = sStr & "                UNION ALL"
    sStr = sStr & "                SELECT A.TCRCD, A.TCRNM, A.SUBJNM,"
    sStr = sStr & "                       B.BASE_CLASS AS CLASSNM,"
    sStr = sStr & "                       B.DAMIM,"
    sStr = sStr & "                       TRIM(TO_CHAR(LESSON))||TRIM(TO_CHAR(WEEKS)) AS IDX,"
    sStr = sStr & "                       B.KAEYOL,"
    sStr = sStr & "                       B.LSNCD, B.LSNCDNM"
    sStr = sStr & "                  FROM (SELECT A.ACID, B.TCRCD, A.LSNCD, GET_LSNNM(A.ACID, A.LSNCD) AS LSNNM, A.LESSON, A.WEEKS,"
    sStr = sStr & "                               B.TCRNM, B.SUBJNM"
    sStr = sStr & "                          FROM SDTRX50TB A, SDTCR01TB B"
    sStr = sStr & "                         WHERE A.ACID   = B.ACID"
    sStr = sStr & "                           AND A.TCRCD  = B.TCRCD"
    sStr = sStr & "                           AND A.SUBJCD = B.SUBJCD"
    sStr = sStr & "                           AND A.YM     = '" & Trim(fpYM.UnFmtText) & "'"
    sStr = sStr & "                           AND A.ACID   = '" & basModule.SchCD & "'"
    sStr = sStr & "                        ) A,"
    sStr = sStr & "                       SDLSN02TB B"
    sStr = sStr & "                 WHERE A.ACID  = B.ACID"
    sStr = sStr & "                   AND A.LSNCD = B.LSNCD"
    sStr = sStr & "                   AND A.ACID  = '" & basModule.SchCD & "'"
    sStr = sStr & "                UNION ALL"
    sStr = sStr & "                SELECT B.TCRCD, B.TCRNM, B.SUBJNM,"
    sStr = sStr & "                       '' AS CLASSNM,"
    sStr = sStr & "                       '' AS DAMIM,"
    sStr = sStr & "                       TRIM(TO_CHAR(LESSON))||TRIM(TO_CHAR(WEEKS)) AS IDX,"
    sStr = sStr & "                       DECODE(LENGTH(PRT_KAEYOL),1,'0'||PRT_KAEYOL, PRT_KAEYOL) AS KAEYOL,"
    sStr = sStr & "                       '00000' AS LSNCD, PRT_LSN AS LSNCDNM"
    sStr = sStr & "                  FROM SDTRX50TB A, SDTCR01TB B"
    sStr = sStr & "                 WHERE A.ACID   = B.ACID"
    sStr = sStr & "                   AND A.TCRCD  = B.TCRCD"
    sStr = sStr & "                   AND A.SUBJCD = B.SUBJCD"
    sStr = sStr & "                   AND A.YM     = '" & Trim(fpYM.UnFmtText) & "'"
    sStr = sStr & "                   AND A.ACID   = '" & basModule.SchCD & "'"
    sStr = sStr & "                   AND A.LSNCD  = '00000'"
    sStr = sStr & "               ) B "
    sStr = sStr & "         WHERE A.KAEYOL  = B.KAEYOL"
    sStr = sStr & "           AND A.LSNCDNM = B.LSNCDNM"
    '>> ∞≠ªÁ∏Ì
    If txtTeacher(1).Text > " " Then
        sStr = sStr & "       AND TCRNM = '" & Trim(txtTeacher(1).Text) & "'"
    End If
    sStr = sStr & "        UNION ALL"
    sStr = sStr & "        SELECT TCRCD, TCRNM, SUBJNM, "
    sStr = sStr & "               DECODE(KAEYOL,'01','¿Œ','02','¿⁄','X')||'-'||LSNCDNM AS CLASSNM,"
    sStr = sStr & "               DAMIM, IDX"
    sStr = sStr & "          FROM (SELECT A.TCRCD, A.TCRNM, A.SUBJNM,"
    sStr = sStr & "                       B.BASE_CLASS AS CLASSNM,"
    sStr = sStr & "                       B.DAMIM,"
    sStr = sStr & "                       TRIM(TO_CHAR(LESSON))||TRIM(TO_CHAR(WEEKS)) AS IDX,"
    sStr = sStr & "                       B.KAEYOL,"
    sStr = sStr & "                       B.LSNCD, B.LSNCDNM"
    sStr = sStr & "                  FROM (SELECT A.ACID, B.TCRCD, A.LSNCD, GET_LSNNM(A.ACID, A.LSNCD) AS LSNNM, A.LESSON, A.WEEKS,"
    sStr = sStr & "                               B.TCRNM, B.SUBJNM"
    sStr = sStr & "                          FROM SDTRX50TB A, SDTCR01TB B"
    sStr = sStr & "                         WHERE A.ACID   = B.ACID"
    sStr = sStr & "                           AND A.TCRCD  = B.TCRCD"
    sStr = sStr & "                           AND A.SUBJCD = B.SUBJCD"
    sStr = sStr & "                           AND A.YM     = '" & Trim(fpYM.UnFmtText) & "'"
    sStr = sStr & "                           AND A.ACID   = '" & basModule.SchCD & "'"
    sStr = sStr & "                        ) A,"
    sStr = sStr & "                       SDLSN01TB B"
    sStr = sStr & "                 WHERE A.ACID  = B.ACID"
    sStr = sStr & "                   AND A.LSNCD = B.LSNCD"
    sStr = sStr & "                   AND A.ACID  = '" & basModule.SchCD & "'"
    sStr = sStr & "                UNION ALL"
    sStr = sStr & "                SELECT A.TCRCD, A.TCRNM, A.SUBJNM,"
    sStr = sStr & "                       B.BASE_CLASS AS CLASSNM,"
    sStr = sStr & "                       B.DAMIM,"
    sStr = sStr & "                       TRIM(TO_CHAR(LESSON))||TRIM(TO_CHAR(WEEKS)) AS IDX,"
    sStr = sStr & "                       B.KAEYOL,"
    sStr = sStr & "                       B.LSNCD, B.LSNCDNM"
    sStr = sStr & "                  FROM (SELECT A.ACID, B.TCRCD, A.LSNCD, GET_LSNNM(A.ACID, A.LSNCD) AS LSNNM, A.LESSON, A.WEEKS,"
    sStr = sStr & "                               B.TCRNM, B.SUBJNM"
    sStr = sStr & "                          FROM SDTRX50TB A, SDTCR01TB B"
    sStr = sStr & "                         WHERE A.ACID   = B.ACID"
    sStr = sStr & "                           AND A.TCRCD  = B.TCRCD"
    sStr = sStr & "                           AND A.SUBJCD = B.SUBJCD"
    sStr = sStr & "                           AND A.YM     = '" & Trim(fpYM.UnFmtText) & "'"
    sStr = sStr & "                           AND A.ACID   = '" & basModule.SchCD & "'"
    sStr = sStr & "                        ) A,"
    sStr = sStr & "                       SDLSN02TB B"
    sStr = sStr & "                 WHERE A.ACID  = B.ACID"
    sStr = sStr & "                   AND A.LSNCD = B.LSNCD"
    sStr = sStr & "                   AND A.ACID  = '" & basModule.SchCD & "'"
    sStr = sStr & "                UNION ALL"
    sStr = sStr & "                SELECT B.TCRCD, B.TCRNM, B.SUBJNM,"
    sStr = sStr & "                       '' AS CLASSNM,"
    sStr = sStr & "                       '' AS DAMIM,"
    sStr = sStr & "                       TRIM(TO_CHAR(LESSON))||TRIM(TO_CHAR(WEEKS)) AS IDX,"
    sStr = sStr & "                       DECODE(LENGTH(PRT_KAEYOL),1,'0'||PRT_KAEYOL, PRT_KAEYOL) AS KAEYOL,"
    sStr = sStr & "                       '00000' AS LSNCD, PRT_LSN AS LSNCDNM"
    sStr = sStr & "                  FROM SDTRX50TB A, SDTCR01TB B"
    sStr = sStr & "                 WHERE A.ACID   = B.ACID"
    sStr = sStr & "                   AND A.TCRCD  = B.TCRCD"
    sStr = sStr & "                   AND A.SUBJCD = B.SUBJCD"
    sStr = sStr & "                   AND A.YM     = '" & Trim(fpYM.UnFmtText) & "'"
    sStr = sStr & "                   AND A.ACID   = '" & basModule.SchCD & "'"
    sStr = sStr & "                   AND A.LSNCD  = '00000'"
    sStr = sStr & "               )"
    sStr = sStr & "         WHERE IDX > ' ' "
'>> ∞≠ªÁ∏Ì
    If txtTeacher(1).Text > " " Then
        sStr = sStr & "       AND TCRNM = '" & Trim(txtTeacher(1).Text) & "'"
    End If
    sStr = sStr & "       )"
    
    sStr = sStr & " ORDER BY TCRCD "
    
    
    Set DBCmd = New ADODB.Command
    Set DBRec = New ADODB.Recordset
    Set DBParam = New ADODB.Parameter
    
    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
    DBCmd.CommandText = sStr
    DBCmd.CommandType = adCmdText
    DBCmd.CommandTimeout = 30
    
    
    
''>> ∫–ø¯
'        sTmp = Trim(basModule.SchCD)
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
    
    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic«¸≈¬∑Œ ø≠∞‘µ«∏È record count∏¶ «“ ºˆ æ¯¿Ω.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop
    
    With DBRec
    
        If .RecordCount > 0 Then
            .MoveFirst
            
            
            For nRec = 1 To .RecordCount Step 1
                '< ∞≠ªÁ∏Ì √º≈© - ±‚¡ÿ¡° >
                sTeacher = "":          If IsNull(.Fields("TCRNM")) = False Then sTeacher = Trim(.Fields("TCRNM"))
                
                
                '## µ•¿Ã≈Õ √º≈© << π›, ±≥Ω√, ø‰¿œ¿Ã ∏¬æ∆æﬂ «‘.
                ninDex = 0
                If sTeacher > " " Then      '-----------------------------------------------------------------------------------------------------------------------
                    If UBound(uTimeTable) = 0 Then
                        ReDim uTimeTable(1) As tTimeTable
                        
                        ninDex = 1              ' INDEX - 1     √≥¿Ω index
                        
                    Else
                        For ni = 1 To UBound(uTimeTable) Step 1
                            If StrComp(uTimeTable(ni).TCRNM, sTeacher, vbTextCompare) = 0 Then
                               
                                ninDex = ni     ' INDEX - NI    ±‚¡∏ µÓ∑œµ» ≥ªøÎ¿∏∑Œ ≥÷¿Ω
                                
                            End If
                        Next ni
                    End If
                    
                    If ninDex = 0 Then
                        ninDex = UBound(uTimeTable) + 1
                        ReDim Preserve uTimeTable(ninDex) As tTimeTable      '<< ªı∑ŒøÓ index ª˝º∫
                    End If
                    
                    If ninDex > 0 Then
                    '>> data µÓ∑œ >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                        uTimeTable(ninDex).TCRNM = sTeacher
                        uTimeTable(ninDex).DAMIM = "":      If IsNull(.Fields("DAMIM")) = False Then uTimeTable(ninDex).DAMIM = Trim(.Fields("DAMIM"))
                        
                        nArray = 0
                        If IsNull(.Fields("IDX")) = False Then
                            nArray = CLng(.Fields("IDX"))       '< πËø≠¿ßƒ°
                            
                            If IsNull(.Fields("SUBJNM")) = False Then uTimeTable(ninDex).DATA(nArray, 1) = Trim(.Fields("SUBJNM"))
                            If IsNull(.Fields("CLASSNM")) = False Then uTimeTable(ninDex).DATA(nArray, 2) = Trim(.Fields("CLASSNM"))
                        End If
                    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
                    End If
                    
                End If      '## If sTeacher > " " Then ---------------------------------------------------------------------------------------------------------------
                
                .MoveNext
            Next nRec       '## recordcount
        End If
    End With
            
    
    '## ∏µÁ µ•¿Ã≈Õ¥¬ ¿¸ø™∫Øºˆ √≥∏Æµ«æÓ ¿÷¿Ω.
    Call Disp_TimeTable_All_Data(1)
    
    

    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    VScroll1.Enabled = True
    
    Exit Sub
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    On Error GoTo 0
    MsgBox "Ω√∞£«• ¡∂»∏Ω√ ø°∑Ø∞° πﬂª˝«œø¥Ω¿¥œ¥Ÿ.", vbCritical + vbOKOnly, "Ω√∞£«• ¡∂»∏"
End Sub


'## Ω√∞£«• µ•¿Ã≈Õ »≠∏È¿∏∑Œ view
Private Sub Disp_TimeTable_All_Data(ByVal aindex As Long)
    
    Dim UsrCtl      As Control
    Dim nRec        As Long
    
    If UBound(uTimeTable) = 0 Then
        MsgBox "Ω√∞£«•∏¶ ¡∂»∏«œººø‰.", vbExclamation + vbOKOnly, "Ω√∞£«• ¡∂»∏"
        Exit Sub
    End If
    
    If UBound(uTimeTable) < aindex Or aindex < 1 Then
        MsgBox "¥ı¿ÃªÛ ¡∂»∏«“ Ω√∞£«•∞° æ¯Ω¿¥œ¥Ÿ.", vbExclamation + vbOKOnly, "Ω√∞£«• ¡∂»∏"
        Exit Sub
    End If
    
    VScroll1.Min = 1
    VScroll1.Max = UBound(uTimeTable)
    VScroll1.Enabled = True
    
    'ainDex¿« ¿⁄∑·∏∏ ∫∏ø©¡‹
    If UBound(uTimeTable) >= aindex Then
    
        txtPage.Text = Trim(CStr(aindex)) & "/" & Trim(CStr(UBound(uTimeTable)))
    
        '>> √ ±‚»≠
        For Each UsrCtl In Me
            With UsrCtl
                If UCase(UsrCtl.Name) = "LS" Or _
                   UCase(UsrCtl.Name) = "LT" Or _
                   UCase(UsrCtl.Name) = "RS" Or _
                   UCase(UsrCtl.Name) = "RT" Or _
                   UCase(UsrCtl.Name) = "LHD" Or _
                   UCase(UsrCtl.Name) = "RHD" Then
                    .Caption = ""
                End If
            End With
        Next
    
        With uTimeTable(aindex)
        
        '// 1. header
            LHD(0).Caption = "∞≠ªÁ : " & .TCRNM:        RHD(0).Caption = "∞≠ªÁ : " & .TCRNM
            
            If StrComp(.DAMIM, "¥„¿”", vbTextCompare) = 0 Then
                LHD(1).Caption = "°º¥„¿”°Ω∞≠ªÁ":                RHD(1).Caption = "°º¥„¿”°Ω∞≠ªÁ"
            Else
                LHD(1).Caption = " ":                   RHD(1).Caption = " "
            End If
            LHD(2).Caption = " ":                       RHD(2).Caption = " "
            LHD(3).Caption = " ":                       RHD(3).Caption = " "
        
        '// 2. Ω√∞£«• π◊ æ»≥ª¥¬ ¡∂»∏Ω√ ∏µŒ √≥∏Æµ .
        
        '// 3. Ω√∞£«• ºº∫Œ≥ªø™
            For nRec = 1 To UBound(.DATA) Step 1
                If .DATA(nRec, 1) > " " Then
                    LS(nRec).Caption = .DATA(nRec, 1):      RS(nRec).Caption = .DATA(nRec, 1)
                    LT(nRec).Caption = .DATA(nRec, 2):      RT(nRec).Caption = .DATA(nRec, 2)
                    
                End If
            Next nRec
        
        End With
    End If
    
End Sub





Private Sub pReportViewer_Click()
    sprTeacher.Visible = False
    sprTime.Visible = False
    sprinFo.Visible = False
    
End Sub

Private Sub txtTeacher_Click(Index As Integer)
    '
End Sub

Private Sub txtTeacher_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF10
            sprTeacher.Visible = False
        
            txtTeacher(1).Text = ""
            Call Find_TeacherData
            
        Case vbKeyCancel
            sprTeacher.Visible = False
            sprTime.Visible = False
            sprinFo.Visible = False
            
        Case vbKeyBack
            txtTeacher(1).Text = ""
            
    End Select
End Sub

Private Sub txtTeacher_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Button
        Case vbRightButton
            sprTeacher.Visible = False
        
            txtTeacher(1).Text = ""
            Call Find_TeacherData
            
    End Select
End Sub

'>> scroll ¿Ãµø
Private Sub VScroll1_Change()
    If Me.Tag = "LOAD" Then Exit Sub
    
    VScroll1.Enabled = False
        Call Disp_TimeTable_All_Data(VScroll1.Value)
    VScroll1.Enabled = True
    
End Sub



Private Sub Find_TeacherData()
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    
    Dim nLength     As Long
    Dim sStr        As String
    Dim ni          As Long
    Dim nRec        As Long
    Dim sTmp        As String
    
    On Error GoTo ErrStmt
    
    sprTeacher.MaxRows = 0
    
    sStr = ""
    sStr = sStr & "      SELECT MAX(SISUCD) AS SISUCD, TCRNM "
    sStr = sStr & "        FROM (SELECT SISUCD, TCRNM"
    sStr = sStr & "                From SDTCR01TB "
    sStr = sStr & "               WHERE ACID = '" & Trim(basModule.SchCD) & "'"
    If Trim(txtTeacher(0).Text) = "" Then
        sStr = sStr & "             AND TCRNM  LIKE '%" & Trim(txtTeacher(0).Text) & "%'"
    End If
    sStr = sStr & "              ) "
    sStr = sStr & "       GROUP BY TCRNM"
    

    Set DBCmd = New ADODB.Command
    Set DBRec = New ADODB.Recordset
    Set DBParam = New ADODB.Parameter
    
    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
    DBCmd.CommandText = sStr
    DBCmd.CommandType = adCmdText
    DBCmd.CommandTimeout = 30
    
    
''>> ∫–ø¯
'        sTmp = Trim(basModule.SchCD)
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
    
    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic«¸≈¬∑Œ ø≠∞‘µ«∏È record count∏¶ «“ ºˆ æ¯¿Ω.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop
    
    With DBRec
    
        If .RecordCount > 0 Then
            .MoveFirst
            
            
            For nRec = 1 To .RecordCount Step 1
                sprTeacher.MaxRows = sprTeacher.MaxRows + 1
                sprTeacher.Row = sprTeacher.MaxRows
                
                sprTeacher.Col = 1
                    sTmp = " ":     If IsNull(.Fields("SISUCD")) = False Then sTmp = Trim(.Fields("SISUCD"))
                        Call basFunction.Set_SprType_Text(sprTeacher, "CENTER", "LEFT", basFunction.LenKor(sTmp), sTmp)
                sprTeacher.Col = 2
                    sTmp = " ":     If IsNull(.Fields("TCRNM")) = False Then sTmp = Trim(.Fields("TCRNM"))
                        Call basFunction.Set_SprType_Text(sprTeacher, "CENTER", "LEFT", basFunction.LenKor(sTmp), sTmp)
                
                .MoveNext
            Next nRec       '## recordcount
            
            sprTeacher.Visible = True

        End If
    End With

    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    Exit Sub
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    On Error GoTo 0
    MsgBox "π› ¡∂»∏Ω√ ø°∑Ø∞° πﬂª˝«œø¥Ω¿¥œ¥Ÿ.", vbCritical + vbOKOnly, "π› ¡∂»∏"
End Sub

'π› º±≈√
Private Sub sprTeacher_DblClick(ByVal Col As Long, ByVal Row As Long)
    If Row < 1 Then Exit Sub
    If Col < 1 Then Exit Sub
    
    With sprTeacher
        .Row = Row
        .Col = 2
            txtTeacher(1).Text = Trim(.Text)
        .Col = 2
            txtTeacher(0).Text = Trim(.Text)
    End With
    
    sprTeacher.Visible = False
End Sub





'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' √‚  ∑¬
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'## ¿¸√º√‚∑¬
Private Sub cmdPrintAll_Click()

    Dim nRec        As Long
    Dim bChk        As Boolean

    If UBound(uTimeTable) < 1 Then
        MsgBox "Ω√∞£«• √‚∑¬«“ ≥ªøÎ¿Ã æ¯Ω¿¥œ¥Ÿ.", vbExclamation + vbOKOnly, "Ω√∞£«• ¿¸√ºπ› √‚∑¬"
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
        MsgBox "¿Œº‚√Îº“«’¥œ¥Ÿ.", vbExclamation + vbOKOnly, "Ω√∞£«• ¿¸√ºπ› √‚∑¬"
        Exit Sub
    End If
    
    On Error GoTo 0
    On Error GoTo ErrStmt
    
    nRec = 0
    cmdPrint.Tag = "ALL"
    
    Do
        nRec = nRec + 1
        txtPage.Text = "1" & "/" & Trim(CStr(UBound(uTimeTable)))
        
        Call Disp_TimeTable_All_Data(nRec)                      '<< Ω√∞£«• ¡∂»∏≥ªø™ ∫∏¿Ã±‚
        
        
        
        Me.Tag = "LOAD"
            VScroll1.Value = nRec
            Call CmdPrint_Click:        DoEvents                '<< «ˆ¿Á ¡∂»∏µ» Ω√∞£«• √‚∑¬
            
        Me.Tag = ""

    Loop Until nRec = UBound(uTimeTable)
    
    cmdPrint.Tag = ""
    MsgBox "Ω√∞£«• √‚∑¬«œø¥Ω¿¥œ¥Ÿ.", vbInformation + vbOKOnly, "Ω√∞£«• ¿¸√ºπ› √‚∑¬"
    
    Exit Sub
ErrStmt:
    On Error GoTo 0
    cmdPrint.Tag = ""
    
    MsgBox "Ω√∞£«• √‚∑¬Ω√ ø°∑Ø∞° πﬂª˝«œø¥Ω¿¥œ¥Ÿ.", vbCritical + vbOKOnly, "Ω√∞£«• ¿¸√ºπ› √‚∑¬"
    
End Sub

'## «ˆ¿Á ∆‰¿Ã¡ˆ∏∏ √‚∑¬
Public Sub CmdPrint_Click()

    Dim i           As Integer
    Dim X           As Integer
    Dim Y           As Integer
    Dim pRate       As Double


    Dim bChk        As Boolean


'    If UBound(uTimeTable) < 1 Then
'        MsgBox "Ω√∞£«• √‚∑¬«“ ≥ªøÎ¿Ã æ¯Ω¿¥œ¥Ÿ.", vbExclamation + vbOKOnly, "Ω√∞£«• √‚∑¬"
'        Exit Sub
'    End If
    
    
    On Error GoTo ErrPrint
    
    '<< «ˆ¿Á ∆‰¿Ã¡ˆ∏∏ √‚∑¬«œ∏È,
    If cmdPrint.Tag = "" Then
        bChk = False
        With dlgPrint
            .CancelError = True
            .ShowPrinter
            
            bChk = True
        End With
        
ErrPrint:
        If bChk = False Then
            MsgBox "¿Œº‚√Îº“«’¥œ¥Ÿ.", vbExclamation + vbOKOnly, "Ω√∞£«• √‚∑¬"
            Exit Sub
        End If
    End If
    
    On Error GoTo 0
    On Error Resume Next        '<< ø°∑Ø∞° ≥™µµ ¡¯«‡Ω√≈¥
    
    '****************************************************************************************
    ' «¡∏∞≈Õ √‚∑¬√ ±‚»≠∏¶ «—¥Ÿ.
    ' PrintStartDoc (Width, Height, PaperSize, Orientation,TopMargin,LeftMargin
    '****************************************************************************************
    pRate = 1.15
    basFunction.PrintStartDoc pReportViewer.Width * pRate, pReportViewer.Height * pRate, vbPRPSA4, vbPRORLandscape, 1, 1


    '********************************************************************
    '  ƒ√∑∫º«¿ª ¿ÃøÎ«œø© CONTROL¿ª πËø≠∑Œ √≥∏Æ«—¥Ÿ.
    ' ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
    '  °ÿ æ∆∑°¿« º¯º≠∏¶ ¿˝¥Î∑Á πŸ≤Ÿ¡ˆ ∏ª∞Õ....   boss
    '********************************************************************
    Dim UsrCtl      As Control

    For Each UsrCtl In Me
        With UsrCtl

             If (UCase(TypeName(UsrCtl)) = "SHAPE" And UCase(UsrCtl.Name) = "FILLBOXS") Then
                '********************************************************************
                '  ≈◊µŒ∏Æ æ¯¥¬ ªÁ∞¢ π⁄Ω∫∏¶ ∏∏µÈ∞Ì ≥ª∫Œªˆ¿ª ƒ•«—¥Ÿ.
                '********************************************************************
                 Printer.DrawWidth = 0                      ' º±¿« ±Ω±‚
                 Printer.FillStyle = vbFSTransparent        ' ¥‹ªˆ
                 Printer.FillColor = basModule.WhiteColor   ' ªˆ∞• ƒ•«œ±‚
                 PrintFilledBox .Left * pRate, .Top * pRate, .Width * pRate, .Height * pRate, &HC1F1FF
             End If
             
             If (UCase(TypeName(UsrCtl)) = "SHAPE" And UCase(UsrCtl.Name) = "FILLBOXS2") Then
                '********************************************************************
                '  ≈◊µŒ∏Æ æ¯¥¬ ªÁ∞¢ π⁄Ω∫∏¶ ∏∏µÈ∞Ì ≥ª∫Œªˆ¿ª ƒ•«—¥Ÿ.
                '********************************************************************
                 Printer.DrawWidth = 0                   ' º±¿« ±Ω±‚
                 Printer.FillStyle = vbFSTransparent     ' ¥‹ªˆ
                 Printer.FillColor = &HC1F1FF            ' ªˆ∞• ƒ•«œ±‚
                 PrintFilledBox .Left * pRate, .Top * pRate, .Width * pRate, .Height * pRate, &HC1F1FF
             End If
             
        End With
    Next

    For Each UsrCtl In Me
        With UsrCtl
             If (UCase(TypeName(UsrCtl)) = "SHAPE" And UCase(UsrCtl.Name) = "BOXS") Then
                '********************************************************************
                '  line∏¶ ¿ÃøÎ«— box∏∏µÈ±‚(±‚∫ª¿˚¿∏∑Œ shape¥¬ √‚∑¬Ω√ line∏¶ ¿ÃøÎ«—¥Ÿ)
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
                         '  π⁄Ω∫/line∏¶ ±ﬂ¥¬¥Ÿ.
                         '********************************************************************
                          Printer.DrawStyle = IIf(UsrCtl.BorderStyle = 3, 2, UsrCtl.BorderStyle)
                          Printer.DrawWidth = IIf(UsrCtl.BorderStyle = 3, 1, UsrCtl.BorderWidth * 4)
                          Printer.FillStyle = vbFSTransparent
                          PrintLine .X1 * pRate, .Y1 * pRate, .X2 * pRate, .Y2 * pRate

                    Case "LABEL"
                          '********************************************************************
                          '  Label¿ª ±◊¥Î∑Œ √‚∑¬ «—¥Ÿ(º”º∫)
                          '  ¥‹) transparent¥¬ true∑Œ √≥∏Æ«œ∞Ì Ω««‡«—¥Ÿ.
                          '  SetBkMode(Printer.hdc, TRANSPARENT)πÆ¿Â¿∫ MSπˆ±◊∏¶ √≥∏Æ«œ±‚ ¿ß«‘
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
                         '  µ•¿Ã≈Õ √‚∑¬ (DATA¥¬ TEXTBOX∑Œ √≥∏Æ «—¥Ÿ.)
                         '********************************************************************
                          Select Case UCase(.Name)
                            Case "txtTeacher", "TXTPAGE"
                            
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
                          '  ¿ÃπÃ¡ˆ√‚∑¬ : picture ¿œ∞ÊøÏ ¬Ô¿Ω
                          '********************************************************************
'                          If (object.Picture <> 0) Then
'                              Printer.FontTransparent = True
'                              iBKMode = SetBkMode(Printer.hDC, OPAQUE)
'                              ' iBKMode = SetBkMode(Printer.hDC, TRANSPARENT)
'                              PrintPicture .Picture, .Left * pRate, .Top * pRate, .Width * pRate, .Height * pRate
'                          End If
             End Select
        End With
    Next

    Printer.EndDoc     ' «¡∏∞≈Õ∑Œ ∫∏≥Ω¥Ÿ

End Sub

'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

Private Sub cmdShiftLeft_Click()
    Dim sDiv()      As String
    Dim nS          As Long
    Dim nE          As Long
    
    If Trim(txtPage) > " " Then
        sDiv = Split(txtPage.Text, "/", -1, vbTextCompare)
        
        nS = CLng(sDiv(0))
        nE = CLng(sDiv(1))
        
        If (nS - 1) >= 1 Then
            VScroll1.Value = nS - 1
            VScroll1.Enabled = False
                Call Disp_TimeTable_All_Data(VScroll1.Value)
            VScroll1.Enabled = True
        End If
    End If
End Sub

Private Sub cmdShiftRight_Click()
    Dim sDiv()      As String
    Dim nS          As Long
    Dim nE          As Long
    
    If Trim(txtPage) > " " Then
        sDiv = Split(txtPage.Text, "/", -1, vbTextCompare)
        
        nS = CLng(sDiv(0))
        nE = CLng(sDiv(1))
        
        If (nS + 1) <= nE Then
            VScroll1.Value = nS + 1
            VScroll1.Enabled = False
                Call Disp_TimeTable_All_Data(VScroll1.Value)
            VScroll1.Enabled = True
        End If
    End If
End Sub





