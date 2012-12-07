VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   11100
   ClientLeft      =   4005
   ClientTop       =   2865
   ClientWidth     =   19905
   LinkTopic       =   "Form3"
   ScaleHeight     =   11100
   ScaleWidth      =   19905
   Begin VB.PictureBox pReportViewer 
      BackColor       =   &H00FFFFFF&
      Height          =   885
      Left            =   16320
      ScaleHeight     =   825
      ScaleWidth      =   1665
      TabIndex        =   338
      Top             =   2430
      Width           =   1725
   End
   Begin VB.TextBox M5N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   12780
      TabIndex        =   323
      Text            =   "M5"
      Top             =   1410
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.TextBox M4N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   10080
      TabIndex        =   322
      Text            =   "M4"
      Top             =   1410
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.TextBox M3N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   7560
      TabIndex        =   321
      Text            =   "M3"
      Top             =   1410
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.TextBox M2N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   5190
      TabIndex        =   320
      Text            =   "M2"
      Top             =   1410
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.TextBox M1N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   2940
      TabIndex        =   319
      Text            =   "M1"
      Top             =   1410
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.TextBox M1D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   1710
      TabIndex        =   318
      Text            =   "M1"
      Top             =   2100
      Width           =   1005
   End
   Begin VB.TextBox M1D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   2370
      TabIndex        =   317
      Text            =   "M1"
      Top             =   1410
      Width           =   825
   End
   Begin VB.TextBox M1D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   1710
      TabIndex        =   316
      Text            =   "M1"
      Top             =   2340
      Width           =   1005
   End
   Begin VB.TextBox M1D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   1710
      TabIndex        =   315
      Text            =   "M1"
      Top             =   2580
      Width           =   1005
   End
   Begin VB.TextBox M1D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   1710
      TabIndex        =   314
      Text            =   "M1"
      Top             =   2820
      Width           =   1005
   End
   Begin VB.TextBox M1D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   1710
      TabIndex        =   313
      Text            =   "M1"
      Top             =   3060
      Width           =   1005
   End
   Begin VB.TextBox M1D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   6
      Left            =   1710
      TabIndex        =   312
      Text            =   "M1"
      Top             =   3300
      Width           =   1005
   End
   Begin VB.TextBox M1D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   7
      Left            =   1710
      TabIndex        =   311
      Text            =   "M1"
      Top             =   3540
      Width           =   1005
   End
   Begin VB.TextBox M1D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   8
      Left            =   1710
      TabIndex        =   310
      Text            =   "M1"
      Top             =   3780
      Width           =   1005
   End
   Begin VB.TextBox M1D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   9
      Left            =   1710
      TabIndex        =   309
      Text            =   "M1"
      Top             =   4020
      Width           =   1005
   End
   Begin VB.TextBox M1D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   10
      Left            =   1710
      TabIndex        =   308
      Text            =   "M1"
      Top             =   4260
      Width           =   1005
   End
   Begin VB.TextBox M1D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   11
      Left            =   1710
      TabIndex        =   307
      Text            =   "M1"
      Top             =   4500
      Width           =   1005
   End
   Begin VB.TextBox M1D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   12
      Left            =   1710
      TabIndex        =   306
      Text            =   "M1"
      Top             =   4740
      Width           =   1005
   End
   Begin VB.TextBox M1D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   13
      Left            =   1710
      TabIndex        =   305
      Text            =   "M1"
      Top             =   4980
      Width           =   1005
   End
   Begin VB.TextBox M1D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   14
      Left            =   1710
      TabIndex        =   304
      Text            =   "M1"
      Top             =   5220
      Width           =   1005
   End
   Begin VB.TextBox M1D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   15
      Left            =   1710
      TabIndex        =   303
      Text            =   "M1"
      Top             =   5460
      Width           =   1005
   End
   Begin VB.TextBox M1D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   16
      Left            =   1710
      TabIndex        =   302
      Text            =   "M1"
      Top             =   5700
      Width           =   1005
   End
   Begin VB.TextBox M1D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   17
      Left            =   1710
      TabIndex        =   301
      Text            =   "M1"
      Top             =   5940
      Width           =   1005
   End
   Begin VB.TextBox M1D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   18
      Left            =   1710
      TabIndex        =   300
      Text            =   "M1"
      Top             =   6180
      Width           =   1005
   End
   Begin VB.TextBox M1D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   19
      Left            =   1710
      TabIndex        =   299
      Text            =   "M1"
      Top             =   6420
      Width           =   1005
   End
   Begin VB.TextBox M1D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   20
      Left            =   1710
      TabIndex        =   298
      Text            =   "M1"
      Top             =   6660
      Width           =   1005
   End
   Begin VB.TextBox M1D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   21
      Left            =   1710
      TabIndex        =   297
      Text            =   "M1"
      Top             =   6900
      Width           =   1005
   End
   Begin VB.TextBox M1D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   22
      Left            =   1710
      TabIndex        =   296
      Text            =   "M1"
      Top             =   7140
      Width           =   1005
   End
   Begin VB.TextBox M1D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   23
      Left            =   1710
      TabIndex        =   295
      Text            =   "M1"
      Top             =   7380
      Width           =   1005
   End
   Begin VB.TextBox M1D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   24
      Left            =   1710
      TabIndex        =   294
      Text            =   "M1"
      Top             =   7620
      Width           =   1005
   End
   Begin VB.TextBox M1D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   25
      Left            =   1710
      TabIndex        =   293
      Text            =   "M1"
      Top             =   7860
      Width           =   1005
   End
   Begin VB.TextBox M1D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   26
      Left            =   1710
      TabIndex        =   292
      Text            =   "M1"
      Top             =   8100
      Width           =   1005
   End
   Begin VB.TextBox M1D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   27
      Left            =   1710
      TabIndex        =   291
      Text            =   "M1"
      Top             =   8340
      Width           =   1005
   End
   Begin VB.TextBox M1D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   28
      Left            =   1710
      TabIndex        =   290
      Text            =   "M1"
      Top             =   8580
      Width           =   1005
   End
   Begin VB.TextBox M1D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   29
      Left            =   1710
      TabIndex        =   289
      Text            =   "M1"
      Top             =   8820
      Width           =   1005
   End
   Begin VB.TextBox M1D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   30
      Left            =   1710
      TabIndex        =   288
      Text            =   "M1"
      Top             =   9060
      Width           =   1005
   End
   Begin VB.TextBox M1D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   31
      Left            =   1710
      TabIndex        =   287
      Text            =   "M1"
      Top             =   9300
      Width           =   1005
   End
   Begin VB.TextBox M1N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   2940
      TabIndex        =   286
      Text            =   "M1"
      Top             =   2100
      Width           =   825
   End
   Begin VB.TextBox M1N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   2940
      TabIndex        =   285
      Text            =   "M1"
      Top             =   2340
      Width           =   825
   End
   Begin VB.TextBox M1N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   2940
      TabIndex        =   284
      Text            =   "M1"
      Top             =   2580
      Width           =   825
   End
   Begin VB.TextBox M1N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   2940
      TabIndex        =   283
      Text            =   "M1"
      Top             =   2820
      Width           =   825
   End
   Begin VB.TextBox M1N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   2940
      TabIndex        =   282
      Text            =   "M1"
      Top             =   3060
      Width           =   825
   End
   Begin VB.TextBox M1N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   6
      Left            =   2940
      TabIndex        =   281
      Text            =   "M1"
      Top             =   3300
      Width           =   825
   End
   Begin VB.TextBox M1N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   7
      Left            =   2940
      TabIndex        =   280
      Text            =   "M1"
      Top             =   3540
      Width           =   825
   End
   Begin VB.TextBox M1N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   8
      Left            =   2940
      TabIndex        =   279
      Text            =   "M1"
      Top             =   3780
      Width           =   825
   End
   Begin VB.TextBox M1N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   9
      Left            =   2940
      TabIndex        =   278
      Text            =   "M1"
      Top             =   4020
      Width           =   825
   End
   Begin VB.TextBox M1N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   10
      Left            =   2940
      TabIndex        =   277
      Text            =   "M1"
      Top             =   4260
      Width           =   825
   End
   Begin VB.TextBox M1N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   11
      Left            =   2940
      TabIndex        =   276
      Text            =   "M1"
      Top             =   4500
      Width           =   825
   End
   Begin VB.TextBox M1N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   12
      Left            =   2940
      TabIndex        =   275
      Text            =   "M1"
      Top             =   4740
      Width           =   825
   End
   Begin VB.TextBox M1N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   13
      Left            =   2940
      TabIndex        =   274
      Text            =   "M1"
      Top             =   4980
      Width           =   825
   End
   Begin VB.TextBox M1N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   14
      Left            =   2940
      TabIndex        =   273
      Text            =   "M1"
      Top             =   5220
      Width           =   825
   End
   Begin VB.TextBox M1N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   15
      Left            =   2940
      TabIndex        =   272
      Text            =   "M1"
      Top             =   5460
      Width           =   825
   End
   Begin VB.TextBox M1N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   16
      Left            =   2940
      TabIndex        =   271
      Text            =   "M1"
      Top             =   5700
      Width           =   825
   End
   Begin VB.TextBox M1N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   17
      Left            =   2940
      TabIndex        =   270
      Text            =   "M1"
      Top             =   5940
      Width           =   825
   End
   Begin VB.TextBox M1N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   18
      Left            =   2940
      TabIndex        =   269
      Text            =   "M1"
      Top             =   6180
      Width           =   825
   End
   Begin VB.TextBox M1N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   19
      Left            =   2940
      TabIndex        =   268
      Text            =   "M1"
      Top             =   6420
      Width           =   825
   End
   Begin VB.TextBox M1N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   20
      Left            =   2940
      TabIndex        =   267
      Text            =   "M1"
      Top             =   6660
      Width           =   825
   End
   Begin VB.TextBox M1N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   21
      Left            =   2940
      TabIndex        =   266
      Text            =   "M1"
      Top             =   6900
      Width           =   825
   End
   Begin VB.TextBox M1N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   22
      Left            =   2940
      TabIndex        =   265
      Text            =   "M1"
      Top             =   7140
      Width           =   825
   End
   Begin VB.TextBox M1N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   23
      Left            =   2940
      TabIndex        =   264
      Text            =   "M1"
      Top             =   7380
      Width           =   825
   End
   Begin VB.TextBox M1N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   24
      Left            =   2940
      TabIndex        =   263
      Text            =   "M1"
      Top             =   7620
      Width           =   825
   End
   Begin VB.TextBox M1N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   25
      Left            =   2940
      TabIndex        =   262
      Text            =   "M1"
      Top             =   7860
      Width           =   825
   End
   Begin VB.TextBox M1N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   26
      Left            =   2940
      TabIndex        =   261
      Text            =   "M1"
      Top             =   8100
      Width           =   825
   End
   Begin VB.TextBox M1N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   27
      Left            =   2940
      TabIndex        =   260
      Text            =   "M1"
      Top             =   8340
      Width           =   825
   End
   Begin VB.TextBox M1N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   28
      Left            =   2940
      TabIndex        =   259
      Text            =   "M1"
      Top             =   8580
      Width           =   825
   End
   Begin VB.TextBox M1N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   29
      Left            =   2940
      TabIndex        =   258
      Text            =   "M1"
      Top             =   8820
      Width           =   825
   End
   Begin VB.TextBox M1N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   30
      Left            =   2940
      TabIndex        =   257
      Text            =   "M1"
      Top             =   9060
      Width           =   825
   End
   Begin VB.TextBox M1N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   31
      Left            =   2940
      TabIndex        =   256
      Text            =   "M1"
      Top             =   9300
      Width           =   825
   End
   Begin VB.TextBox M2N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   31
      Left            =   5280
      TabIndex        =   255
      Text            =   "M2"
      Top             =   9300
      Width           =   825
   End
   Begin VB.TextBox M2N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   30
      Left            =   5280
      TabIndex        =   254
      Text            =   "M2"
      Top             =   9060
      Width           =   825
   End
   Begin VB.TextBox M2N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   29
      Left            =   5280
      TabIndex        =   253
      Text            =   "M2"
      Top             =   8820
      Width           =   825
   End
   Begin VB.TextBox M2N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   28
      Left            =   5280
      TabIndex        =   252
      Text            =   "M2"
      Top             =   8580
      Width           =   825
   End
   Begin VB.TextBox M2N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   27
      Left            =   5280
      TabIndex        =   251
      Text            =   "M2"
      Top             =   8340
      Width           =   825
   End
   Begin VB.TextBox M2N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   26
      Left            =   5280
      TabIndex        =   250
      Text            =   "M2"
      Top             =   8100
      Width           =   825
   End
   Begin VB.TextBox M2N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   25
      Left            =   5280
      TabIndex        =   249
      Text            =   "M2"
      Top             =   7860
      Width           =   825
   End
   Begin VB.TextBox M2N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   24
      Left            =   5280
      TabIndex        =   248
      Text            =   "M2"
      Top             =   7620
      Width           =   825
   End
   Begin VB.TextBox M2N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   23
      Left            =   5280
      TabIndex        =   247
      Text            =   "M2"
      Top             =   7380
      Width           =   825
   End
   Begin VB.TextBox M2N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   22
      Left            =   5280
      TabIndex        =   246
      Text            =   "M2"
      Top             =   7140
      Width           =   825
   End
   Begin VB.TextBox M2N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   21
      Left            =   5280
      TabIndex        =   245
      Text            =   "M2"
      Top             =   6900
      Width           =   825
   End
   Begin VB.TextBox M2N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   20
      Left            =   5280
      TabIndex        =   244
      Text            =   "M2"
      Top             =   6660
      Width           =   825
   End
   Begin VB.TextBox M2N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   19
      Left            =   5280
      TabIndex        =   243
      Text            =   "M2"
      Top             =   6420
      Width           =   825
   End
   Begin VB.TextBox M2N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   18
      Left            =   5280
      TabIndex        =   242
      Text            =   "M2"
      Top             =   6180
      Width           =   825
   End
   Begin VB.TextBox M2N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   17
      Left            =   5280
      TabIndex        =   241
      Text            =   "M2"
      Top             =   5940
      Width           =   825
   End
   Begin VB.TextBox M2N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   16
      Left            =   5280
      TabIndex        =   240
      Text            =   "M2"
      Top             =   5700
      Width           =   825
   End
   Begin VB.TextBox M2N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   15
      Left            =   5280
      TabIndex        =   239
      Text            =   "M2"
      Top             =   5460
      Width           =   825
   End
   Begin VB.TextBox M2N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   14
      Left            =   5280
      TabIndex        =   238
      Text            =   "M2"
      Top             =   5220
      Width           =   825
   End
   Begin VB.TextBox M2N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   13
      Left            =   5280
      TabIndex        =   237
      Text            =   "M2"
      Top             =   4980
      Width           =   825
   End
   Begin VB.TextBox M2N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   12
      Left            =   5280
      TabIndex        =   236
      Text            =   "M2"
      Top             =   4740
      Width           =   825
   End
   Begin VB.TextBox M2N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   11
      Left            =   5280
      TabIndex        =   235
      Text            =   "M2"
      Top             =   4500
      Width           =   825
   End
   Begin VB.TextBox M2N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   10
      Left            =   5280
      TabIndex        =   234
      Text            =   "M2"
      Top             =   4260
      Width           =   825
   End
   Begin VB.TextBox M2N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   9
      Left            =   5280
      TabIndex        =   233
      Text            =   "M2"
      Top             =   4020
      Width           =   825
   End
   Begin VB.TextBox M2N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   8
      Left            =   5280
      TabIndex        =   232
      Text            =   "M2"
      Top             =   3780
      Width           =   825
   End
   Begin VB.TextBox M2N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   7
      Left            =   5280
      TabIndex        =   231
      Text            =   "M2"
      Top             =   3540
      Width           =   825
   End
   Begin VB.TextBox M2N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   6
      Left            =   5280
      TabIndex        =   230
      Text            =   "M2"
      Top             =   3300
      Width           =   825
   End
   Begin VB.TextBox M2N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   5280
      TabIndex        =   229
      Text            =   "M2"
      Top             =   3060
      Width           =   825
   End
   Begin VB.TextBox M2N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   5280
      TabIndex        =   228
      Text            =   "M2"
      Top             =   2820
      Width           =   825
   End
   Begin VB.TextBox M2N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   5280
      TabIndex        =   227
      Text            =   "M2"
      Top             =   2580
      Width           =   825
   End
   Begin VB.TextBox M2N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   5280
      TabIndex        =   226
      Text            =   "M2"
      Top             =   2340
      Width           =   825
   End
   Begin VB.TextBox M2N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   5280
      TabIndex        =   225
      Text            =   "M2"
      Top             =   2100
      Width           =   825
   End
   Begin VB.TextBox M2D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   31
      Left            =   4110
      TabIndex        =   224
      Text            =   "M2"
      Top             =   9300
      Width           =   1005
   End
   Begin VB.TextBox M2D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   30
      Left            =   4110
      TabIndex        =   223
      Text            =   "M2"
      Top             =   9060
      Width           =   1005
   End
   Begin VB.TextBox M2D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   29
      Left            =   4110
      TabIndex        =   222
      Text            =   "M2"
      Top             =   8820
      Width           =   1005
   End
   Begin VB.TextBox M2D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   28
      Left            =   4110
      TabIndex        =   221
      Text            =   "M2"
      Top             =   8580
      Width           =   1005
   End
   Begin VB.TextBox M2D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   27
      Left            =   4110
      TabIndex        =   220
      Text            =   "M2"
      Top             =   8340
      Width           =   1005
   End
   Begin VB.TextBox M2D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   26
      Left            =   4110
      TabIndex        =   219
      Text            =   "M2"
      Top             =   8100
      Width           =   1005
   End
   Begin VB.TextBox M2D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   25
      Left            =   4110
      TabIndex        =   218
      Text            =   "M2"
      Top             =   7860
      Width           =   1005
   End
   Begin VB.TextBox M2D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   24
      Left            =   4110
      TabIndex        =   217
      Text            =   "M2"
      Top             =   7620
      Width           =   1005
   End
   Begin VB.TextBox M2D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   23
      Left            =   4110
      TabIndex        =   216
      Text            =   "M2"
      Top             =   7380
      Width           =   1005
   End
   Begin VB.TextBox M2D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   22
      Left            =   4110
      TabIndex        =   215
      Text            =   "M2"
      Top             =   7140
      Width           =   1005
   End
   Begin VB.TextBox M2D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   21
      Left            =   4110
      TabIndex        =   214
      Text            =   "M2"
      Top             =   6900
      Width           =   1005
   End
   Begin VB.TextBox M2D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   20
      Left            =   4110
      TabIndex        =   213
      Text            =   "M2"
      Top             =   6660
      Width           =   1005
   End
   Begin VB.TextBox M2D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   19
      Left            =   4110
      TabIndex        =   212
      Text            =   "M2"
      Top             =   6420
      Width           =   1005
   End
   Begin VB.TextBox M2D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   18
      Left            =   4110
      TabIndex        =   211
      Text            =   "M2"
      Top             =   6180
      Width           =   1005
   End
   Begin VB.TextBox M2D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   17
      Left            =   4110
      TabIndex        =   210
      Text            =   "M2"
      Top             =   5940
      Width           =   1005
   End
   Begin VB.TextBox M2D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   16
      Left            =   4110
      TabIndex        =   209
      Text            =   "M2"
      Top             =   5700
      Width           =   1005
   End
   Begin VB.TextBox M2D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   15
      Left            =   4110
      TabIndex        =   208
      Text            =   "M2"
      Top             =   5460
      Width           =   1005
   End
   Begin VB.TextBox M2D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   14
      Left            =   4110
      TabIndex        =   207
      Text            =   "M2"
      Top             =   5220
      Width           =   1005
   End
   Begin VB.TextBox M2D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   13
      Left            =   4110
      TabIndex        =   206
      Text            =   "M2"
      Top             =   4980
      Width           =   1005
   End
   Begin VB.TextBox M2D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   12
      Left            =   4110
      TabIndex        =   205
      Text            =   "M2"
      Top             =   4740
      Width           =   1005
   End
   Begin VB.TextBox M2D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   11
      Left            =   4110
      TabIndex        =   204
      Text            =   "M2"
      Top             =   4500
      Width           =   1005
   End
   Begin VB.TextBox M2D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   10
      Left            =   4110
      TabIndex        =   203
      Text            =   "M2"
      Top             =   4260
      Width           =   1005
   End
   Begin VB.TextBox M2D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   9
      Left            =   4110
      TabIndex        =   202
      Text            =   "M2"
      Top             =   4020
      Width           =   1005
   End
   Begin VB.TextBox M2D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   8
      Left            =   4110
      TabIndex        =   201
      Text            =   "M2"
      Top             =   3780
      Width           =   1005
   End
   Begin VB.TextBox M2D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   7
      Left            =   4110
      TabIndex        =   200
      Text            =   "M2"
      Top             =   3540
      Width           =   1005
   End
   Begin VB.TextBox M2D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   6
      Left            =   4110
      TabIndex        =   199
      Text            =   "M2"
      Top             =   3300
      Width           =   1005
   End
   Begin VB.TextBox M2D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   4110
      TabIndex        =   198
      Text            =   "M2"
      Top             =   3060
      Width           =   1005
   End
   Begin VB.TextBox M2D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   4110
      TabIndex        =   197
      Text            =   "M2"
      Top             =   2820
      Width           =   1005
   End
   Begin VB.TextBox M2D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   4110
      TabIndex        =   196
      Text            =   "M2"
      Top             =   2580
      Width           =   1005
   End
   Begin VB.TextBox M2D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   4110
      TabIndex        =   195
      Text            =   "M2"
      Top             =   2340
      Width           =   1005
   End
   Begin VB.TextBox M2D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   4620
      TabIndex        =   194
      Text            =   "M2"
      Top             =   1410
      Width           =   825
   End
   Begin VB.TextBox M2D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   4110
      TabIndex        =   193
      Text            =   "M2"
      Top             =   2100
      Width           =   1005
   End
   Begin VB.TextBox M3N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   31
      Left            =   7710
      TabIndex        =   192
      Text            =   "M3"
      Top             =   9300
      Width           =   825
   End
   Begin VB.TextBox M3N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   30
      Left            =   7710
      TabIndex        =   191
      Text            =   "M3"
      Top             =   9060
      Width           =   825
   End
   Begin VB.TextBox M3N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   29
      Left            =   7710
      TabIndex        =   190
      Text            =   "M3"
      Top             =   8820
      Width           =   825
   End
   Begin VB.TextBox M3N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   28
      Left            =   7710
      TabIndex        =   189
      Text            =   "M3"
      Top             =   8580
      Width           =   825
   End
   Begin VB.TextBox M3N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   27
      Left            =   7710
      TabIndex        =   188
      Text            =   "M3"
      Top             =   8340
      Width           =   825
   End
   Begin VB.TextBox M3N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   26
      Left            =   7710
      TabIndex        =   187
      Text            =   "M3"
      Top             =   8100
      Width           =   825
   End
   Begin VB.TextBox M3N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   25
      Left            =   7710
      TabIndex        =   186
      Text            =   "M3"
      Top             =   7860
      Width           =   825
   End
   Begin VB.TextBox M3N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   24
      Left            =   7710
      TabIndex        =   185
      Text            =   "M3"
      Top             =   7620
      Width           =   825
   End
   Begin VB.TextBox M3N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   23
      Left            =   7710
      TabIndex        =   184
      Text            =   "M3"
      Top             =   7380
      Width           =   825
   End
   Begin VB.TextBox M3N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   22
      Left            =   7710
      TabIndex        =   183
      Text            =   "M3"
      Top             =   7140
      Width           =   825
   End
   Begin VB.TextBox M3N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   21
      Left            =   7710
      TabIndex        =   182
      Text            =   "M3"
      Top             =   6900
      Width           =   825
   End
   Begin VB.TextBox M3N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   20
      Left            =   7710
      TabIndex        =   181
      Text            =   "M3"
      Top             =   6660
      Width           =   825
   End
   Begin VB.TextBox M3N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   19
      Left            =   7710
      TabIndex        =   180
      Text            =   "M3"
      Top             =   6420
      Width           =   825
   End
   Begin VB.TextBox M3N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   18
      Left            =   7710
      TabIndex        =   179
      Text            =   "M3"
      Top             =   6180
      Width           =   825
   End
   Begin VB.TextBox M3N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   17
      Left            =   7710
      TabIndex        =   178
      Text            =   "M3"
      Top             =   5940
      Width           =   825
   End
   Begin VB.TextBox M3N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   16
      Left            =   7710
      TabIndex        =   177
      Text            =   "M3"
      Top             =   5700
      Width           =   825
   End
   Begin VB.TextBox M3N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   15
      Left            =   7710
      TabIndex        =   176
      Text            =   "M3"
      Top             =   5460
      Width           =   825
   End
   Begin VB.TextBox M3N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   14
      Left            =   7710
      TabIndex        =   175
      Text            =   "M3"
      Top             =   5220
      Width           =   825
   End
   Begin VB.TextBox M3N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   13
      Left            =   7710
      TabIndex        =   174
      Text            =   "M3"
      Top             =   4980
      Width           =   825
   End
   Begin VB.TextBox M3N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   12
      Left            =   7710
      TabIndex        =   173
      Text            =   "M3"
      Top             =   4740
      Width           =   825
   End
   Begin VB.TextBox M3N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   11
      Left            =   7710
      TabIndex        =   172
      Text            =   "M3"
      Top             =   4500
      Width           =   825
   End
   Begin VB.TextBox M3N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   10
      Left            =   7710
      TabIndex        =   171
      Text            =   "M3"
      Top             =   4260
      Width           =   825
   End
   Begin VB.TextBox M3N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   9
      Left            =   7710
      TabIndex        =   170
      Text            =   "M3"
      Top             =   4020
      Width           =   825
   End
   Begin VB.TextBox M3N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   8
      Left            =   7710
      TabIndex        =   169
      Text            =   "M3"
      Top             =   3780
      Width           =   825
   End
   Begin VB.TextBox M3N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   7
      Left            =   7710
      TabIndex        =   168
      Text            =   "M3"
      Top             =   3540
      Width           =   825
   End
   Begin VB.TextBox M3N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   6
      Left            =   7710
      TabIndex        =   167
      Text            =   "M3"
      Top             =   3300
      Width           =   825
   End
   Begin VB.TextBox M3N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   7710
      TabIndex        =   166
      Text            =   "M3"
      Top             =   3060
      Width           =   825
   End
   Begin VB.TextBox M3N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   7710
      TabIndex        =   165
      Text            =   "M3"
      Top             =   2820
      Width           =   825
   End
   Begin VB.TextBox M3N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   7710
      TabIndex        =   164
      Text            =   "M3"
      Top             =   2580
      Width           =   825
   End
   Begin VB.TextBox M3N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   7710
      TabIndex        =   163
      Text            =   "M3"
      Top             =   2340
      Width           =   825
   End
   Begin VB.TextBox M3N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   7710
      TabIndex        =   162
      Text            =   "M3"
      Top             =   2100
      Width           =   825
   End
   Begin VB.TextBox M3D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   31
      Left            =   6510
      TabIndex        =   161
      Text            =   "M3"
      Top             =   9300
      Width           =   1005
   End
   Begin VB.TextBox M3D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   30
      Left            =   6510
      TabIndex        =   160
      Text            =   "M3"
      Top             =   9060
      Width           =   1005
   End
   Begin VB.TextBox M3D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   29
      Left            =   6510
      TabIndex        =   159
      Text            =   "M3"
      Top             =   8820
      Width           =   1005
   End
   Begin VB.TextBox M3D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   28
      Left            =   6510
      TabIndex        =   158
      Text            =   "M3"
      Top             =   8580
      Width           =   1005
   End
   Begin VB.TextBox M3D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   27
      Left            =   6510
      TabIndex        =   157
      Text            =   "M3"
      Top             =   8340
      Width           =   1005
   End
   Begin VB.TextBox M3D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   26
      Left            =   6510
      TabIndex        =   156
      Text            =   "M3"
      Top             =   8100
      Width           =   1005
   End
   Begin VB.TextBox M3D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   25
      Left            =   6510
      TabIndex        =   155
      Text            =   "M3"
      Top             =   7860
      Width           =   1005
   End
   Begin VB.TextBox M3D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   24
      Left            =   6510
      TabIndex        =   154
      Text            =   "M3"
      Top             =   7620
      Width           =   1005
   End
   Begin VB.TextBox M3D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   23
      Left            =   6510
      TabIndex        =   153
      Text            =   "M3"
      Top             =   7380
      Width           =   1005
   End
   Begin VB.TextBox M3D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   22
      Left            =   6510
      TabIndex        =   152
      Text            =   "M3"
      Top             =   7140
      Width           =   1005
   End
   Begin VB.TextBox M3D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   21
      Left            =   6510
      TabIndex        =   151
      Text            =   "M3"
      Top             =   6900
      Width           =   1005
   End
   Begin VB.TextBox M3D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   20
      Left            =   6510
      TabIndex        =   150
      Text            =   "M3"
      Top             =   6660
      Width           =   1005
   End
   Begin VB.TextBox M3D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   19
      Left            =   6510
      TabIndex        =   149
      Text            =   "M3"
      Top             =   6420
      Width           =   1005
   End
   Begin VB.TextBox M3D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   18
      Left            =   6510
      TabIndex        =   148
      Text            =   "M3"
      Top             =   6180
      Width           =   1005
   End
   Begin VB.TextBox M3D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   17
      Left            =   6510
      TabIndex        =   147
      Text            =   "M3"
      Top             =   5940
      Width           =   1005
   End
   Begin VB.TextBox M3D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   16
      Left            =   6510
      TabIndex        =   146
      Text            =   "M3"
      Top             =   5700
      Width           =   1005
   End
   Begin VB.TextBox M3D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   15
      Left            =   6510
      TabIndex        =   145
      Text            =   "M3"
      Top             =   5460
      Width           =   1005
   End
   Begin VB.TextBox M3D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   14
      Left            =   6510
      TabIndex        =   144
      Text            =   "M3"
      Top             =   5220
      Width           =   1005
   End
   Begin VB.TextBox M3D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   13
      Left            =   6510
      TabIndex        =   143
      Text            =   "M3"
      Top             =   4980
      Width           =   1005
   End
   Begin VB.TextBox M3D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   12
      Left            =   6510
      TabIndex        =   142
      Text            =   "M3"
      Top             =   4740
      Width           =   1005
   End
   Begin VB.TextBox M3D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   11
      Left            =   6510
      TabIndex        =   141
      Text            =   "M3"
      Top             =   4500
      Width           =   1005
   End
   Begin VB.TextBox M3D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   10
      Left            =   6510
      TabIndex        =   140
      Text            =   "M3"
      Top             =   4260
      Width           =   1005
   End
   Begin VB.TextBox M3D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   9
      Left            =   6510
      TabIndex        =   139
      Text            =   "M3"
      Top             =   4020
      Width           =   1005
   End
   Begin VB.TextBox M3D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   8
      Left            =   6510
      TabIndex        =   138
      Text            =   "M3"
      Top             =   3780
      Width           =   1005
   End
   Begin VB.TextBox M3D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   7
      Left            =   6510
      TabIndex        =   137
      Text            =   "M3"
      Top             =   3540
      Width           =   1005
   End
   Begin VB.TextBox M3D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   6
      Left            =   6510
      TabIndex        =   136
      Text            =   "M3"
      Top             =   3300
      Width           =   1005
   End
   Begin VB.TextBox M3D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   6510
      TabIndex        =   135
      Text            =   "M3"
      Top             =   3060
      Width           =   1005
   End
   Begin VB.TextBox M3D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   6510
      TabIndex        =   134
      Text            =   "M3"
      Top             =   2820
      Width           =   1005
   End
   Begin VB.TextBox M3D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   6510
      TabIndex        =   133
      Text            =   "M3"
      Top             =   2580
      Width           =   1005
   End
   Begin VB.TextBox M3D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   6510
      TabIndex        =   132
      Text            =   "M3"
      Top             =   2340
      Width           =   1005
   End
   Begin VB.TextBox M3D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   6990
      TabIndex        =   131
      Text            =   "M3"
      Top             =   1410
      Width           =   825
   End
   Begin VB.TextBox M3D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   6510
      TabIndex        =   130
      Text            =   "M3"
      Top             =   2100
      Width           =   1005
   End
   Begin VB.TextBox M4N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   31
      Left            =   10230
      TabIndex        =   129
      Text            =   "M4"
      Top             =   9300
      Width           =   825
   End
   Begin VB.TextBox M4N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   30
      Left            =   10230
      TabIndex        =   128
      Text            =   "M4"
      Top             =   9060
      Width           =   825
   End
   Begin VB.TextBox M4N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   29
      Left            =   10230
      TabIndex        =   127
      Text            =   "M4"
      Top             =   8820
      Width           =   825
   End
   Begin VB.TextBox M4N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   28
      Left            =   10230
      TabIndex        =   126
      Text            =   "M4"
      Top             =   8580
      Width           =   825
   End
   Begin VB.TextBox M4N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   27
      Left            =   10230
      TabIndex        =   125
      Text            =   "M4"
      Top             =   8340
      Width           =   825
   End
   Begin VB.TextBox M4N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   26
      Left            =   10230
      TabIndex        =   124
      Text            =   "M4"
      Top             =   8100
      Width           =   825
   End
   Begin VB.TextBox M4N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   25
      Left            =   10230
      TabIndex        =   123
      Text            =   "M4"
      Top             =   7860
      Width           =   825
   End
   Begin VB.TextBox M4N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   24
      Left            =   10230
      TabIndex        =   122
      Text            =   "M4"
      Top             =   7620
      Width           =   825
   End
   Begin VB.TextBox M4N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   23
      Left            =   10230
      TabIndex        =   121
      Text            =   "M4"
      Top             =   7380
      Width           =   825
   End
   Begin VB.TextBox M4N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   22
      Left            =   10230
      TabIndex        =   120
      Text            =   "M4"
      Top             =   7140
      Width           =   825
   End
   Begin VB.TextBox M4N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   21
      Left            =   10230
      TabIndex        =   119
      Text            =   "M4"
      Top             =   6900
      Width           =   825
   End
   Begin VB.TextBox M4N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   20
      Left            =   10230
      TabIndex        =   118
      Text            =   "M4"
      Top             =   6660
      Width           =   825
   End
   Begin VB.TextBox M4N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   19
      Left            =   10230
      TabIndex        =   117
      Text            =   "M4"
      Top             =   6420
      Width           =   825
   End
   Begin VB.TextBox M4N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   18
      Left            =   10230
      TabIndex        =   116
      Text            =   "M4"
      Top             =   6180
      Width           =   825
   End
   Begin VB.TextBox M4N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   17
      Left            =   10230
      TabIndex        =   115
      Text            =   "M4"
      Top             =   5940
      Width           =   825
   End
   Begin VB.TextBox M4N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   16
      Left            =   10230
      TabIndex        =   114
      Text            =   "M4"
      Top             =   5700
      Width           =   825
   End
   Begin VB.TextBox M4N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   15
      Left            =   10230
      TabIndex        =   113
      Text            =   "M4"
      Top             =   5460
      Width           =   825
   End
   Begin VB.TextBox M4N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   14
      Left            =   10230
      TabIndex        =   112
      Text            =   "M4"
      Top             =   5220
      Width           =   825
   End
   Begin VB.TextBox M4N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   13
      Left            =   10230
      TabIndex        =   111
      Text            =   "M4"
      Top             =   4980
      Width           =   825
   End
   Begin VB.TextBox M4N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   12
      Left            =   10230
      TabIndex        =   110
      Text            =   "M4"
      Top             =   4740
      Width           =   825
   End
   Begin VB.TextBox M4N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   11
      Left            =   10230
      TabIndex        =   109
      Text            =   "M4"
      Top             =   4500
      Width           =   825
   End
   Begin VB.TextBox M4N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   10
      Left            =   10230
      TabIndex        =   108
      Text            =   "M4"
      Top             =   4260
      Width           =   825
   End
   Begin VB.TextBox M4N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   9
      Left            =   10230
      TabIndex        =   107
      Text            =   "M4"
      Top             =   4020
      Width           =   825
   End
   Begin VB.TextBox M4N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   8
      Left            =   10230
      TabIndex        =   106
      Text            =   "M4"
      Top             =   3780
      Width           =   825
   End
   Begin VB.TextBox M4N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   7
      Left            =   10230
      TabIndex        =   105
      Text            =   "M4"
      Top             =   3540
      Width           =   825
   End
   Begin VB.TextBox M4N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   6
      Left            =   10230
      TabIndex        =   104
      Text            =   "M4"
      Top             =   3300
      Width           =   825
   End
   Begin VB.TextBox M4N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   10230
      TabIndex        =   103
      Text            =   "M4"
      Top             =   3060
      Width           =   825
   End
   Begin VB.TextBox M4N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   10230
      TabIndex        =   102
      Text            =   "M4"
      Top             =   2820
      Width           =   825
   End
   Begin VB.TextBox M4N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   10230
      TabIndex        =   101
      Text            =   "M4"
      Top             =   2580
      Width           =   825
   End
   Begin VB.TextBox M4N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   10230
      TabIndex        =   100
      Text            =   "M4"
      Top             =   2340
      Width           =   825
   End
   Begin VB.TextBox M4N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   10230
      TabIndex        =   99
      Text            =   "M4"
      Top             =   2100
      Width           =   825
   End
   Begin VB.TextBox M4D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   31
      Left            =   8970
      TabIndex        =   98
      Text            =   "M4"
      Top             =   9300
      Width           =   1005
   End
   Begin VB.TextBox M4D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   30
      Left            =   8970
      TabIndex        =   97
      Text            =   "M4"
      Top             =   9060
      Width           =   1005
   End
   Begin VB.TextBox M4D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   29
      Left            =   8970
      TabIndex        =   96
      Text            =   "M4"
      Top             =   8820
      Width           =   1005
   End
   Begin VB.TextBox M4D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   28
      Left            =   8970
      TabIndex        =   95
      Text            =   "M4"
      Top             =   8580
      Width           =   1005
   End
   Begin VB.TextBox M4D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   27
      Left            =   8970
      TabIndex        =   94
      Text            =   "M4"
      Top             =   8340
      Width           =   1005
   End
   Begin VB.TextBox M4D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   26
      Left            =   8970
      TabIndex        =   93
      Text            =   "M4"
      Top             =   8100
      Width           =   1005
   End
   Begin VB.TextBox M4D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   25
      Left            =   8970
      TabIndex        =   92
      Text            =   "M4"
      Top             =   7860
      Width           =   1005
   End
   Begin VB.TextBox M4D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   24
      Left            =   8970
      TabIndex        =   91
      Text            =   "M4"
      Top             =   7620
      Width           =   1005
   End
   Begin VB.TextBox M4D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   23
      Left            =   8970
      TabIndex        =   90
      Text            =   "M4"
      Top             =   7380
      Width           =   1005
   End
   Begin VB.TextBox M4D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   22
      Left            =   8970
      TabIndex        =   89
      Text            =   "M4"
      Top             =   7140
      Width           =   1005
   End
   Begin VB.TextBox M4D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   21
      Left            =   8970
      TabIndex        =   88
      Text            =   "M4"
      Top             =   6900
      Width           =   1005
   End
   Begin VB.TextBox M4D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   20
      Left            =   8970
      TabIndex        =   87
      Text            =   "M4"
      Top             =   6660
      Width           =   1005
   End
   Begin VB.TextBox M4D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   19
      Left            =   8970
      TabIndex        =   86
      Text            =   "M4"
      Top             =   6420
      Width           =   1005
   End
   Begin VB.TextBox M4D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   18
      Left            =   8970
      TabIndex        =   85
      Text            =   "M4"
      Top             =   6180
      Width           =   1005
   End
   Begin VB.TextBox M4D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   17
      Left            =   8970
      TabIndex        =   84
      Text            =   "M4"
      Top             =   5940
      Width           =   1005
   End
   Begin VB.TextBox M4D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   16
      Left            =   8970
      TabIndex        =   83
      Text            =   "M4"
      Top             =   5700
      Width           =   1005
   End
   Begin VB.TextBox M4D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   15
      Left            =   8970
      TabIndex        =   82
      Text            =   "M4"
      Top             =   5460
      Width           =   1005
   End
   Begin VB.TextBox M4D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   14
      Left            =   8970
      TabIndex        =   81
      Text            =   "M4"
      Top             =   5220
      Width           =   1005
   End
   Begin VB.TextBox M4D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   13
      Left            =   8970
      TabIndex        =   80
      Text            =   "M4"
      Top             =   4980
      Width           =   1005
   End
   Begin VB.TextBox M4D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   12
      Left            =   8970
      TabIndex        =   79
      Text            =   "M4"
      Top             =   4740
      Width           =   1005
   End
   Begin VB.TextBox M4D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   11
      Left            =   8970
      TabIndex        =   78
      Text            =   "M4"
      Top             =   4500
      Width           =   1005
   End
   Begin VB.TextBox M4D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   10
      Left            =   8970
      TabIndex        =   77
      Text            =   "M4"
      Top             =   4260
      Width           =   1005
   End
   Begin VB.TextBox M4D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   9
      Left            =   8970
      TabIndex        =   76
      Text            =   "M4"
      Top             =   4020
      Width           =   1005
   End
   Begin VB.TextBox M4D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   8
      Left            =   8970
      TabIndex        =   75
      Text            =   "M4"
      Top             =   3780
      Width           =   1005
   End
   Begin VB.TextBox M4D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   7
      Left            =   8970
      TabIndex        =   74
      Text            =   "M4"
      Top             =   3540
      Width           =   1005
   End
   Begin VB.TextBox M4D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   6
      Left            =   8970
      TabIndex        =   73
      Text            =   "M4"
      Top             =   3300
      Width           =   1005
   End
   Begin VB.TextBox M4D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   8970
      TabIndex        =   72
      Text            =   "M4"
      Top             =   3060
      Width           =   1005
   End
   Begin VB.TextBox M4D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   8970
      TabIndex        =   71
      Text            =   "M4"
      Top             =   2820
      Width           =   1005
   End
   Begin VB.TextBox M4D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   8970
      TabIndex        =   70
      Text            =   "M4"
      Top             =   2580
      Width           =   1005
   End
   Begin VB.TextBox M4D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   8970
      TabIndex        =   69
      Text            =   "M4"
      Top             =   2340
      Width           =   1005
   End
   Begin VB.TextBox M4D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   9510
      TabIndex        =   68
      Text            =   "M4"
      Top             =   1410
      Width           =   825
   End
   Begin VB.TextBox M4D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   8970
      TabIndex        =   67
      Text            =   "M4"
      Top             =   2100
      Width           =   1005
   End
   Begin VB.TextBox M5N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   31
      Left            =   12780
      TabIndex        =   66
      Text            =   "M5"
      Top             =   9300
      Width           =   825
   End
   Begin VB.TextBox M5N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   30
      Left            =   12780
      TabIndex        =   65
      Text            =   "M5"
      Top             =   9060
      Width           =   825
   End
   Begin VB.TextBox M5N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   29
      Left            =   12780
      TabIndex        =   64
      Text            =   "M5"
      Top             =   8820
      Width           =   825
   End
   Begin VB.TextBox M5N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   28
      Left            =   12780
      TabIndex        =   63
      Text            =   "M5"
      Top             =   8580
      Width           =   825
   End
   Begin VB.TextBox M5N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   27
      Left            =   12780
      TabIndex        =   62
      Text            =   "M5"
      Top             =   8340
      Width           =   825
   End
   Begin VB.TextBox M5N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   26
      Left            =   12780
      TabIndex        =   61
      Text            =   "M5"
      Top             =   8100
      Width           =   825
   End
   Begin VB.TextBox M5N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   25
      Left            =   12780
      TabIndex        =   60
      Text            =   "M5"
      Top             =   7860
      Width           =   825
   End
   Begin VB.TextBox M5N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   24
      Left            =   12780
      TabIndex        =   59
      Text            =   "M5"
      Top             =   7620
      Width           =   825
   End
   Begin VB.TextBox M5N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   23
      Left            =   12780
      TabIndex        =   58
      Text            =   "M5"
      Top             =   7380
      Width           =   825
   End
   Begin VB.TextBox M5N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   22
      Left            =   12780
      TabIndex        =   57
      Text            =   "M5"
      Top             =   7140
      Width           =   825
   End
   Begin VB.TextBox M5N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   21
      Left            =   12780
      TabIndex        =   56
      Text            =   "M5"
      Top             =   6900
      Width           =   825
   End
   Begin VB.TextBox M5N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   20
      Left            =   12780
      TabIndex        =   55
      Text            =   "M5"
      Top             =   6660
      Width           =   825
   End
   Begin VB.TextBox M5N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   19
      Left            =   12780
      TabIndex        =   54
      Text            =   "M5"
      Top             =   6420
      Width           =   825
   End
   Begin VB.TextBox M5N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   18
      Left            =   12780
      TabIndex        =   53
      Text            =   "M5"
      Top             =   6180
      Width           =   825
   End
   Begin VB.TextBox M5N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   17
      Left            =   12780
      TabIndex        =   52
      Text            =   "M5"
      Top             =   5940
      Width           =   825
   End
   Begin VB.TextBox M5N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   16
      Left            =   12780
      TabIndex        =   51
      Text            =   "M5"
      Top             =   5700
      Width           =   825
   End
   Begin VB.TextBox M5N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   15
      Left            =   12780
      TabIndex        =   50
      Text            =   "M5"
      Top             =   5460
      Width           =   825
   End
   Begin VB.TextBox M5N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   14
      Left            =   12780
      TabIndex        =   49
      Text            =   "M5"
      Top             =   5220
      Width           =   825
   End
   Begin VB.TextBox M5N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   13
      Left            =   12780
      TabIndex        =   48
      Text            =   "M5"
      Top             =   4980
      Width           =   825
   End
   Begin VB.TextBox M5N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   12
      Left            =   12780
      TabIndex        =   47
      Text            =   "M5"
      Top             =   4740
      Width           =   825
   End
   Begin VB.TextBox M5N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   11
      Left            =   12780
      TabIndex        =   46
      Text            =   "M5"
      Top             =   4500
      Width           =   825
   End
   Begin VB.TextBox M5N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   10
      Left            =   12780
      TabIndex        =   45
      Text            =   "M5"
      Top             =   4260
      Width           =   825
   End
   Begin VB.TextBox M5N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   9
      Left            =   12780
      TabIndex        =   44
      Text            =   "M5"
      Top             =   4020
      Width           =   825
   End
   Begin VB.TextBox M5N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   8
      Left            =   12780
      TabIndex        =   43
      Text            =   "M5"
      Top             =   3780
      Width           =   825
   End
   Begin VB.TextBox M5N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   7
      Left            =   12780
      TabIndex        =   42
      Text            =   "M5"
      Top             =   3540
      Width           =   825
   End
   Begin VB.TextBox M5N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   6
      Left            =   12780
      TabIndex        =   41
      Text            =   "M5"
      Top             =   3300
      Width           =   825
   End
   Begin VB.TextBox M5N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   12780
      TabIndex        =   40
      Text            =   "M5"
      Top             =   3060
      Width           =   825
   End
   Begin VB.TextBox M5N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   12780
      TabIndex        =   39
      Text            =   "M5"
      Top             =   2820
      Width           =   825
   End
   Begin VB.TextBox M5N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   12780
      TabIndex        =   38
      Text            =   "M5"
      Top             =   2580
      Width           =   825
   End
   Begin VB.TextBox M5N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   12780
      TabIndex        =   37
      Text            =   "M5"
      Top             =   2340
      Width           =   825
   End
   Begin VB.TextBox M5N 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   12780
      TabIndex        =   36
      Text            =   "M5"
      Top             =   2100
      Width           =   825
   End
   Begin VB.TextBox M5D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   31
      Left            =   11550
      TabIndex        =   35
      Text            =   "M5"
      Top             =   9300
      Width           =   1005
   End
   Begin VB.TextBox M5D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   30
      Left            =   11550
      TabIndex        =   34
      Text            =   "M5"
      Top             =   9060
      Width           =   1005
   End
   Begin VB.TextBox M5D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   29
      Left            =   11550
      TabIndex        =   33
      Text            =   "M5"
      Top             =   8820
      Width           =   1005
   End
   Begin VB.TextBox M5D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   28
      Left            =   11550
      TabIndex        =   32
      Text            =   "M5"
      Top             =   8580
      Width           =   1005
   End
   Begin VB.TextBox M5D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   27
      Left            =   11550
      TabIndex        =   31
      Text            =   "M5"
      Top             =   8340
      Width           =   1005
   End
   Begin VB.TextBox M5D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   26
      Left            =   11550
      TabIndex        =   30
      Text            =   "M5"
      Top             =   8100
      Width           =   1005
   End
   Begin VB.TextBox M5D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   25
      Left            =   11550
      TabIndex        =   29
      Text            =   "M5"
      Top             =   7860
      Width           =   1005
   End
   Begin VB.TextBox M5D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   24
      Left            =   11550
      TabIndex        =   28
      Text            =   "M5"
      Top             =   7620
      Width           =   1005
   End
   Begin VB.TextBox M5D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   23
      Left            =   11550
      TabIndex        =   27
      Text            =   "M5"
      Top             =   7380
      Width           =   1005
   End
   Begin VB.TextBox M5D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   22
      Left            =   11550
      TabIndex        =   26
      Text            =   "M5"
      Top             =   7140
      Width           =   1005
   End
   Begin VB.TextBox M5D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   21
      Left            =   11550
      TabIndex        =   25
      Text            =   "M5"
      Top             =   6900
      Width           =   1005
   End
   Begin VB.TextBox M5D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   20
      Left            =   11550
      TabIndex        =   24
      Text            =   "M5"
      Top             =   6660
      Width           =   1005
   End
   Begin VB.TextBox M5D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   19
      Left            =   11550
      TabIndex        =   23
      Text            =   "M5"
      Top             =   6420
      Width           =   1005
   End
   Begin VB.TextBox M5D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   18
      Left            =   11550
      TabIndex        =   22
      Text            =   "M5"
      Top             =   6180
      Width           =   1005
   End
   Begin VB.TextBox M5D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   17
      Left            =   11550
      TabIndex        =   21
      Text            =   "M5"
      Top             =   5940
      Width           =   1005
   End
   Begin VB.TextBox M5D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   16
      Left            =   11550
      TabIndex        =   20
      Text            =   "M5"
      Top             =   5700
      Width           =   1005
   End
   Begin VB.TextBox M5D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   15
      Left            =   11550
      TabIndex        =   19
      Text            =   "M5"
      Top             =   5460
      Width           =   1005
   End
   Begin VB.TextBox M5D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   14
      Left            =   11550
      TabIndex        =   18
      Text            =   "M5"
      Top             =   5220
      Width           =   1005
   End
   Begin VB.TextBox M5D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   13
      Left            =   11550
      TabIndex        =   17
      Text            =   "M5"
      Top             =   4980
      Width           =   1005
   End
   Begin VB.TextBox M5D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   12
      Left            =   11550
      TabIndex        =   16
      Text            =   "M5"
      Top             =   4740
      Width           =   1005
   End
   Begin VB.TextBox M5D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   11
      Left            =   11550
      TabIndex        =   15
      Text            =   "M5"
      Top             =   4500
      Width           =   1005
   End
   Begin VB.TextBox M5D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   10
      Left            =   11550
      TabIndex        =   14
      Text            =   "M5"
      Top             =   4260
      Width           =   1005
   End
   Begin VB.TextBox M5D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   9
      Left            =   11550
      TabIndex        =   13
      Text            =   "M5"
      Top             =   4020
      Width           =   1005
   End
   Begin VB.TextBox M5D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   8
      Left            =   11550
      TabIndex        =   12
      Text            =   "M5"
      Top             =   3780
      Width           =   1005
   End
   Begin VB.TextBox M5D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   7
      Left            =   11550
      TabIndex        =   11
      Text            =   "M5"
      Top             =   3540
      Width           =   1005
   End
   Begin VB.TextBox M5D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   6
      Left            =   11550
      TabIndex        =   10
      Text            =   "M5"
      Top             =   3300
      Width           =   1005
   End
   Begin VB.TextBox M5D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   11550
      TabIndex        =   9
      Text            =   "M5"
      Top             =   3060
      Width           =   1005
   End
   Begin VB.TextBox M5D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   11550
      TabIndex        =   8
      Text            =   "M5"
      Top             =   2820
      Width           =   1005
   End
   Begin VB.TextBox M5D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   11550
      TabIndex        =   7
      Text            =   "M5"
      Top             =   2580
      Width           =   1005
   End
   Begin VB.TextBox M5D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   11550
      TabIndex        =   6
      Text            =   "M5"
      Top             =   2340
      Width           =   1005
   End
   Begin VB.TextBox M5D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   12210
      TabIndex        =   5
      Text            =   "M5"
      Top             =   1410
      Width           =   825
   End
   Begin VB.TextBox M5D 
      Appearance      =   0  'Æò¸é
      BorderStyle     =   0  '¾øÀ½
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   11550
      TabIndex        =   4
      Text            =   "M5"
      Top             =   2100
      Width           =   1005
   End
   Begin VB.TextBox txtStdCD1 
      BorderStyle     =   0  '¾øÀ½
      Height          =   285
      Left            =   3120
      TabIndex        =   3
      Text            =   "txtStdCD1"
      Top             =   750
      Width           =   615
   End
   Begin VB.TextBox txtStdNM1 
      BorderStyle     =   0  '¾øÀ½
      Height          =   285
      Left            =   5280
      TabIndex        =   2
      Text            =   "txtStdNM1"
      Top             =   750
      Width           =   1005
   End
   Begin VB.TextBox txtBan 
      BorderStyle     =   0  '¾øÀ½
      Height          =   285
      Left            =   1950
      TabIndex        =   1
      Text            =   "txtBan"
      Top             =   750
      Width           =   495
   End
   Begin VB.TextBox txtGaeyol 
      BorderStyle     =   0  '¾øÀ½
      Height          =   285
      Left            =   960
      TabIndex        =   0
      Text            =   "txtGaeyol"
      Top             =   750
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Åõ¸í
      Caption         =   "ÀÏ ÀÏ ´Ü¾î½ÃÇè ¼ºÀûÇ¥"
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   18
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   4350
      TabIndex        =   337
      Top             =   0
      Width           =   4605
   End
   Begin VB.Shape Boxs 
      BorderColor     =   &H00FF0000&
      Height          =   555
      Index           =   2
      Left            =   690
      Top             =   570
      Width           =   5925
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Åõ¸í
      Caption         =   "¹Ý"
      Height          =   210
      Left            =   2520
      TabIndex        =   336
      Top             =   750
      Width           =   495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Åõ¸í
      Caption         =   "¹ø"
      Height          =   210
      Left            =   3840
      TabIndex        =   335
      Top             =   750
      Width           =   495
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Åõ¸í
      Caption         =   "¼º¸í"
      Height          =   210
      Left            =   4740
      TabIndex        =   334
      Top             =   750
      Width           =   495
   End
   Begin VB.Line Lines 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   28
      X1              =   720
      X2              =   13860
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Lines 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   0
      X1              =   720
      X2              =   720
      Y1              =   1320
      Y2              =   9630
   End
   Begin VB.Line Lines 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   1
      X1              =   720
      X2              =   13830
      Y1              =   1950
      Y2              =   1950
   End
   Begin VB.Line Lines 
      BorderColor     =   &H00FF0000&
      Index           =   16
      X1              =   1530
      X2              =   1530
      Y1              =   1290
      Y2              =   9600
   End
   Begin VB.Line Lines 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   2
      X1              =   720
      X2              =   13830
      Y1              =   9630
      Y2              =   9630
   End
   Begin VB.Line Lines 
      BorderColor     =   &H00FF0000&
      Index           =   3
      X1              =   3930
      X2              =   3930
      Y1              =   1290
      Y2              =   9600
   End
   Begin VB.Line Lines 
      BorderColor     =   &H00FF0000&
      Index           =   4
      X1              =   6300
      X2              =   6300
      Y1              =   1290
      Y2              =   9600
   End
   Begin VB.Line Lines 
      BorderColor     =   &H00FF0000&
      Index           =   5
      X1              =   8760
      X2              =   8760
      Y1              =   1290
      Y2              =   9600
   End
   Begin VB.Line Lines 
      BorderColor     =   &H00FF0000&
      Index           =   6
      X1              =   11310
      X2              =   11310
      Y1              =   1290
      Y2              =   9600
   End
   Begin VB.Line Lines 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   7
      X1              =   13860
      X2              =   13860
      Y1              =   1290
      Y2              =   9630
   End
   Begin VB.Line Lines 
      BorderColor     =   &H00FF0000&
      Index           =   8
      X1              =   2820
      X2              =   2820
      Y1              =   1710
      Y2              =   9600
   End
   Begin VB.Line Lines 
      BorderColor     =   &H00FF0000&
      Index           =   9
      X1              =   5130
      X2              =   5130
      Y1              =   1680
      Y2              =   9600
   End
   Begin VB.Line Lines 
      BorderColor     =   &H00FF0000&
      Index           =   10
      X1              =   7530
      X2              =   7530
      Y1              =   1710
      Y2              =   9600
   End
   Begin VB.Line Lines 
      BorderColor     =   &H00FF0000&
      Index           =   11
      X1              =   9990
      X2              =   9990
      Y1              =   1680
      Y2              =   9600
   End
   Begin VB.Line Lines 
      BorderColor     =   &H00FF0000&
      Index           =   12
      X1              =   12570
      X2              =   12570
      Y1              =   1710
      Y2              =   9630
   End
   Begin VB.Line Lines 
      BorderColor     =   &H00FF0000&
      Index           =   13
      X1              =   1530
      X2              =   13830
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line Lines 
      BorderColor     =   &H00FF0000&
      Index           =   14
      X1              =   720
      X2              =   1590
      Y1              =   1350
      Y2              =   1950
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Åõ¸í
      Caption         =   "³¯Â¥"
      Height          =   210
      Left            =   2070
      TabIndex        =   333
      Top             =   1710
      Width           =   495
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Åõ¸í
      Caption         =   "Á¡¼ö"
      Height          =   210
      Left            =   3120
      TabIndex        =   332
      Top             =   1710
      Width           =   495
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Åõ¸í
      Caption         =   "³¯Â¥"
      Height          =   210
      Left            =   4440
      TabIndex        =   331
      Top             =   1710
      Width           =   495
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Åõ¸í
      Caption         =   "Á¡¼ö"
      Height          =   210
      Left            =   5490
      TabIndex        =   330
      Top             =   1710
      Width           =   495
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Åõ¸í
      Caption         =   "³¯Â¥"
      Height          =   210
      Left            =   6720
      TabIndex        =   329
      Top             =   1710
      Width           =   495
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Åõ¸í
      Caption         =   "Á¡¼ö"
      Height          =   210
      Left            =   7860
      TabIndex        =   328
      Top             =   1710
      Width           =   495
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Åõ¸í
      Caption         =   "³¯Â¥"
      Height          =   210
      Left            =   9210
      TabIndex        =   327
      Top             =   1710
      Width           =   495
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Åõ¸í
      Caption         =   "Á¡¼ö"
      Height          =   210
      Left            =   10410
      TabIndex        =   326
      Top             =   1710
      Width           =   495
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Åõ¸í
      Caption         =   "³¯Â¥"
      Height          =   210
      Left            =   11760
      TabIndex        =   325
      Top             =   1710
      Width           =   495
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Åõ¸í
      Caption         =   "Á¡¼ö"
      Height          =   210
      Left            =   12930
      TabIndex        =   324
      Top             =   1710
      Width           =   495
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

