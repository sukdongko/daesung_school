VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form TMR012 
   Caption         =   "시간표 만들기 >> 강사 및 시수넣기"
   ClientHeight    =   9090
   ClientLeft      =   2340
   ClientTop       =   2325
   ClientWidth     =   15645
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9090
   ScaleWidth      =   15645
   Begin FPSpread.vaSpread sprTcr 
      Height          =   2655
      Left            =   90
      TabIndex        =   7
      Top             =   10440
      Width           =   2955
      _Version        =   393216
      _ExtentX        =   5212
      _ExtentY        =   4683
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
      SpreadDesigner  =   "TMR012.frx":0000
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0082C8E8&
      BorderStyle     =   0  '없음
      Caption         =   "Frame1"
      Height          =   735
      Left            =   60
      TabIndex        =   8
      Top             =   30
      Width           =   15435
      Begin VB.Frame Frame2 
         BackColor       =   &H00D2EAF5&
         BorderStyle     =   0  '없음
         Caption         =   "Frame2"
         Height          =   675
         Left            =   30
         TabIndex        =   9
         Top             =   30
         Width           =   15375
         Begin VB.CommandButton cmdTcr 
            Caption         =   "강사내역"
            Height          =   495
            Left            =   13950
            TabIndex        =   15
            Top             =   75
            Width           =   1335
         End
         Begin VB.CommandButton cmdSaveTmr 
            Caption         =   "시수내역 등록 (&S)"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   11100
            TabIndex        =   14
            Top             =   75
            Width           =   2325
         End
         Begin VB.ComboBox cboKaeyol 
            Height          =   300
            Left            =   2340
            Style           =   2  '드롭다운 목록
            TabIndex        =   1
            Top             =   165
            Width           =   1065
         End
         Begin VB.ComboBox cboSubjGbn 
            Height          =   300
            Left            =   7080
            Style           =   2  '드롭다운 목록
            TabIndex        =   4
            Top             =   165
            Width           =   1305
         End
         Begin VB.TextBox txtTcrNM 
            Height          =   300
            IMEMode         =   10  '한글 
            Left            =   4740
            TabIndex        =   3
            Text            =   "txtTcrNM"
            Top             =   150
            Width           =   1455
         End
         Begin VB.ComboBox cboTcrGbn 
            Height          =   300
            Left            =   9300
            Style           =   2  '드롭다운 목록
            TabIndex        =   5
            Top             =   165
            Width           =   1305
         End
         Begin VB.CommandButton cmdFindTmr 
            Caption         =   "조 회 (&F)"
            Height          =   495
            Left            =   210
            TabIndex        =   0
            Top             =   75
            Width           =   1515
         End
         Begin EditLib.fpMask fpTcrCD 
            Height          =   300
            Left            =   4110
            TabIndex        =   2
            Top             =   150
            Width           =   615
            _Version        =   196608
            _ExtentX        =   1085
            _ExtentY        =   529
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
            Mask            =   "999"
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
         Begin VB.Label Label2 
            BackStyle       =   0  '투명
            Caption         =   "                 삭제는 [ DEL ] 키 가능"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   7530
            TabIndex        =   16
            Top             =   510
            Width           =   3555
         End
         Begin VB.Label Label4 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "계열"
            Height          =   210
            Left            =   1320
            TabIndex        =   13
            Top             =   210
            Width           =   975
         End
         Begin VB.Label Label1 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "과목구분"
            Height          =   210
            Left            =   6030
            TabIndex        =   12
            Top             =   210
            Width           =   975
         End
         Begin VB.Label Label26 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "강사명"
            Height          =   210
            Left            =   3270
            TabIndex        =   11
            Top             =   210
            Width           =   765
         End
         Begin VB.Label Label3 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "강사구분"
            Height          =   210
            Left            =   8250
            TabIndex        =   10
            Top             =   210
            Width           =   975
         End
      End
   End
   Begin FPSpread.vaSpread sprTmr 
      Height          =   8505
      Left            =   30
      TabIndex        =   6
      Top             =   810
      Width           =   15465
      _Version        =   393216
      _ExtentX        =   27279
      _ExtentY        =   15002
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
      ProcessTab      =   -1  'True
      SpreadDesigner  =   "TMR012.frx":1832
   End
End
Attribute VB_Name = "TMR012"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################################################
'   시 스 템  명 : 대성학원 입학사정, 반배정 & 시간표 프로그램
'   서브시스템명 :
'   모   듈   명 : TRM012
'   모 듈  목 적 : 강사 및 시수넣기
'
'   작   성   일 : 2007/12/27
'   작   성   자 : 유하균
' --------------------------------------------------------------------------------------------------------------
' --------------------------------------------------------------------------------------------------------------
'                 수     정     내     용
' --------------------------------------------------------------------------------------------------------------
'   1. 수정일 :
'   2. 내  용 :
'################################################################################################################

Option Explicit

Private Type tSisu_Data
    ACID        As String
    TCRCD       As String
    SUBJCD      As String
    
    LSNCD       As String
    SISU        As Long
End Type
Private uSisu_Data()    As tSisu_Data

Private Sub cmdTcr_Click()
    Load TMR011
    TMR011.Show
    TMR011.ZOrder 0
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload TMR011
    
End Sub

Private Sub Form_Load()
    Me.Move 0, 0, 15700, 9980
    
    basFunction.RemoveContextMenu txtTcrNM
    'basFunction.RemoveContextMenu fpTcrCD
    
    With sprTmr
        .ShadowColor = basModule.ShadowColor2
        .ShadowDark = basModule.ShadowDark2
        .ShadowText = basModule.ShadowText2
        .GridColor = basModule.GridColor2
        .GrayAreaBackColor = basModule.GrayAreaBackColor2
        
        .MaxRows = 0
    End With
    
    With sprTcr
        .ShadowColor = basModule.ShadowColor1
        .ShadowDark = basModule.ShadowDark1
        .ShadowText = basModule.ShadowText1
        .GridColor = basModule.GridColor1
        .GrayAreaBackColor = basModule.GrayAreaBackColor1
        
        .MaxRows = 0
        .ZOrder 0
        .Left = 6690
        .Top = 210
    End With
    
    With cboSubjGbn
        .Clear
        
        .AddItem "전체" & Space(50) & "ALL"
        .AddItem "언어" & Space(50) & "10"
        .AddItem "수리" & Space(50) & "20"
        .AddItem "외국어" & Space(50) & "30"
        .AddItem "사탐" & Space(50) & "40"
        .AddItem "과탐" & Space(50) & "50"
        
        .ListIndex = 0
    End With
    
    With cboTcrGbn
        .Clear
        
        .AddItem "없음" & Space(50) & "ALL"
        .AddItem "강남 출강" & Space(50) & "10"
        .AddItem "송파 출강" & Space(50) & "20"
        
        .ListIndex = 0
    End With
    
    With cboKaeyol
        .Clear
        .AddItem "전체" & Space(30) & "ALL"
        .AddItem "인문" & Space(30) & "01"
        .AddItem "자연" & Space(30) & "02"
        '.AddItem "예체" & Space(30) & "03"
        
        .ListIndex = 0
    End With
    
    ReDim uSisu_Data(0) As tSisu_Data
    
    Me.Tag = "LOAD"
        Call initData
    
    Me.Tag = ""

End Sub

Private Sub initData()
    fpTcrCD.Text = ""
    txtTcrNM.Text = ""
    
    sprTcr.Visible = False
    
    With sprTmr
        .MaxCols = 0
        .MaxRows = 0
    End With
    
End Sub



'>> 강사조회
Private Sub fpTcrCD_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    
    Dim nLength     As Long
    Dim sStr        As String
    Dim ni          As Integer
    Dim nRec        As Long
    
    Dim sTmp        As String
    
    On Error GoTo ErrStmt
    
    Select Case KeyCode
        Case vbKeyEscape
            sprTcr.Visible = False
            Exit Sub
            
        Case vbKeyReturn
            sStr = ""
            sStr = sStr & "  SELECT ACID, TCRCD, SUBJCD, SUBJGBN, TCRGBN, TCRNM, SUBJNM, TCR_CL"
            sStr = sStr & "    From SDTCR01TB"
            sStr = sStr & "   WHERE ACID   = '" & Trim(basModule.SchCD) & "'"
            sStr = sStr & "     AND TCRCD  LIKE '" & Trim(fpTcrCD.UnFmtText) & "%'"
                
        Case vbKeyF10
            sStr = ""
            sStr = sStr & "  SELECT ACID, TCRCD, SUBJCD, SUBJGBN, TCRGBN, TCRNM, SUBJNM, TCR_CL"
            sStr = sStr & "    From SDTCR01TB"
            sStr = sStr & "   WHERE ACID   = '" & Trim(basModule.SchCD) & "'"
            If Trim(fpTcrCD.UnFmtText) > " " Then
                sStr = sStr & "     AND TCRNM  LIKE '" & Trim(txtTcrNM.Text) & "%'"
            End If
            
        Case Else
            Exit Sub
    End Select
    
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
    
    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic형태로 열게되면 record count를 할 수 없음.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop
    
    With DBRec
        If .RecordCount = 1 Then
            .MoveFirst
            
            fpTcrCD.Text = "":      If IsNull(.Fields("TCRCD")) = False Then fpTcrCD.Text = Trim(.Fields("TCRCD"))
            txtTcrNM.Text = " ":    If IsNull(.Fields("TCRNM")) = False Then txtTcrNM.Text = Trim(.Fields("TCRNM"))
        ElseIf .RecordCount > 1 Then
            sprTcr.Visible = True
            sprTcr.MaxRows = 0
            
            .MoveFirst
            
            For nRec = 1 To .RecordCount Step 1
                sprTcr.MaxRows = sprTcr.MaxRows + 1
                sprTcr.Row = sprTcr.MaxRows
                
                sprTcr.Col = 1:     sTmp = "":      If IsNull(.Fields("TCRCD")) = False Then sTmp = Trim(.Fields("TCRCD"))
                    Call basFunction.Set_SprType_Text(sprTcr, "CENTER", "CENTER", basFunction.LenKor(sTmp), sTmp)
                sprTcr.Col = 2:     sTmp = "":      If IsNull(.Fields("TCRNM")) = False Then sTmp = Trim(.Fields("TCRNM"))
                    Call basFunction.Set_SprType_Text(sprTcr, "CENTER", "LEFT", basFunction.LenKor(sTmp), sTmp)
                    
                .MoveNext
            Next nRec
        End If
    End With
    
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    fpTcrCD.SetFocus
            
    Exit Sub
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    On Error GoTo 0
    MsgBox "강사 조회시 에러가 발생하였습니다.", vbCritical + vbOKOnly, "강사조회"
End Sub

Private Sub fpTcrCD_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    
    Dim nLength     As Long
    Dim sStr        As String
    Dim ni          As Integer
    Dim nRec        As Long
    
    Dim sTmp        As String
    
    On Error GoTo ErrStmt
    
    Select Case Button
        Case vbRightButton
            sStr = ""
            sStr = sStr & "  SELECT ACID, TCRCD, SUBJCD, SUBJGBN, TCRGBN, TCRNM, SUBJNM, TCR_CL"
            sStr = sStr & "    From SDTCR01TB"
            sStr = sStr & "   WHERE ACID   = '" & Trim(basModule.SchCD) & "'"
            If Trim(fpTcrCD.UnFmtText) > " " Then
                sStr = sStr & "     AND TCRNM  LIKE '" & Trim(txtTcrNM.Text) & "%'"
            End If
            
        Case Else
            Exit Sub
    End Select
    
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
    
    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic형태로 열게되면 record count를 할 수 없음.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop
    
    With DBRec
        If .RecordCount = 1 Then
            .MoveFirst
            
            fpTcrCD.Text = "":      If IsNull(.Fields("TCRCD")) = False Then fpTcrCD.Text = Trim(.Fields("TCRCD"))
            txtTcrNM.Text = " ":    If IsNull(.Fields("TCRNM")) = False Then txtTcrNM.Text = Trim(.Fields("TCRNM"))
        ElseIf .RecordCount > 1 Then
            sprTcr.Visible = True
            sprTcr.MaxRows = 0
            
            .MoveFirst
            
            For nRec = 1 To .RecordCount Step 1
                sprTcr.MaxRows = sprTcr.MaxRows + 1
                sprTcr.Row = sprTcr.MaxRows
                
                sprTcr.Col = 1:     sTmp = "":      If IsNull(.Fields("TCRCD")) = False Then sTmp = Trim(.Fields("TCRCD"))
                    Call basFunction.Set_SprType_Text(sprTcr, "CENTER", "CENTER", basFunction.LenKor(sTmp), sTmp)
                sprTcr.Col = 2:     sTmp = "":      If IsNull(.Fields("TCRNM")) = False Then sTmp = Trim(.Fields("TCRNM"))
                    Call basFunction.Set_SprType_Text(sprTcr, "CENTER", "LEFT", basFunction.LenKor(sTmp), sTmp)
                    
                .MoveNext
            Next nRec
        End If
    End With
    
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    fpTcrCD.SetFocus
            
    Exit Sub
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    On Error GoTo 0
    MsgBox "강사 조회시 에러가 발생하였습니다.", vbCritical + vbOKOnly, "강사조회"
    
End Sub



Private Sub txtTcrNM_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    
    Dim nLength     As Long
    Dim sStr        As String
    Dim ni          As Integer
    Dim nRec        As Long
    
    Dim sTmp        As String
    
    On Error GoTo ErrStmt
    
    Select Case KeyCode
        Case vbKeyBack
            fpTcrCD.Text = ""
            Exit Sub
            
        Case vbKeyEscape
            sprTcr.Visible = False
            Exit Sub
                
        Case vbKeyReturn
            sStr = ""
            sStr = sStr & "  SELECT ACID, TCRCD, SUBJCD, SUBJGBN, TCRGBN, TCRNM, SUBJNM, TCR_CL"
            sStr = sStr & "    From SDTCR01TB"
            sStr = sStr & "   WHERE ACID   = '" & Trim(basModule.SchCD) & "'"
            sStr = sStr & "     AND TCRNM  LIKE '" & Trim(txtTcrNM.Text) & "%'"
                
        Case vbKeyF10
            sStr = ""
            sStr = sStr & "  SELECT ACID, TCRCD, SUBJCD, SUBJGBN, TCRGBN, TCRNM, SUBJNM, TCR_CL"
            sStr = sStr & "    From SDTCR01TB"
            sStr = sStr & "   WHERE ACID   = '" & Trim(basModule.SchCD) & "'"
            If Trim(txtTcrNM.Text) > " " Then
                sStr = sStr & "     AND TCRNM  LIKE '" & Trim(txtTcrNM.Text) & "%'"
            End If
        
        Case Else
            Exit Sub
            
    End Select
    
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
    
    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic형태로 열게되면 record count를 할 수 없음.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop
    
    With DBRec
        If .RecordCount = 1 Then
            .MoveFirst
            
            fpTcrCD.Text = "":      If IsNull(.Fields("TCRCD")) = False Then fpTcrCD.Text = Trim(.Fields("TCRCD"))
            txtTcrNM.Text = " ":    If IsNull(.Fields("TCRNM")) = False Then txtTcrNM.Text = Trim(.Fields("TCRNM"))
        ElseIf .RecordCount > 1 Then
            sprTcr.Visible = True
            sprTcr.MaxRows = 0
            
            .MoveFirst
            
            For nRec = 1 To .RecordCount Step 1
                sprTcr.MaxRows = sprTcr.MaxRows + 1
                sprTcr.Row = sprTcr.MaxRows
                
                sprTcr.Col = 1:     sTmp = "":      If IsNull(.Fields("TCRCD")) = False Then sTmp = Trim(.Fields("TCRCD"))
                    Call basFunction.Set_SprType_Text(sprTcr, "CENTER", "CENTER", basFunction.LenKor(sTmp), sTmp)
                sprTcr.Col = 2:     sTmp = "":      If IsNull(.Fields("TCRNM")) = False Then sTmp = Trim(.Fields("TCRNM"))
                    Call basFunction.Set_SprType_Text(sprTcr, "CENTER", "LEFT", basFunction.LenKor(sTmp), sTmp)
                    
                .MoveNext
            Next nRec
        End If
    End With
    
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    txtTcrNM.SetFocus
            
    Exit Sub
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    On Error GoTo 0
    MsgBox "강사 조회시 에러가 발생하였습니다.", vbCritical + vbOKOnly, "강사조회"
End Sub

Private Sub txtTcrNM_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    
    Dim nLength     As Long
    Dim sStr        As String
    Dim ni          As Integer
    Dim nRec        As Long
    
    Dim sTmp        As String
    
    On Error GoTo ErrStmt
    
    Select Case Button
        Case vbRightButton
            sStr = ""
            sStr = sStr & "  SELECT ACID, TCRCD, SUBJCD, SUBJGBN, TCRGBN, TCRNM, SUBJNM, TCR_CL"
            sStr = sStr & "    From SDTCR01TB"
            sStr = sStr & "   WHERE ACID   = '" & Trim(basModule.SchCD) & "'"
            If Trim(txtTcrNM.Text) > " " Then
                sStr = sStr & "     AND TCRNM  LIKE '" & Trim(txtTcrNM.Text) & "%'"
            End If
            
        Case Else
            Exit Sub
            
    End Select
    
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
    
    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic형태로 열게되면 record count를 할 수 없음.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop
    
    With DBRec
        If .RecordCount = 1 Then
            .MoveFirst
            
            fpTcrCD.Text = "":      If IsNull(.Fields("TCRCD")) = False Then fpTcrCD.Text = Trim(.Fields("TCRCD"))
            txtTcrNM.Text = " ":    If IsNull(.Fields("TCRNM")) = False Then txtTcrNM.Text = Trim(.Fields("TCRNM"))
        ElseIf .RecordCount > 1 Then
            sprTcr.Visible = True
            sprTcr.MaxRows = 0
            
            .MoveFirst
            
            For nRec = 1 To .RecordCount Step 1
                sprTcr.MaxRows = sprTcr.MaxRows + 1
                sprTcr.Row = sprTcr.MaxRows
                
                sprTcr.Col = 1:     sTmp = "":      If IsNull(.Fields("TCRCD")) = False Then sTmp = Trim(.Fields("TCRCD"))
                    Call basFunction.Set_SprType_Text(sprTcr, "CENTER", "CENTER", basFunction.LenKor(sTmp), sTmp)
                sprTcr.Col = 2:     sTmp = "":      If IsNull(.Fields("TCRNM")) = False Then sTmp = Trim(.Fields("TCRNM"))
                    Call basFunction.Set_SprType_Text(sprTcr, "CENTER", "LEFT", basFunction.LenKor(sTmp), sTmp)
                    
                .MoveNext
            Next nRec
        End If
    End With
    
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    txtTcrNM.SetFocus
            
    Exit Sub
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    On Error GoTo 0
    MsgBox "강사 조회시 에러가 발생하였습니다.", vbCritical + vbOKOnly, "강사조회"
End Sub



Private Sub sprTcr_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            sprTcr.Visible = False
            
    End Select
End Sub

Private Sub sprTcr_Click(ByVal Col As Long, ByVal Row As Long)
    Dim ni      As Long
    
    With sprTcr
        If Row < 1 Then Exit Sub
        If .MaxRows = 0 Then Exit Sub
        
        If Trim(.Tag) = "" Then .Tag = "1"
        
        .Row = CLng(.Tag):  .Row2 = .Row
        .Col = 1:           .Col2 = .MaxCols
        .BlockMode = True
            .BackColor = basModule.WhiteColor
            .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
        
        .Row = Row:         .Row2 = .Row
        .Col = 1:           .Col2 = .MaxCols
        .BlockMode = True
            .BackColor = basModule.SelectColor1
            .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
        
        .Tag = Trim(CStr(Row))
        
    End With
End Sub

Private Sub sprTcr_DblClick(ByVal Col As Long, ByVal Row As Long)
    Dim ni      As Long
    
    With sprTcr
        If Row < 1 Then Exit Sub
        If .MaxRows = 0 Then Exit Sub
        
        If Trim(.Tag) = "" Then .Tag = "1"
        
        .Row = CLng(.Tag):  .Row2 = .Row
        .Col = 1:           .Col2 = .MaxCols
        .BlockMode = True
            .BackColor = basModule.WhiteColor
            .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
        
        .Row = Row:         .Row2 = .Row
        .Col = 1:           .Col2 = .MaxCols
        .BlockMode = True
            .BackColor = basModule.SelectColor1
            .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
        
        .Tag = Trim(CStr(Row))
        
        '>> 데이터 보여주기
        .Row = Row
        .Col = 1:       fpTcrCD.Text = Trim(.Text)
        .Col = 2:       txtTcrNM.Text = Trim(.Text)
        
        .Visible = False
    End With
End Sub




'#######################################################################################################################################################################
' 강사별 시수내역 처리
'#######################################################################################################################################################################
Private Sub cmdFindTmr_Click()
    Dim nCol        As Long
    Dim nColChk     As Long
    
    
    
    sprTmr.MaxRows = 0
    sprTmr.MaxCols = 0
    
    sprTmr.Col = 0:   sprTmr.ColHidden = False
    sprTmr.Row = 0:   sprTmr.RowHidden = False
    
    sprTmr.RowHeaderCols = 1
    sprTmr.ColHeaderRows = 1
    
    ReDim uSisu_Data(0) As tSisu_Data                       '<< 초기화
    
    Call Display_SprTmr_Col_SpreadHeader                    '<< ROW 로 진행하는 COLUMN의 헤더 작성
    
    If sprTmr.RowHeaderCols > 3 Then                            '<< 조회되어진 강사가 있는가를 체크함.
    
        Call Display_SprTmr_Row_SpreadHeader                    '<< COL 로 진행하는 ROW의 헤더 작성
        
        If sprTmr.ColHeaderRows < 4 Then
            sprTmr.MaxCols = 0
            sprTmr.MaxRows = 0
    
            sprTmr.ColHeaderRows = 1
            sprTmr.RowHeaderCols = 1
        Else
            Call Construct_Spread_Sisu_Data(sprTmr.MaxRows, sprTmr.MaxCols)
            
            If sprTmr.ColHeaderRows = 4 Then
                sprTmr.Row = SpreadHeader
                    sprTmr.RowMerge = MergeAlways
                  
                sprTmr.AddCellSpan SpreadHeader, SpreadHeader, 5, 4
                
                sprTmr.Row = SpreadHeader + 1:          sprTmr.RowHidden = True
                
                sprTmr.AddCellSpan SpreadHeader, SpreadHeader + 4, 4, 1
                
            End If
            sprTmr.Col = sprTmr.MaxCols:                                            sprTmr.ColHidden = True
            
            sprTmr.Row = SpreadHeader
            sprTmr.Col = SpreadHeader
            
'            If sprTmr.ColHidden = False Then
'                sprTmr.ColHidden = True
'            End If
            
            
            '## 데이터 넣기
            Call Find_input_SisuData
            
        End If
    End If
    
    With sprTmr
        If .MaxRows >= 2 Then
            .SetActiveCell 1, 2
            .SetFocus
            
        End If
    End With
    
End Sub


'>> ROW 로 진행하는 헤더 작성
Private Sub Display_SprTmr_Col_SpreadHeader()
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    
    Dim nLength     As Long
    Dim sStr        As String
    Dim ni          As Integer
    Dim nRec        As Long
    Dim nHeaders    As Integer
    
    Dim sTmp        As String
    
    If Me.Tag = "LOAD" Then Exit Sub
        
    On Error GoTo ErrStmt
    
    sStr = ""
    sStr = sStr & "  SELECT TCRCD, SUBJCD, TCRNM, SUBJNM, "
    
'>> 강사구분 추가시 반드시 변경해야 함.----------------------------------------------------------------------
    sStr = sStr & "         DECODE(TCRGBN,'99','','10','담임','20','강남출강','30','송파출강' ) AS TCRGBN "
'------------------------------------------------------------------------------------------------------------
    
    sStr = sStr & "    From SDTCR01TB "
    sStr = sStr & "   WHERE ACID   = '" & Trim(basModule.SchCD) & "'"
    If Trim(fpTcrCD.UnFmtText) > " " Then
        sStr = sStr & " AND TCRCD  = '" & Trim(fpTcrCD.UnFmtText) & "'"
    End If
    If Trim(Right(cboSubjGbn.Text, 30)) <> "ALL" Then
        sStr = sStr & " AND SUBJGBN = '" & Trim(Right(cboSubjGbn.Text, 30)) & "'"
    End If
    If Trim(Right(cboTcrGbn.Text, 30)) <> "ALL" Then
        sStr = sStr & " AND TCRGBN  = '" & Trim(Right(cboTcrGbn.Text, 30)) & "'"
    End If
    sStr = sStr & "   ORDER BY TCRCD, SUBJCD "
    
    
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
'        If Trim(Right(cboFindTcrGbn.Text, 30)) <> "ALL" Then
    
    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic형태로 열게되면 record count를 할 수 없음.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop
    
    With DBRec
        If .RecordCount > 0 Then
        
            sprTmr.MaxRows = .RecordCount + 1
            sprTmr.RowHeaderCols = 5
            
            .MoveFirst
            
            
            sprTmr.Row = 1
            sprTmr.Col = SpreadHeader + 4:  sTmp = "소 계"
                sprTmr.Text = sTmp:     sprTmr.ColWidth(sprTmr.Col) = 5
                sprTmr.RowHeight(sprTmr.Row) = 14             '<< 처음 행 : 합계처리
            
            
            For nRec = 1 To .RecordCount Step 1
                sprTmr.Row = nRec + 1
                
                sprTmr.Col = SpreadHeader:      sTmp = "":  If IsNull(.Fields("TCRCD")) = False Then sTmp = Trim(.Fields("TCRCD"))
                    sprTmr.Text = sTmp:     sprTmr.ColWidth(sprTmr.Col) = 3.5
                    sprTmr.FontSize = 8
                    sprTmr.FontBold = False
                sprTmr.Col = SpreadHeader + 1:  sTmp = "":  If IsNull(.Fields("SUBJCD")) = False Then sTmp = Trim(.Fields("SUBJCD"))
                    sprTmr.Text = sTmp:     sprTmr.ColWidth(sprTmr.Col) = 3
                    sprTmr.FontSize = 8
                    sprTmr.FontBold = False
                sprTmr.Col = SpreadHeader + 2:  sTmp = "":  If IsNull(.Fields("TCRNM")) = False Then sTmp = Trim(.Fields("TCRNM"))
                    sprTmr.Text = sTmp:     sprTmr.ColWidth(sprTmr.Col) = 8
                    sprTmr.TypeHAlign = TypeHAlignLeft
                    sprTmr.TypeVAlign = TypeVAlignCenter
                    sprTmr.FontSize = 12
                    sprTmr.FontBold = True
                sprTmr.Col = SpreadHeader + 3:  sTmp = "":  If IsNull(.Fields("SUBJNM")) = False Then sTmp = Trim(.Fields("SUBJNM"))
                    sprTmr.Text = sTmp:     sprTmr.ColWidth(sprTmr.Col) = 8
                    sprTmr.TypeHAlign = TypeHAlignLeft
                    sprTmr.TypeVAlign = TypeVAlignCenter
                    sprTmr.FontSize = 12
                    sprTmr.FontBold = True
                sprTmr.Col = SpreadHeader + 4:  sTmp = " ": If IsNull(.Fields("TCRGBN")) = False Then sTmp = Trim(.Fields("TCRGBN"))
                    sprTmr.Text = sTmp
                    sprTmr.TypeHAlign = TypeHAlignLeft
                    sprTmr.TypeVAlign = TypeVAlignCenter
                    sprTmr.FontSize = 6
                    sprTmr.FontBold = False
                
                sprTmr.RowHeight(sprTmr.Row) = 14
                
                .MoveNext
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
    MsgBox "강사 및 과목조회시 에러가 발생하였습니다.", vbCritical + vbOKOnly, "COL 헤더처리"
End Sub


'>> COL로 진행하는 헤더 작성
Private Sub Display_SprTmr_Row_SpreadHeader()

    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    
    Dim nLength     As Long
    Dim sStr        As String
    Dim ni          As Integer
    Dim nRec        As Long
    
    Dim sTmp        As String
    Dim sKaeyol     As String
    
    If Me.Tag = "LOAD" Then Exit Sub
    
    On Error GoTo ErrStmt
    
    sStr = ""
'    sStr = sStr & "  SELECT DECODE(KAEYOL,'01','인문',"
'    sStr = sStr & "                       '02','자연',"
'    sStr = sStr & "                       '03','예체') KAEYOL,"
'    sStr = sStr & "         LSNCD , LSNNM, LSNCDNM "
'    sStr = sStr & "    From SDLSN01TB "
'    sStr = sStr & "   WHERE ACID   = '" & Trim(basModule.SchCD) & "'"
'    If Trim(Right(cboKaeyol.Text, 30)) <> "ALL" Then
'        sStr = sStr & " AND KAEYOL = '" & Trim(Right(cboKaeyol.Text, 30)) & "'"
'    End If
'    sStr = sStr & "   ORDER BY KAEYOL, LSNCDNM"
    
    
    
    sStr = ""
    sStr = sStr & "    SELECT ACID, LSNCD, LSNNM, LSNCDNM, "
    sStr = sStr & "           DECODE(KAEYOL,'01','인문',"
    sStr = sStr & "                         '02','자연',"
    sStr = sStr & "                         '03','예체') KAEYOL"
    sStr = sStr & "      FROM (SELECT ACID, LSNCD, LSNNM, LSNCDNM, KAEYOL "
    sStr = sStr & "              FROM SDLSN01TB "
    sStr = sStr & "             WHERE ACID = '" & Trim(basModule.SchCD) & "'"
    If Trim(Right(cboKaeyol.Text, 30)) <> "ALL" Then
        sStr = sStr & "           AND KAEYOL = '" & Trim(Right(cboKaeyol.Text, 30)) & "'"
    End If
    sStr = sStr & "            UNION ALL"
    sStr = sStr & "            SELECT ACID, LSNCD, LSNNM, LSNCDNM, KAEYOL"
    sStr = sStr & "              FROM SDLSN02TB"
    sStr = sStr & "             WHERE (ACID, LSNCD)"
    sStr = sStr & "                IN (SELECT ACID, LSNCD"
    sStr = sStr & "                      FROM (SELECT ACID, LSNCD,"
    sStr = sStr & "                                   CASE WHEN (TAMGU1 +"
    sStr = sStr & "                                              TAMGU2 +"
    sStr = sStr & "                                              TAMGU3 +"
    sStr = sStr & "                                              TAMGU4 +"
    sStr = sStr & "                                              TAMGU5 +"
    sStr = sStr & "                                              TAMGU6 +"
    sStr = sStr & "                                              TAMGU7 +"
    sStr = sStr & "                                              TAMGU8 +"
    sStr = sStr & "                                              TAMGU9 +"
    sStr = sStr & "                                              TAMGU10+"
    sStr = sStr & "                                              TAMGU11+"
    sStr = sStr & "                                              J2SEL  +"
    sStr = sStr & "                                              NONSUL1+"
    sStr = sStr & "                                              NONSUL2+"
    sStr = sStr & "                                              NONSUL3+"
    sStr = sStr & "                                              NONSUL4) > 0 THEN"
    sStr = sStr & "                                       1"
    sStr = sStr & "                                   ELSE"
    sStr = sStr & "                                       0"
    sStr = sStr & "                                   END INWON,"
    sStr = sStr & "                                   CASE WHEN (DECODE(TAMGU_CL1  , 16777215, 0, 1)+"
    sStr = sStr & "                                              DECODE(TAMGU_CL2  , 16777215, 0, 1)+"
    sStr = sStr & "                                              DECODE(TAMGU_CL3  , 16777215, 0, 1)+"
    sStr = sStr & "                                              DECODE(TAMGU_CL4  , 16777215, 0, 1)+"
    sStr = sStr & "                                              DECODE(TAMGU_CL5  , 16777215, 0, 1)+"
    sStr = sStr & "                                              DECODE(TAMGU_CL6  , 16777215, 0, 1)+"
    sStr = sStr & "                                              DECODE(TAMGU_CL7  , 16777215, 0, 1)+"
    sStr = sStr & "                                              DECODE(TAMGU_CL8  , 16777215, 0, 1)+"
    sStr = sStr & "                                              DECODE(TAMGU_CL9  , 16777215, 0, 1)+"
    sStr = sStr & "                                              DECODE(TAMGU_CL10 , 16777215, 0, 1)+"
    sStr = sStr & "                                              DECODE(TAMGU_CL11 , 16777215, 0, 1)+"
    sStr = sStr & "                                              DECODE(J2SEL_CL   , 16777215, 0, 1)+"
    sStr = sStr & "                                              DECODE(NONSUL1_CL , 16777215, 0, 1)+"
    sStr = sStr & "                                              DECODE(NONSUL2_CL , 16777215, 0, 1)+"
    sStr = sStr & "                                              DECODE(NONSUL3_CL , 16777215, 0, 1)+"
    sStr = sStr & "                                              DECODE(NONSUL4_CL , 16777215, 0, 1)) > 0 THEN"
    sStr = sStr & "                                       1"
    sStr = sStr & "                                   ELSE"
    sStr = sStr & "                                       0"
    sStr = sStr & "                                   END NCOL"
    sStr = sStr & "                              FROM SDLSN05TB"
    sStr = sStr & "                             WHERE ACID = '" & Trim(basModule.SchCD) & "'"
    If Trim(Right(cboKaeyol.Text, 30)) <> "ALL" Then
        sStr = sStr & "                           AND KAEYOL = '" & Trim(Right(cboKaeyol.Text, 30)) & "'"
    End If
    sStr = sStr & "                            )"
    sStr = sStr & "                     WHERE (INWON > 0 OR NCOL > 0)"
    sStr = sStr & "                       AND LSNCD >= '90000'"
    sStr = sStr & "                    )"
    sStr = sStr & "            )"
    sStr = sStr & "     ORDER BY KAEYOL, LSNCDNM"
    
    
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
'        If Trim(Right(cboFindTcrGbn.Text, 30)) <> "ALL" Then
'    ' KAEYOL
'            sTmp = Trim(Right(cboKaeyol.Text, 30))
'            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'                Set DBParam = DBCmd.CreateParameter("KAEYOL", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
'
'        End If
    
    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic형태로 열게되면 record count를 할 수 없음.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop
    
    With DBRec
        
        If .RecordCount > 0 Then
        
            sprTmr.MaxCols = .RecordCount + 2
            sprTmr.ColHeaderRows = 4
            
            .MoveFirst
            
            For nRec = 1 To .RecordCount Step 1
                sprTmr.Col = nRec
                
                sprTmr.Row = SpreadHeader:      sTmp = "":  If IsNull(.Fields("KAEYOL")) = False Then sTmp = Trim(.Fields("KAEYOL"))
                    sprTmr.Text = sTmp
                    sprTmr.FontSize = 8
                    sprTmr.FontBold = False
                    
                    If nRec = 1 Then sKaeyol = sTmp
                    If StrComp(sKaeyol, sTmp, vbTextCompare) <> 0 Then
                        sprTmr.SetCellBorder sprTmr.Col, 1, sprTmr.Col, sprTmr.MaxRows, 1, basModule.SectionColor1, CellBorderStyleSolid
                        sKaeyol = sTmp
                    End If
                
                sprTmr.Row = SpreadHeader + 1:  sTmp = "":  If IsNull(.Fields("LSNCD")) = False Then sTmp = Trim(.Fields("LSNCD"))
                    sprTmr.Text = sTmp:     sprTmr.ColWidth(sprTmr.Col) = 7
                    sprTmr.FontSize = 8
                    sprTmr.FontBold = False
                sprTmr.Row = SpreadHeader + 2:  sTmp = "":  If IsNull(.Fields("LSNCDNM")) = False Then sTmp = Trim(.Fields("LSNCDNM"))
                    sprTmr.Text = sTmp:     sprTmr.ColWidth(sprTmr.Col) = 7
                    sprTmr.FontSize = 8
                    sprTmr.FontBold = False
                sprTmr.Row = SpreadHeader + 3:  sTmp = "":  If IsNull(.Fields("LSNNM")) = False Then sTmp = Trim(.Fields("LSNNM"))
                    sprTmr.Text = sTmp:     sprTmr.ColWidth(sprTmr.Col) = 7
                    sprTmr.FontSize = 12
                    sprTmr.FontBold = True
                
                
                
                
                .MoveNext
            Next nRec
            
            sprTmr.Row = SpreadHeader + 3
            sprTmr.Col = sprTmr.MaxCols - 1
                sTmp = "합 계":             sprTmr.Text = sTmp
            
            sprTmr.SetCellBorder sprTmr.MaxCols - 1, 1, sprTmr.MaxCols - 1, sprTmr.MaxRows, 1, basModule.SectionColor1, CellBorderStyleSolid
            
        End If
    End With

    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    Exit Sub
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    On Error GoTo 0
    MsgBox "강사 및 과목조회시 에러가 발생하였습니다.", vbCritical + vbOKOnly, "ROW 헤더처리"

End Sub


'## 등록된 내용 조회
Private Sub Find_input_SisuData()
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    
    Dim nLength     As Long
    Dim sStr        As String
    Dim ni          As Integer
    Dim nRec        As Long
    
    Dim nRow        As Long
    Dim nCol        As Long
    
    Dim sTcrCD      As String
    Dim sSubjCD     As String
    
    Dim sLsnCD      As String
    
    Dim nTmp        As Long
    
    On Error GoTo ErrStmt
    
    With sprTmr
        If .MaxRows = 0 Then Exit Sub
        If .MaxCols = 0 Then Exit Sub
        
            
        sStr = ""
        sStr = sStr & "  SELECT A.ACID, TCRCD, SUBJCD, A.LSNCD, SISU "
        sStr = sStr & "    FROM SDTCR11TB A, SDLSN01TB B "
        sStr = sStr & "   WHERE A.ACID  = B.ACID "
        sStr = sStr & "     AND A.LSNCD = B.LSNCD"
        sStr = sStr & "     AND A.ACID  = '" & Trim(basModule.SchCD) & "'"
        sStr = sStr & "  UNION ALL "
        sStr = sStr & "  SELECT ACID, TCRCD, SUBJCD, LSNCD, SISU"
        sStr = sStr & "    From SDTCR11TB"
        sStr = sStr & "   WHERE ACID = '" & Trim(basModule.SchCD) & "'"
        sStr = sStr & "     AND LSNCD >= '90000'"
        
        
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
    
        DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic형태로 열게되면 record count를 할 수 없음.
        Do While DBRec.State And adStateExecuting
            DoEvents
        Loop
        
        With DBRec
            If .RecordCount > 0 Then
                ReDim uSisu_Data(.RecordCount) As tSisu_Data            '<< 데이터 베이스 설정
                
                .MoveFirst
                For nRec = 1 To .RecordCount Step 1
                    
                    If IsNull(.Fields("TCRCD")) = False And IsNull(.Fields("SUBJCD")) = False Then
                       
                            uSisu_Data(nRec).ACID = Trim(.Fields("ACID"))
                            uSisu_Data(nRec).TCRCD = Trim(.Fields("TCRCD"))
                            uSisu_Data(nRec).SUBJCD = Trim(.Fields("SUBJCD"))
                            
                            uSisu_Data(nRec).LSNCD = Trim(.Fields("LSNCD"))
                            uSisu_Data(nRec).SISU = CLng(.Fields("SISU"))
                    End If
                    
                    .MoveNext
                Next nRec
            End If
        End With
        
    End With
    
    '> 데이터 내용 SPREAD에 뿌려주기
    If UBound(uSisu_Data) > 0 Then
        With sprTmr
        
            For nRow = 2 To .MaxRows Step 1
                .Row = nRow
                    .Col = SpreadHeader:            sTcrCD = Trim(.Text)
                    .Col = SpreadHeader + 1:        sSubjCD = Trim(.Text)
                
                For nCol = 1 To (.MaxCols - 1) Step 1
                    .Col = nCol:    .Row = SpreadHeader + 1:    sLsnCD = Trim(.Text)
                    
                    For nRec = 1 To UBound(uSisu_Data) Step 1
                        If StrComp(uSisu_Data(nRec).TCRCD, sTcrCD, vbTextCompare) = 0 And _
                           StrComp(uSisu_Data(nRec).SUBJCD, sSubjCD, vbTextCompare) = 0 And _
                           StrComp(uSisu_Data(nRec).LSNCD, sLsnCD, vbTextCompare) = 0 Then
                           
                            .Row = nRow
                            .Col = nCol
                                nTmp = uSisu_Data(nRec).SISU
                                If nTmp > 0 Then Call basFunction.Set_SprType_Numeric(sprTmr, 0, 0, 999, "", nTmp)
                            
                            Exit For
                            
                        End If
                    Next nRec
                Next nCol
                
            Next nRow
        End With
    End If
        
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    MsgBox "조회하였습니다.", vbInformation + vbOKOnly, "강사 및 과목내역 조회"
    
    Exit Sub
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    On Error GoTo 0
    
    MsgBox "시수 상세내역 조회시 오류가 발생하였습니다.", vbCritical + vbOKOnly, "강사 및 과목내역 조회"
    
End Sub

Private Sub Construct_Spread_Sisu_Data(ByVal aRow As Long, ByVal aCol As Long)
    Dim nCol        As Long
    Dim nRow        As Long
    
    Dim nRowCols    As Long
    Dim sRowEtxt    As String       ' sum row 값 처리 : start
    
    With sprTmr
    
        If aCol < 1 Then
            MsgBox "반의 수가 너무 작습니다.", vbExclamation + vbOKOnly, "합계처리"
            Exit Sub
        End If
        
        If aRow < 1 Then
            MsgBox "선생님의 수가 너무 작습니다.", vbExclamation + vbOKOnly, "합계처리"
            Exit Sub
        End If
        
        '.MaxRows = 0:           .MaxCols = 0                    '## TEST 시에 사용
        '.MaxRows = aRow:        .MaxCols = aCol + 2             '<< row 는 강사 : col 은 시수이고, col에서 maxcols-1(소계) maxcol(선택)
        
        .Row = 1:   .Row2 = .MaxRows
        .Col = 1:   .Col2 = .MaxCols
        .BlockMode = True
            .BackColor = basModule.WhiteColor
            .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
        
        
'        For nCol = 1 To .MaxCols Step 1                         '<< 열의 간격조정. 단, row는 기본값
'            .ColWidth(nCol) = 6
'        Next nCol
        
        
        For nRow = 1 To .MaxRows Step 1
            For nCol = 1 To .MaxCols Step 1                  '<< col 마지막은 선택여부
                
                .Row = nRow
                
                If nCol = .MaxCols Then
                    If .Row = 1 Then
                    
                    Else
                        .Col = nCol
                        
                        .CellType = CellTypeCheckBox
                        .TypeHAlign = TypeHAlignCenter
                        .TypeVAlign = TypeVAlignCenter
                        .Value = 0
                    End If
                    
                Else
                    
                    .Col = nCol
                    
                    .CellType = CellTypeNumber
                    .TypeVAlign = TypeVAlignCenter
                    .TypeNumberDecPlaces = 0
                    .TypeNumberMin = 0
                    .TypeNumberMax = 99
                    
                    .TypeNumberShowSep = False
                End If
                
            Next nCol
        Next nRow
        
       '>> 열 합계 -------------------------------------------------------
            For nCol = 1 To (.MaxCols - 2) Step 1               '<<
                .Row = 1
                .Col = nCol
                .FormulaSync = True
                .Formula = "SUM(#2:#" & Trim(CStr(.MaxRows)) & ")"
                
            Next nCol
            '>> 첫번째 행을 locking
            .Row = 1:       .Row2 = 1
            .Col = 1:       .Col2 = .MaxCols - 1
            .BlockMode = True
                .Lock = True
                .Protect = True
                
                .BackColor = basModule.SelectColor1
                .BackColorStyle = BackColorStyleUnderGrid
            .BlockMode = False
            .SetCellBorder 1, 1, .MaxCols, 1, 8, basModule.SectionColor1, CellBorderStyleSolid
        '-----------------------------------------------------------------
        
        
        '>> 행 합계 ------------------------------------------------------
            '## 선행 값   SUM( A#: ?#)      <- 여기서 ? 항목   x , AA, BA, CA, ... 로 진행 << 처음시작은 A#
                nRowCols = Fix((.MaxCols - 2) / 26)
                If nRowCols = 0 Then
                    sRowEtxt = ""
                Else
                    sRowEtxt = Chr$(64 + nRowCols)
                End If
            '## 후행값
                nRowCols = ((.MaxCols - 2) Mod 26)
                sRowEtxt = sRowEtxt & Chr$(64 + nRowCols)
        
            For nRow = 1 To .MaxRows Step 1
                .Row = nRow
                .Col = .MaxCols - 1
                .FormulaSync = True
                .Formula = "SUM(A#:" & Trim(sRowEtxt) & "#)"
            Next nRow
            
            '>> 마지막 열을 locking
            .Row = 2:               .Row2 = .MaxRows
            .Col = .MaxCols - 1:    .Col2 = .MaxCols
            .BlockMode = True
                .Lock = True
                .Protect = True
                
                .BackColor = basModule.SelectColor2
                .BackColorStyle = BackColorStyleUnderGrid
            .BlockMode = False
            .SetCellBorder .MaxCols, 1, .MaxCols, .MaxRows, 1, basModule.SectionColor1, CellBorderStyleSolid
            
        '----------------------------------------------------------------
        
    End With
End Sub














'>> 등록위한 guide line
Private Sub sprTmr_Click(ByVal Col As Long, ByVal Row As Long)
    
    With sprTmr
        If Row < 2 Then Exit Sub
        If Col < 1 Then Exit Sub
        If Col > .MaxCols - 2 Then Exit Sub
    
        '--------------------------------------------------------------
        .Row = 2:       .Row2 = .MaxRows
        .Col = 1:       .Col2 = .MaxCols - 2
        .BlockMode = True
            .BackColor = basModule.WhiteColor
            .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
        
        '>> 첫번째 행 색
            .Row = 1:       .Row2 = 1
            .Col = 1:       .Col2 = .MaxCols - 1
            .BlockMode = True
                .Lock = True
                .Protect = True
                
                .BackColor = basModule.SelectColor1
                .BackColorStyle = BackColorStyleUnderGrid
            .BlockMode = False
        
        '>> 마지막 열 색
        .Row = 2:               .Row2 = .MaxRows
        .Col = .MaxCols - 1:    .Col2 = .MaxCols
        .BlockMode = True
            .BackColor = basModule.SelectColor2
            .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
        '--------------------------------------------------------------
        
        .Row = 1:       .Row2 = Row
        .Col = Col:     .Col2 = Col
        .BlockMode = True
            .BackColor = basModule.MargentaColor
            .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
        
        .Row = Row:     .Row2 = Row
        .Col = 1:       .Col2 = Col
        .BlockMode = True
            .BackColor = basModule.MargentaColor
            .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
        
    End With
    
End Sub

Private Sub sprTmr_KeyUp(KeyCode As Integer, Shift As Integer)
    
    Dim sTcrCD      As String
    Dim sSubjCD     As String
    Dim sLsnCD      As String
    
    Dim bRet        As Boolean
    
    With sprTmr
        If .ActiveRow < 2 Then Exit Sub
        If .ActiveCol < 1 Then Exit Sub
        If .ActiveCol > .MaxCols - 2 Then Exit Sub
    
        '--------------------------------------------------------------
        .Row = 2:       .Row2 = .MaxRows
        .Col = 1:       .Col2 = .MaxCols - 2
        .BlockMode = True
            .BackColor = basModule.WhiteColor
            .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
        
        '>> 첫번째 행 색
            .Row = 1:       .Row2 = 1
            .Col = 1:       .Col2 = .MaxCols - 1
            .BlockMode = True
                .Lock = True
                .Protect = True
                
                .BackColor = basModule.SelectColor1
                .BackColorStyle = BackColorStyleUnderGrid
            .BlockMode = False
        
        '>> 마지막 열 색
        .Row = 2:               .Row2 = .MaxRows
        .Col = .MaxCols - 1:    .Col2 = .MaxCols
        .BlockMode = True
            .BackColor = basModule.SelectColor2
            .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
        '--------------------------------------------------------------
    
        .Row = 1:               .Row2 = .ActiveRow
        .Col = .ActiveCol:      .Col2 = .ActiveCol
        .BlockMode = True
            .BackColor = basModule.MargentaColor
            .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
        
        .Row = .ActiveRow:      .Row2 = .ActiveRow
        .Col = 1:               .Col2 = .ActiveCol
        .BlockMode = True
            .BackColor = basModule.MargentaColor
            .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
        
        .Row = .ActiveRow:  .Col = .ActiveCol
        If Trim(.Text) <> "" Then
            If .Value > 0 Then
                .Row = .ActiveRow
                .Col = .MaxCols
                    .Value = 1
            End If
        End If
        
        
        If KeyCode = vbKeyDelete Then
            .Row = .ActiveRow
                .Col = SpreadHeader:            sTcrCD = Trim(.Text)
                .Col = SpreadHeader + 1:        sSubjCD = Trim(.Text)
            
            .Col = .ActiveCol
                .Row = SpreadHeader + 1:        sLsnCD = Trim(.Text)
            
            bRet = Del_SisuData(sTcrCD, sSubjCD, sLsnCD)
            If bRet = True Then
                .Row = .ActiveRow
                .Col = .ActiveCol
                    .Text = ""
            End If
        End If
        
    End With
    
End Sub


Private Sub sprTmr_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    Dim nRow        As Long
    Dim nS          As Long
    Dim nE          As Long
    
    With sprTmr
        
        If BlockRow < 1 Then BlockRow = 1
        If BlockRow2 < 1 Then BlockRow2 = 1
        
        nS = BlockRow
        nE = BlockRow2
        If BlockRow > BlockRow2 Then
            nS = BlockRow2
            nE = BlockRow
        End If
        
        For nRow = BlockRow To BlockRow2 Step 1
            .Row = nRow
            .Col = .MaxCols
                .Value = 1
        Next nRow
        
    End With
    
End Sub









Private Function Del_SisuData(ByVal aTcrCD As String, ByVal aSubjCD As String, ByVal aLsnCD As String) As Boolean
    Dim DBCmd       As ADODB.Command
    Dim DBParam     As ADODB.Parameter
    
    Dim sStr        As String
    Dim nLength     As Long
    Dim nExe        As Long
    
    Dim sTmp        As String
    Dim ni          As Long
    
    Dim bRet        As Boolean
    
    On Error GoTo ErrStmt
    bRet = False
    
    basDataBase.DBConn.BeginTrans
    
    Set DBCmd = New ADODB.Command
    Set DBParam = New ADODB.Parameter
    
    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
                


        
    sStr = ""
    sStr = sStr & " DELETE "
    sStr = sStr & "   From SDTCR11TB "
    sStr = sStr & "  WHERE ACID   = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "    AND TCRCD  = '" & aTcrCD & "'"
    sStr = sStr & "    AND SUBJCD = '" & aSubjCD & "'"
    sStr = sStr & "    AND LSNCD  = '" & aLsnCD & "'"
    
    DBCmd.CommandText = sStr
    DBCmd.CommandType = adCmdText
    DBCmd.CommandTimeout = 30
    
    DBCmd.Execute nExe, , -1

    Do While basDataBase.DBConn.State And adStateExecuting
        DoEvents
    Loop
    
    If nExe = 0 Then
        basDataBase.DBConn.RollbackTrans
        bRet = True
    ElseIf nExe = 1 Then
        basDataBase.DBConn.CommitTrans
        bRet = True
    Else
    
ErrStmt:
        basDataBase.DBConn.RollbackTrans
    End If
    
    Set DBCmd = Nothing
    Set DBParam = Nothing
    
    Del_SisuData = bRet
    
End Function





'#######################################################################################################################################################################
' 강사별 시수내역 등록하기
'#######################################################################################################################################################################
Private Sub cmdSaveTmr_Click()
    Dim nChk        As Long
    Dim nRow        As Long
    
    On Error GoTo ErrStmt
    
    With sprTmr
        If .MaxRows = 0 Then Exit Sub
        If .MaxCols = 0 Then Exit Sub
        
        nChk = 0
        
        For nRow = 2 To .MaxRows Step 1
            .Row = nRow
            .Col = .MaxCols
            
            If .Value = 1 Then
                nChk = nChk + 1
            End If
        Next nRow
        
        If nChk = 0 Then
            MsgBox "시수를 넣으신 후 등록버튼을 클릭하세요.", vbExclamation + vbOKOnly, "강사 시수내역 등록"
            Exit Sub
        End If
        
            
        '## 데이터 저장
        Call Save_Detail_Data
        
        
    End With
    
    Exit Sub
ErrStmt:
    
    MsgBox "강사 시수내역 등록시 에러가 발생하였습니다." & vbCrLf & CStr(Err.Number) & vbCrLf & Err.Description, vbCritical + vbOKOnly, "강사 시수내역 등록"
    On Error GoTo 0
    
End Sub

'>> 강사 및 과목내역 등록
Private Sub Save_Detail_Data()
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    
    Dim nLength     As Long
    Dim sStr        As String
    Dim ni          As Integer
    
    Dim nRow        As Long
    Dim nCol        As Long
    
    Dim sTcrCD      As String           ' 강사코드
    Dim sSubjCD     As String           ' 과목코드
    
    Dim sLsnCD      As String           ' 반코드 : header에 있음
    Dim nSisu       As Long             ' 시수
    
    Dim nTotExe     As Long             ' insert/update 되어질 것
    Dim nAddExe     As Long             '               처리된 결과 합
    Dim nExe        As Long             '               처리
    
    Dim nCounts     As Long
    
    On Error GoTo ErrStmt
    
    
    basDataBase.DBConn.BeginTrans

    Set DBCmd = New ADODB.Command
    Set DBRec = New ADODB.Recordset
    Set DBParam = New ADODB.Parameter
    
    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
    
    
    '## 등록할 내용 조회
    With sprTmr
    
        nTotExe = 0
        nAddExe = 0
    
    
        For nRow = 2 To .MaxRows Step 1
            .Row = nRow:
            
            .Col = SpreadHeader:        sTcrCD = Trim(.Text)            '< 강사코드 : HEADER
            .Col = SpreadHeader + 1:    sSubjCD = Trim(.Text)           '< 과목코드 : HEADER + 1
            
            .Col = .MaxCols
            If .Value = 1 Then      '< 시수변경이 생긴 내용만 저장함.
            
                For nCol = 1 To (.MaxCols - 2) Step 1
                    .Col = nCol:                    .Row = SpreadHeader + 1:            sLsnCD = Trim(.Text)        '< 반코드
                    
                    .Row = nRow
                    .Col = nCol
                        If Trim(.Text) > " " Then       '< 데이터 있는 것만 작업함. '0' 포함
                            
                            nTotExe = nTotExe + 1       '<< 작업
                            nSisu = .Value
                            
                            
                            '## SELECT
                            sStr = ""
                            
                            sStr = sStr & " SELECT MAX(CNT) AS CNT"
                            sStr = sStr & "   FROM ("
                            sStr = sStr & "         SELECT 0 AS CNT "
                            sStr = sStr & "           FROM DUAL"
                            sStr = sStr & "         UNION ALL"
                            
                            'sStr = sStr & "        SELECT ACID, TCRCD, SUBJCD, LSNCD, SISU "
                            sStr = sStr & "         SELECT COUNT(*) AS CNT "
                            sStr = sStr & "           FROM SDTCR11TB "
                            sStr = sStr & "          WHERE ACID   = '" & Trim(basModule.SchCD) & "'"
                            sStr = sStr & "            AND TCRCD  = '" & sTcrCD & "'"
                            sStr = sStr & "            AND SUBJCD = '" & sSubjCD & "'"
                            sStr = sStr & "            AND LSNCD  = '" & sLsnCD & "'"
                            sStr = sStr & "         )"
                            
                            DBCmd.CommandText = sStr
                            DBCmd.CommandType = adCmdText
                            DBCmd.CommandTimeout = 30
                    


                                        
'                                ' ACID
'                                    sTmp = Trim(basModule.SchCD)
'                                    nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'                                        Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
                                
                            DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic형태로 열게되면 record count를 할 수 없음.
                            Do While DBRec.State And adStateExecuting
                                DoEvents
                            Loop
                            
                            nCounts = CLng(DBRec.Fields(0))
                            DBRec.Close
                            
                            Select Case nCounts
                                Case 0
                        '< insert >
                                    sStr = ""
                                    sStr = sStr & "  INSERT INTO SDTCR11TB (ACID, TCRCD, SUBJCD, LSNCD, SISU)"
                                    sStr = sStr & "  VALUES ( "
                                    sStr = sStr & "     '" & Trim(basModule.SchCD) & "', "
                                    sStr = sStr & "     '" & sTcrCD & "', "
                                    sStr = sStr & "     '" & sSubjCD & "', "
                                    sStr = sStr & "     '" & sLsnCD & "', "
                                    sStr = sStr & "      " & Trim(CStr(nSisu))
                                    sStr = sStr & "  ) "
                                    
                                    

'                                    ' ACID
'                                        sTmp = Trim(basModule.SchCD)
'                                        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'                                            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
                                            
                                Case Else
                        '< update >
                        
                                    sStr = ""
                                    sStr = sStr & "  UPDATE SDTCR11TB"
                                    sStr = sStr & "     SET SISU   =  " & Trim(CStr(nSisu))
                                    sStr = sStr & "   WHERE ACID   = '" & Trim(basModule.SchCD) & "'"
                                    sStr = sStr & "     AND TCRCD  = '" & Trim(sTcrCD) & "'"
                                    sStr = sStr & "     AND SUBJCD = '" & sSubjCD & "'"
                                    sStr = sStr & "     AND LSNCD  = '" & sLsnCD & "'"
                                        


                            
'                                    ' SISU
'                                        nTmp = nSisu
'                                            Set DBParam = DBCmd.CreateParameter("SISU", adDouble, adParamInput, , nTmp):   DBCmd.Parameters.Append DBParam
'                                    ' ACID
'                                        sTmp = Trim(basModule.SchCD)
'                                        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'                                            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
                            
                                End Select
                                
                                DBCmd.CommandText = sStr
                                DBCmd.CommandType = adCmdText
                                DBCmd.CommandTimeout = 30
                            
                                nExe = 0
                                DBCmd.Execute nExe, , -1
                            
                                Do While basDataBase.DBConn.State And adStateExecuting
                                    DoEvents
                                Loop
                            
                                If nExe = 1 Then
                                    nAddExe = nAddExe + 1
                                End If
                                
                                'DBRec.Close
                                
                            'End If
                        End If
                Next nCol
                
            End If
        Next nRow
    End With
    
    
    If nTotExe = nAddExe Then
        basDataBase.DBConn.CommitTrans
        MsgBox "시수내역 등록하였습니다.", vbInformation + vbOKOnly, "강사 시수내역 등록"
    Else
        basDataBase.DBConn.RollbackTrans
    End If
    
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    Exit Sub
ErrStmt:
    basDataBase.DBConn.RollbackTrans
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    MsgBox "강사 시수내역 등록시 에러가 발생하였습니다." & vbCrLf & CStr(Err.Number) & vbCrLf & Err.Description, vbCritical + vbOKOnly, "강사 시수내역 등록"
    On Error GoTo 0
    
End Sub
