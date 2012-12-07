VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form TMR015 
   Caption         =   "시간표 만들기 >> 강사 강의불가 시간등록"
   ClientHeight    =   9645
   ClientLeft      =   2955
   ClientTop       =   2895
   ClientWidth     =   15840
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9645
   ScaleWidth      =   15840
   Begin FPSpread.vaSpread sprTcr 
      Height          =   2655
      Left            =   180
      TabIndex        =   5
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
      SpreadDesigner  =   "TMR015.frx":0000
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0082C8E8&
      BorderStyle     =   0  '없음
      Caption         =   "Frame1"
      Height          =   735
      Left            =   30
      TabIndex        =   6
      Top             =   30
      Width           =   15435
      Begin VB.Frame Frame2 
         BackColor       =   &H00D2EAF5&
         BorderStyle     =   0  '없음
         Caption         =   "Frame2"
         Height          =   675
         Left            =   30
         TabIndex        =   7
         Top             =   30
         Width           =   15375
         Begin VB.CommandButton cmdSave 
            Caption         =   "제약조건 등록 (&S)"
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
            Left            =   7500
            TabIndex        =   3
            Top             =   90
            Width           =   2325
         End
         Begin VB.CommandButton cmdFind 
            Caption         =   "조  회 (&F)"
            Height          =   495
            Left            =   750
            TabIndex        =   0
            Top             =   90
            Width           =   1245
         End
         Begin VB.TextBox txtTcrNM 
            Height          =   300
            IMEMode         =   10  '한글 
            Left            =   3930
            TabIndex        =   2
            Text            =   "txtTcrNM"
            Top             =   180
            Width           =   1455
         End
         Begin EditLib.fpMask fpTcrCD 
            Height          =   300
            Left            =   3300
            TabIndex        =   1
            Top             =   180
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
         Begin VB.Label Label3 
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
            Left            =   11880
            TabIndex        =   12
            Top             =   240
            Width           =   3555
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '투명
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   210
            Left            =   12720
            TabIndex        =   11
            Top             =   450
            Width           =   225
         End
         Begin VB.Label Label1 
            BackStyle       =   0  '투명
            Caption         =   "9"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C000C0&
            Height          =   210
            Left            =   13830
            TabIndex        =   10
            Top             =   450
            Width           =   225
         End
         Begin VB.Label Label45 
            BackStyle       =   0  '투명
            Caption         =   "※ 입력은 숫자    ,  삭제는     입니다."
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
            Left            =   11310
            TabIndex        =   9
            Top             =   450
            Width           =   3555
         End
         Begin VB.Label Label26 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "강사명"
            Height          =   210
            Left            =   2460
            TabIndex        =   8
            Top             =   225
            Width           =   765
         End
      End
   End
   Begin FPSpread.vaSpread sprTmr 
      Height          =   8535
      Left            =   30
      TabIndex        =   4
      Top             =   810
      Width           =   15435
      _Version        =   393216
      _ExtentX        =   27226
      _ExtentY        =   15055
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
      MaxCols         =   73
      ProcessTab      =   -1  'True
      SpreadDesigner  =   "TMR015.frx":1832
   End
End
Attribute VB_Name = "TMR015"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################################################
'   시 스 템  명 : 대성학원 입학사정, 반배정 & 시간표 프로그램
'   서브시스템명 :
'   모   듈   명 : TRM015
'   모 듈  목 적 : 강사 강의불가 시간등록
'
'   작   성   일 : 2007/12/27
'   작   성   자 : 유하균
' --------------------------------------------------------------------------------------------------------------
' --------------------------------------------------------------------------------------------------------------
'                 수     정     내     용
' --------------------------------------------------------------------------------------------------------------
'   1. 수정일 : TCRGBN 추가시 변경내용 있음.
'   2. 내  용 :
'################################################################################################################

Option Explicit

Private Type tSchdule_Data
    REC         As Long
    ACID        As String
    TCRCD       As String
    SUBJCD      As String
    LESSON      As Long
    WEEKS       As Long
    
    T_SISU      As Long
End Type
Private uSchedule_Data()    As tSchdule_Data



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
        .AddCellSpan SpreadHeader, SpreadHeader, 5, 2
    End With
    
    With sprTcr
        .ShadowColor = basModule.ShadowColor1
        .ShadowDark = basModule.ShadowDark1
        .ShadowText = basModule.ShadowText1
        .GridColor = basModule.GridColor1
        .GrayAreaBackColor = basModule.GrayAreaBackColor1
        
        .MaxRows = 0
        .ZOrder 0
        .Left = 5490
        .Top = 210
    End With
    
    ReDim uSchedule_Data(0) As tSchdule_Data
    
    Me.Tag = "LOAD"
        Call initData
    
    Me.Tag = ""

End Sub

Private Sub initData()
    fpTcrCD.Text = ""
    txtTcrNM.Text = ""
    
    sprTcr.Visible = False
    
    With sprTmr
        .MaxRows = 0
    End With
    
End Sub







'##################################################################################################################################################
'>> 강사조회
'##################################################################################################################################################
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






'##################################################################################################################################################
'>> 강사 강의불가 내역조회
'##################################################################################################################################################
Private Sub cmdFind_Click()

    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter

    Dim sStr        As String
    Dim ni          As Long
    Dim sTmp        As String
    Dim nTmp        As Long
    Dim nRec        As Long

    Dim sAcID       As String
    Dim sTcrCD      As String
    Dim sSubjCD     As String
    
    Dim nRow        As Long

    If Me.Tag = "LOAD" Then Exit Sub

    On Error GoTo ErrStmt

    Set DBCmd = New ADODB.Command
    Set DBRec = New ADODB.Recordset
    Set DBParam = New ADODB.Parameter


    sprTmr.Row = 0

    ReDim uSchedule_Data(0) As tSchdule_Data

    sStr = ""
    sStr = sStr & "  SELECT A.ACID, A.TCRCD||A.SUBJCD AS ID, A.TCRCD, A.SUBJCD, "
    sStr = sStr & "         A.TCRNM, A.SUBJNM, "

'>> 강사구분 추가시 반드시 변경해야 함.----------------------------------------------------------------------
    sStr = sStr & "         DECODE(TCRGBN,'99','','10','담임','20','강남출강','30','송파출강' ) AS TCRGBN, "
'------------------------------------------------------------------------------------------------------------

    sStr = sStr & "         GET_TCR_T_SISU(A.ACID, A.TCRCD, A.SUBJCD) AS T_SISU, "
    sStr = sStr & "         GET_TCR_NOT_SISU(A.ACID, A.TCRCD, A.SUBJCD) AS NOT_SISU, "
    sStr = sStr & "         NVL(LESSON,0) AS LESSON, NVL(WEEKS,0) AS WEEKS"
    sStr = sStr & "    FROM SDTCR01TB A, SDTCR15TB B "
    sStr = sStr & "   WHERE A.ACID   = B.ACID(+) "
    sStr = sStr & "     AND A.TCRCD  = B.TCRCD (+) "
    sStr = sStr & "     AND A.SUBJCD = B.SUBJCD (+) "
    sStr = sStr & "     AND A.ACID   = '" & Trim(basModule.SchCD) & "'"
    If Trim(fpTcrCD.UnFmtText) > " " Then
        sStr = sStr & " AND A.TCRCD  = '" & Trim(fpTcrCD.UnFmtText) & "'"
    End If
    sStr = sStr & "   ORDER BY A.ACID, A.TCRCD, A.SUBJCD"


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

        ReDim uSchedule_Data(0) As tSchdule_Data
        sprTmr.MaxRows = 0
        nRow = 0

        If .RecordCount > 0 Then

            ReDim uSchedule_Data(.RecordCount) As tSchdule_Data

            sAcID = ""
            sTcrCD = ""
            sSubjCD = ""
            sprTmr.RowHeaderCols = 5                '< 강사내역 처리

            sprTmr.MaxRows = sprTmr.MaxRows + 1



            sprTmr.Row = 1
            sprTmr.Col = SpreadHeader + 4:  sTmp = "소 계"
                sprTmr.Text = sTmp:     sprTmr.ColWidth(sprTmr.Col) = 5
                sprTmr.RowHeight(sprTmr.Row) = 14             '<< 처음 행 : 합계처리
            sprTmr.AddCellSpan SpreadHeader, SpreadHeader + 2, 4, 1
            sprTmr.ColsFrozen = 2

            .MoveFirst
            For nRec = 1 To .RecordCount Step 1

                uSchedule_Data(nRec).ACID = "":      If IsNull(.Fields("ACID")) = False Then uSchedule_Data(nRec).ACID = Trim(.Fields("ACID"))
                uSchedule_Data(nRec).TCRCD = "":     If IsNull(.Fields("TCRCD")) = False Then uSchedule_Data(nRec).TCRCD = Trim(.Fields("TCRCD"))
                uSchedule_Data(nRec).SUBJCD = "":    If IsNull(.Fields("SUBJCD")) = False Then uSchedule_Data(nRec).SUBJCD = Trim(.Fields("SUBJCD"))

                uSchedule_Data(nRec).LESSON = 0:     If IsNumeric(.Fields("LESSON")) = True Then uSchedule_Data(nRec).LESSON = CLng(.Fields("LESSON"))
                uSchedule_Data(nRec).WEEKS = 0:      If IsNumeric(.Fields("WEEKS")) = True Then uSchedule_Data(nRec).WEEKS = CLng(.Fields("WEEKS"))
                
                uSchedule_Data(nRec).T_SISU = 0:     If IsNumeric(.Fields("T_SISU")) = True Then uSchedule_Data(nRec).T_SISU = CLng(.Fields("T_SISU"))

            '>> ROW 추가
                If (StrComp(sAcID, uSchedule_Data(nRec).ACID, vbTextCompare) <> 0) Or _
                   (StrComp(sTcrCD, uSchedule_Data(nRec).TCRCD, vbTextCompare) <> 0) Or _
                   (StrComp(sSubjCD, uSchedule_Data(nRec).SUBJCD, vbTextCompare) <> 0) Then
                    sprTmr.MaxRows = sprTmr.MaxRows + 1
                    sprTmr.Row = sprTmr.MaxRows:        sprTmr.RowHeight(sprTmr.Row) = 14
                    
                    nRow = sprTmr.Row
                    

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
                        
                        
                    sprTmr.Col = 1:         nTmp = 0:   If IsNumeric(.Fields("T_SISU")) = True Then nTmp = Trim(.Fields("T_SISU"))
                        Call basFunction.Set_SprType_Numeric(sprTmr, 0, -999, 999, "", nTmp)
                        
                    
                    '< 제약시수는 내용 DISPLAY 하면서 SUM >
                    
                    sAcID = uSchedule_Data(nRec).ACID
                    sTcrCD = uSchedule_Data(nRec).TCRCD
                    sSubjCD = uSchedule_Data(nRec).SUBJCD
    
                End If
                
                uSchedule_Data(nRec).REC = nRow
                
                .MoveNext
            Next nRec

            If sprTmr.MaxRows > 1 Then Call Construct_Spread_Not_Sisu_Data(1, 1)
            
        End If
    End With
    
    
'>> 데이터 넣음.
    For nRec = 1 To UBound(uSchedule_Data) Step 1
        With uSchedule_Data(nRec)

            sprTmr.Row = .REC

            Select Case .WEEKS
                Case 2 To 7
                    sprTmr.Col = 1 + (10 * (.WEEKS - 2)) + 1
                    sprTmr.Col = sprTmr.Col + .LESSON
                        Call basFunction.Set_SprType_Numeric(sprTmr, 0, 1, 9, "", 1)            '< 선택내용
                        
                Case 1
                    sprTmr.Col = 1 + (10 * (8 - 2)) + 1
                    sprTmr.Col = sprTmr.Col + .LESSON
                        Call basFunction.Set_SprType_Numeric(sprTmr, 0, 1, 9, "", 1)            '< 선택내용
                        
                Case Else
                    'NO ACTION
            End Select
        End With
    Next nRec
    
    With sprTmr
        .Col = .MaxCols:        .ColHidden = True
        
        If .MaxCols > 3 Then
            For ni = 3 To .MaxCols Step 10
                Call .SetCellBorder(ni, 1, ni, .MaxRows, 1, basModule.SectionColor1, CellBorderStyleSolid)
            Next ni
        End If
    End With
    
    Set DBCmd = Nothing
    Set DBParam = Nothing
    Set DBRec = Nothing

    Exit Sub
ErrStmt:
    Set DBCmd = Nothing
    Set DBParam = Nothing
    Set DBRec = Nothing

    On Error GoTo 0
    MsgBox "강사 강의불가 내용조회시 에러가 발생하였습니다.", vbCritical + vbOKOnly, "강사 강의불가 내용조회"
End Sub


Private Sub Construct_Spread_Not_Sisu_Data(ByVal aRow As Long, ByVal aCol As Long)
    Dim nCol        As Long
    Dim nRow        As Long
    
    Dim nRowCols    As Long
    Dim sRowEtxt    As String       ' sum row 값 처리 : start
    
    With sprTmr
        
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
                    .TypeNumberMax = 9
                    
                    .TypeNumberShowSep = False
                End If
                
            Next nCol
        Next nRow
        
       '>> 열 합계 -------------------------------------------------------
            For nCol = 1 To (.MaxCols - 1) Step 1               '<<
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
                nRowCols = Fix((.MaxCols - 1) / 26)
                If nRowCols = 0 Then
                    sRowEtxt = ""
                Else
                    sRowEtxt = Chr$(64 + nRowCols)
                End If
            '## 후행값
                nRowCols = ((.MaxCols - 1) Mod 26)
                sRowEtxt = sRowEtxt & Chr$(64 + nRowCols)
        
            For nRow = 1 To .MaxRows Step 1
                .Row = nRow
                .Col = 2
                .FormulaSync = True
                .Formula = "SUM(C#:" & Trim(sRowEtxt) & "#)"
            Next nRow
            
            '>> 1,2 and 마지막 열을 locking
            .Row = 2:               .Row2 = .MaxRows
            .Col = 1:               .Col2 = 2
            .BlockMode = True
                .Lock = True
                .Protect = True
                
                .BackColor = basModule.SelectColor2
                .BackColorStyle = BackColorStyleUnderGrid
            .BlockMode = False
            .SetCellBorder 2, 1, 2, .MaxRows, 2, basModule.SectionColor1, CellBorderStyleSolid
            
            .Row = 1:              .Row2 = .MaxRows
            .Col = .MaxCols:       .Col2 = .MaxCols
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
        If Row <= 1 Then Exit Sub
        If Col <= 2 Then Exit Sub
        '--------------------------------------------------------------
        .Row = 2:       .Row2 = .MaxRows
        .Col = 3:       .Col2 = .MaxCols - 1
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
        
        '>> 1,2 열 색
        .Row = 2:               .Row2 = .MaxRows
        .Col = 1:               .Col2 = 2
        .BlockMode = True
            .BackColor = basModule.SelectColor2
            .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
        
        '>> 마지막 열 색
        .Row = 2:               .Row2 = .MaxRows
        .Col = .MaxCols:        .Col2 = .MaxCols
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
    Dim sWeek       As String
    Dim sLesson     As String
    
    With sprTmr
        If .ActiveRow <= 1 Then Exit Sub
        If .ActiveCol <= 2 Then Exit Sub
        If .ActiveCol >= .MaxCols Then Exit Sub
    
        '--------------------------------------------------------------
        .Row = 2:       .Row2 = .MaxRows
        .Col = 3:       .Col2 = .MaxCols - 1
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
        
        '>> 1,2 열 색
        .Row = 2:               .Row2 = .MaxRows
        .Col = 1:               .Col2 = 2
        .BlockMode = True
            .BackColor = basModule.SelectColor2
            .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
        
        '>> 마지막 열 색
        .Row = 2:               .Row2 = .MaxRows
        .Col = .MaxCols:        .Col2 = .MaxCols
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
            .Col = .ActiveCol
                .Text = ""
                
            .Row = .ActiveRow
            .Col = SpreadHeader:            sTcrCD = Trim(.Text)
            .Col = SpreadHeader + 1:        sSubjCD = Trim(.Text)
            
            .Col = .ActiveCol
            Select Case (.Col - 2)              '< 요일 계산
                Case 1 To 10
                    sWeek = "2"
                Case 11 To 20
                    sWeek = "3"
                Case 21 To 30
                    sWeek = "4"
                Case 31 To 40
                    sWeek = "5"
                Case 41 To 50
                    sWeek = "6"
                Case 51 To 60
                    sWeek = "7"
                Case 61 To 70
                    sWeek = "1"
            End Select
            .Row = SpreadHeader + 1:        sLesson = Trim(CLng(.Text))
            
            Call Del_NotTeaching(sTcrCD, sSubjCD, sWeek, sLesson)
            
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



Private Sub Del_NotTeaching(ByVal aTcrCD As String, ByVal aSubjCD As String, ByVal aWeek As String, ByVal aLesson As String)
    Dim DBCmd       As ADODB.Command
    Dim DBParam     As ADODB.Parameter
    
    Dim sStr        As String
    Dim nLength     As Long
    Dim nExe        As Long
    
    Dim sTmp        As String
    Dim ni          As Long
    
    On Error GoTo ErrStmt
    
    basDataBase.DBConn.BeginTrans
    
    Set DBCmd = New ADODB.Command
    Set DBParam = New ADODB.Parameter
    
    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
                


        
    sStr = ""
    sStr = sStr & " DELETE "
    sStr = sStr & "   From SDTCR15TB"
    sStr = sStr & "  WHERE ACID   = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "    AND TCRCD  = '" & aTcrCD & "'"
    sStr = sStr & "    AND SUBJCD = '" & aSubjCD & "'"
    sStr = sStr & "    AND WEEKS  = " & aWeek
    sStr = sStr & "    AND LESSON = " & aLesson
    
    DBCmd.CommandText = sStr
    DBCmd.CommandType = adCmdText
    DBCmd.CommandTimeout = 30
    
    DBCmd.Execute nExe, , -1

    Do While basDataBase.DBConn.State And adStateExecuting
        DoEvents
    Loop
    
    
    If nExe = 1 Then
        basDataBase.DBConn.CommitTrans
    Else
    
ErrStmt:
        basDataBase.DBConn.RollbackTrans
    End If
    
    Set DBCmd = Nothing
    Set DBParam = Nothing
    
End Sub






'#######################################################################################################################################################################
' 강사별 시수내역 등록하기
'#######################################################################################################################################################################
Private Sub cmdSave_Click()
    Dim nChk        As Long
    Dim nRow        As Long
    
    On Error GoTo ErrStmt
    
    With sprTmr
        If .MaxRows <= 1 Then Exit Sub
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
            MsgBox "등록할 제약조건이 없습니다." & vbCrLf & _
                   "선택 후 제약조건 등록 버튼을 클릭하십시요.", vbExclamation + vbOKOnly, "강사 강의불가 시간등록"
            Exit Sub
        End If
        
        
        '## 데이터 저장
        Call Save_Detail_Data
        
        
    End With
    
    Exit Sub
ErrStmt:
    On Error GoTo 0
    MsgBox "강사 강의불가 시간등록시 에러가 발생하였습니다.", vbCritical + vbOKOnly, "강사 강의불가 시간등록"
    
End Sub

'>> 강사 강의불가 시간등록
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
    Dim sSubjCD     As String
    
    Dim nLesson     As Long             ' 교시 계산
    Dim nWeek       As Long             ' 요일 계산
    
    Dim nTotExe     As Long             ' insert/update 되어질 것
    Dim nAddExe     As Long             '               처리된 결과 합
    Dim nExe        As Long             '               처리
    
    Dim nCountn     As Long
    
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
            If .Value = 1 Then      '< 강의불가 등록 내용만 저장함.
            
                For nCol = 3 To (.MaxCols - 1) Step 1       '< 시작점 column : 3,  종료점 column (max - 1)  / 마지막은 선택
                
                    .Row = nRow
                    .Col = nCol
                        If Trim(.Text) = "1" Then       '< 1 인 내용만 등록
                            
                            Select Case (.Col - 2)              '< 요일 계산
                                Case 1 To 10
                                    nWeek = 2
                                Case 11 To 20
                                    nWeek = 3
                                Case 21 To 30
                                    nWeek = 4
                                Case 31 To 40
                                    nWeek = 5
                                Case 41 To 50
                                    nWeek = 6
                                Case 51 To 60
                                    nWeek = 7
                                Case 61 To 70
                                    nWeek = 1
                            End Select
                            
                            Select Case (.Col - 2) Mod 10       '< 교시 계산
                                Case 1 To 9
                                    nLesson = (.Col - 2) Mod 10
                                Case 0
                                    nLesson = 10
                            End Select
                            
                            
                            '## SELECT
                            sStr = ""
                            
                            sStr = ""
                            sStr = sStr & " SELECT MAX(CNT) AS CNT"
                            sStr = sStr & "   FROM ( "
                            sStr = sStr & "         SELECT 0 AS CNT "
                            sStr = sStr & "           FROM DUAL"
                            sStr = sStr & "         UNION ALL"
                            'sStr = sStr & "        SELECT ACID, TCRCD, SUBJCD, LESSON, WEEKS "
                            sStr = sStr & "         SELECT COUNT(*) AS CNT "
                            sStr = sStr & "           FROM SDTCR15TB "
                            sStr = sStr & "          WHERE ACID   = '" & Trim(basModule.SchCD) & "'"
                            sStr = sStr & "            AND TCRCD  = '" & sTcrCD & "'"
                            sStr = sStr & "            AND SUBJCD = '" & sSubjCD & "'"
                            sStr = sStr & "            AND LESSON = " & Trim(CStr(nLesson))
                            sStr = sStr & "            AND WEEKS  = " & Trim(CStr(nWeek))
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
                            
                            
                            nCountn = CLng(DBRec.Fields(0))
                            DBRec.Close
                            
                            
                            If nCountn = 0 Then
                            
                                nTotExe = nTotExe + 1       '<< 작업
                                
                                sStr = ""
                                sStr = sStr & "  INSERT INTO SDTCR15TB (ACID, TCRCD, SUBJCD, LESSON, WEEKS)"
                                sStr = sStr & "  VALUES ( "
                                sStr = sStr & "     '" & Trim(basModule.SchCD) & "', "
                                sStr = sStr & "     '" & sTcrCD & "', "
                                sStr = sStr & "     '" & sSubjCD & "', "
                                sStr = sStr & "      " & Trim(CStr(nLesson)) & ", "
                                sStr = sStr & "      " & Trim(CStr(nWeek))
                                sStr = sStr & "  ) "
                                
                                    


                                    
    '                            ' ACID
    '                                sTmp = Trim(basModule.SchCD)
    '                                nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
    '                                    Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
                                        
                                
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
                            End If
                            
                            
                        
                        ElseIf Trim(.Text) = "9" Then   '< 등록내역 삭제
                        
                            Select Case (.Col - 2)              '< 요일 계산
                                Case 1 To 10
                                    nWeek = 2
                                Case 11 To 20
                                    nWeek = 3
                                Case 21 To 30
                                    nWeek = 4
                                Case 31 To 40
                                    nWeek = 5
                                Case 41 To 50
                                    nWeek = 6
                                Case 51 To 60
                                    nWeek = 7
                                Case 61 To 70
                                    nWeek = 1
                            End Select
                            
                            Select Case (.Col - 2) Mod 10       '< 교시 계산
                                Case 1 To 9
                                    nLesson = (.Col - 2) Mod 10
                                Case 0
                                    nLesson = 10
                            End Select
                            
                            '## SELECT
                            sStr = ""
                            sStr = sStr & " SELECT MAX(CNT) AS CNT"
                            sStr = sStr & "   FROM ("
                            sStr = sStr & "         SELECT 0 AS CNT "
                            sStr = sStr & "           FROM DUAL"
                            sStr = sStr & "         UNION ALL"
                            'sStr = sStr & "        SELECT ACID, TCRCD, SUBJCD, LESSON, WEEKS "
                            sStr = sStr & "         SELECT COUNT(*) AS CNT "
                            sStr = sStr & "           FROM SDTCR15TB "
                            sStr = sStr & "          WHERE ACID   = '" & Trim(basModule.SchCD) & "'"
                            sStr = sStr & "            AND TCRCD  = '" & sTcrCD & "'"
                            sStr = sStr & "            AND SUBJCD = '" & sSubjCD & "'"
                            sStr = sStr & "            AND LESSON = " & Trim(CStr(nLesson))
                            sStr = sStr & "            AND WEEKS  = " & Trim(CStr(nWeek))
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
                            
                            nCountn = CLng(DBRec.Fields(0))
                            DBRec.Close
                            
                            If nCountn = 1 Then
                                
                                nTotExe = nTotExe + 1       '<< 작업
                                
                                sStr = ""
                                sStr = sStr & "  DELETE"
                                sStr = sStr & "    FROM SDTCR15TB"
                                sStr = sStr & "  WHERE ACID   = '" & Trim(basModule.SchCD) & "'"
                                sStr = sStr & "    AND TCRCD  = '" & sTcrCD & "'"
                                sStr = sStr & "    AND SUBJCD = '" & sSubjCD & "'"
                                sStr = sStr & "    AND LESSON = " & Trim(CStr(nLesson))
                                sStr = sStr & "    AND WEEKS  = " & Trim(CStr(nWeek))
                                    


                    
'                                ' SISU
'                                    nTmp = nSisu
'                                        Set DBParam = DBCmd.CreateParameter("SISU", adDouble, adParamInput, , nTmp):   DBCmd.Parameters.Append DBParam
'                                ' ACID
'                                    sTmp = Trim(basModule.SchCD)
'                                    nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'                                        Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
                                
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
                                
                            End If
                            
                        End If
                        
                Next nCol
                
            End If
        Next nRow
    End With
    
    
    If nTotExe = nAddExe Then
        basDataBase.DBConn.CommitTrans
        MsgBox "강사 강의불가내역 등록하였습니다.", vbInformation + vbOKOnly, "강사 강의불가 내역등록"
    Else
        basDataBase.DBConn.RollbackTrans
        MsgBox "등록시 에러가 발생하였습니다.", vbCritical + vbOKOnly, "강사 강의불가 내역등록"
    End If
    
    Set DBCmd = Nothing
    Set DBRec = Nothing
    Set DBParam = Nothing
    
    Exit Sub
ErrStmt:
    basDataBase.DBConn.RollbackTrans
    
    Set DBCmd = Nothing
    Set DBRec = Nothing
    Set DBParam = Nothing
    
    On Error GoTo 0
    MsgBox "강사 강의불가내역 등록시 에러가 발생하였습니다.", vbCritical + vbOKOnly, "강사 강의불가 내역등록"
    
End Sub



























