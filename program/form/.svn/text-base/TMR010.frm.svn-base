VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form TMR010 
   Caption         =   "시간표 만들기 >> 강사 및 시수넣기"
   ClientHeight    =   9450
   ClientLeft      =   315
   ClientTop       =   2925
   ClientWidth     =   15630
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9450
   ScaleWidth      =   15630
   Begin VB.Frame fraMain 
      BackColor       =   &H00C6AD84&
      BorderStyle     =   0  '없음
      Caption         =   "Frame1"
      Height          =   885
      Left            =   30
      TabIndex        =   21
      Top             =   30
      Width           =   15465
      Begin VB.Frame Frame3 
         BackColor       =   &H00F7EFE7&
         BorderStyle     =   0  '없음
         Caption         =   "Frame3"
         Height          =   825
         Left            =   30
         TabIndex        =   23
         Top             =   30
         Width           =   11175
         Begin VB.CommandButton cmdFindTmr 
            Caption         =   "강사 및 과목내역 조회"
            Height          =   600
            Left            =   3090
            TabIndex        =   2
            Top             =   120
            Width           =   2000
         End
         Begin VB.ComboBox cboKaeyol 
            Height          =   300
            Left            =   870
            Style           =   2  '드롭다운 목록
            TabIndex        =   1
            Top             =   450
            Width           =   2025
         End
         Begin VB.ComboBox cboFindTcrGbn 
            Height          =   300
            Left            =   870
            Style           =   2  '드롭다운 목록
            TabIndex        =   0
            Top             =   120
            Width           =   2025
         End
         Begin VB.CommandButton cmdSaveTmr 
            Caption         =   "시수내역 등록하기"
            Height          =   600
            Left            =   5850
            TabIndex        =   3
            Top             =   120
            Width           =   2000
         End
         Begin VB.Label Label5 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "과목"
            Height          =   210
            Left            =   -180
            TabIndex        =   25
            Top             =   180
            Width           =   975
         End
         Begin VB.Label Label4 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "계열"
            Height          =   210
            Left            =   -180
            TabIndex        =   24
            Top             =   510
            Width           =   975
         End
      End
      Begin VB.Frame fraMain1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  '없음
         Caption         =   "Frame3"
         Height          =   825
         Left            =   11220
         TabIndex        =   22
         Top             =   30
         Width           =   4215
         Begin VB.CommandButton cmdSubj_inSert 
            Caption         =   "강사 및 과목내역 등록"
            Height          =   600
            Left            =   600
            TabIndex        =   4
            Top             =   120
            Width           =   2505
         End
      End
   End
   Begin VB.Frame fraSubj 
      BackColor       =   &H0082C8E8&
      BorderStyle     =   0  '없음
      Caption         =   "Frame1"
      Height          =   7395
      Left            =   4740
      TabIndex        =   14
      Top             =   1560
      Visible         =   0   'False
      Width           =   5025
      Begin VB.Frame Frame2 
         BackColor       =   &H00D2EAF5&
         BorderStyle     =   0  '없음
         Caption         =   "Frame2"
         Height          =   7335
         Left            =   30
         TabIndex        =   15
         Top             =   30
         Width           =   4965
         Begin VB.CommandButton cmdNewSisu 
            Caption         =   "신 규"
            Height          =   400
            Left            =   210
            TabIndex        =   20
            Top             =   510
            Width           =   1000
         End
         Begin VB.CommandButton cmdDeleteSisu 
            Caption         =   "삭 제"
            Height          =   400
            Left            =   3750
            TabIndex        =   12
            Top             =   510
            Width           =   1000
         End
         Begin VB.CommandButton cmdSaveSisu 
            Caption         =   "저 장"
            Height          =   400
            Left            =   2550
            TabIndex        =   11
            Top             =   510
            Width           =   1000
         End
         Begin VB.CommandButton cmdFindSisu 
            Caption         =   "조 회"
            Height          =   400
            Left            =   1380
            TabIndex        =   10
            Top             =   510
            Width           =   1000
         End
         Begin VB.TextBox txtSisuCD 
            Height          =   300
            Left            =   1350
            TabIndex        =   6
            Text            =   "txtSisuCD"
            Top             =   150
            Visible         =   0   'False
            Width           =   1455
         End
         Begin FPSpread.vaSpread sprSubj 
            Height          =   4875
            Left            =   120
            TabIndex        =   13
            Top             =   2310
            Width           =   4695
            _Version        =   393216
            _ExtentX        =   8281
            _ExtentY        =   8599
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
            SpreadDesigner  =   "TMR010.frx":0000
         End
         Begin VB.TextBox txtSubjNM 
            Height          =   300
            IMEMode         =   10  '한글 
            Left            =   1230
            TabIndex        =   9
            Text            =   "txtSubjNM"
            Top             =   1890
            Width           =   1455
         End
         Begin VB.ComboBox cboTcrGbn 
            Height          =   300
            Left            =   1230
            Style           =   2  '드롭다운 목록
            TabIndex        =   8
            Top             =   1560
            Width           =   1455
         End
         Begin VB.TextBox txtTcrNM 
            Height          =   300
            IMEMode         =   10  '한글 
            Left            =   1230
            TabIndex        =   7
            Text            =   "txtTcrNM"
            Top             =   1170
            Width           =   1455
         End
         Begin VB.Label Label6 
            Alignment       =   2  '가운데 맞춤
            BackStyle       =   0  '투명
            Caption         =   "닫기"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00CB5C56&
            Height          =   375
            Left            =   4020
            TabIndex        =   26
            Top             =   120
            Width           =   1035
         End
         Begin VB.Label Label3 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "강사시수코드"
            Height          =   210
            Left            =   60
            TabIndex        =   19
            Top             =   180
            Visible         =   0   'False
            Width           =   1155
         End
         Begin VB.Label Label2 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "과목명"
            Height          =   210
            Left            =   120
            TabIndex        =   18
            Top             =   1950
            Width           =   975
         End
         Begin VB.Label Label1 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "과목구분"
            Height          =   210
            Left            =   120
            TabIndex        =   17
            Top             =   1620
            Width           =   975
         End
         Begin VB.Label Label26 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "강사명"
            Height          =   210
            Left            =   120
            TabIndex        =   16
            Top             =   1200
            Width           =   975
         End
      End
   End
   Begin FPSpread.vaSpread sprTmr 
      Height          =   8385
      Left            =   30
      TabIndex        =   5
      Top             =   960
      Width           =   15465
      _Version        =   393216
      _ExtentX        =   27279
      _ExtentY        =   14790
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
      SpreadDesigner  =   "TMR010.frx":18EE
   End
End
Attribute VB_Name = "TMR010"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################################################
'   시 스 템  명 : 대성학원 입학사정, 반배정 & 시간표 프로그램
'   서브시스템명 :
'   모   듈   명 : TRM010
'   모 듈  목 적 : 강사 및 시수넣기
'
'   작   성   일 : 2007/10/31
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
    SISUCD      As String
    LSNCD       As String
    SISU        As Long
End Type
Private uSisu_Data()    As tSisu_Data

Private Sub Label6_Click()
    fraSubj.Visible = False
End Sub


Private Sub Form_Click()
    fraSubj.Visible = False
End Sub

Private Sub fraMain_Click()
    fraSubj.Visible = False
End Sub

Private Sub fraMain1_Click()
    fraSubj.Visible = False
End Sub

Private Sub Frame3_DragDrop(Source As Control, x As Single, y As Single)
    fraSubj.Visible = False
End Sub


Private Sub Form_Load()
    Me.Move 0, 0, 15700, 9980
    
    With sprTmr
        .ShadowColor = basModule.ShadowColor2
        .ShadowDark = basModule.ShadowDark2
        .ShadowText = basModule.ShadowText2
        .GridColor = basModule.GridColor2
        .GrayAreaBackColor = basModule.GrayAreaBackColor2
    End With

    With sprSubj
        .ShadowColor = basModule.ShadowColor1
        .ShadowDark = basModule.ShadowDark1
        .ShadowText = basModule.ShadowText1
        .GridColor = basModule.GridColor1
        .GrayAreaBackColor = basModule.GrayAreaBackColor1
    End With
    
    With cboTcrGbn
        .Clear
        
        .AddItem "언어" & Space(30) & "10"
        .AddItem "수리" & Space(30) & "20"
        .AddItem "Eng" & Space(30) & "30"
        .AddItem "사탐" & Space(30) & "40"
        .AddItem "과탐" & Space(30) & "50"
        
        .ListIndex = 0
    End With
    
    With cboFindTcrGbn
        .Clear
        
        .AddItem "전체" & Space(30) & "ALL"
        .AddItem "언어" & Space(30) & "10"
        .AddItem "수리" & Space(30) & "20"
        .AddItem "Eng" & Space(30) & "30"
        .AddItem "사탐" & Space(30) & "40"
        .AddItem "과탐" & Space(30) & "50"
        
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
    fraSubj.Visible = False
    
    Me.Tag = "LOAD"
        Call init_Form
    
    Me.Tag = ""



End Sub

Private Sub cmdSubj_inSert_Click()
    fraSubj.Top = fraMain.Top + fraMain1.Top + cmdSubj_inSert.Top + cmdSubj_inSert.Height + 30
    fraSubj.Left = fraMain.Left + fraMain1.Left + cmdSubj_inSert.Left - fraSubj.Width + cmdSubj_inSert.Width
    fraSubj.ZOrder 0
    fraSubj.Visible = True
End Sub

Private Sub init_Form()
    
    txtSisuCD.Text = ""
    txtTcrNM.Text = ""
    txtSubjNM.Text = ""
    
    sprTmr.MaxRows = 0:     sprTmr.MaxCols = 0
    sprSubj.MaxRows = 0
    
        
End Sub




'##########################################################################################################

Private Sub sprSubj_Click(ByVal Col As Long, ByVal Row As Long)
    Dim sTmp        As String
    
    If Row = 0 Then Exit Sub
    
    With sprSubj
        If Trim(.Tag) = "" Then .Tag = "1"
        
        .Row = CLng(.Tag):  .Row2 = .Row
        .Col = 1:           .Col2 = .MaxCols
        .BlockMode = True
            .BackColor = basModule.WhiteColor
            .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
        
        .Row = Row:     .Row2 = .Row
        .Col = 1:       .Col2 = .MaxCols
        .BlockMode = True
            .BackColor = basModule.SelectColor1
            .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
        
        .Tag = Trim(CStr(Row))
        
        .Row = Row
        .Col = 1:           sTmp = Trim(.Text):     txtSisuCD.Text = sTmp
        .Col = .Col + 1:    sTmp = Trim(.Text):     txtTcrNM.Text = sTmp
        .Col = .Col + 1:    sTmp = Trim(.Text)
            Select Case sTmp        '<< 과목구분 내역
                Case "10"
                    cboTcrGbn.ListIndex = 0
                Case "20"
                    cboTcrGbn.ListIndex = 1
                Case "30"
                    cboTcrGbn.ListIndex = 2
                Case "40"
                    cboTcrGbn.ListIndex = 3
                Case "50"
                    cboTcrGbn.ListIndex = 4
            End Select
        .Col = .Col + 1     '<< skip
        .Col = .Col + 1:    sTmp = Trim(.Text):     txtSubjNM.Text = sTmp
    End With
End Sub


'>> 강사 및 과목내역 삭제
Private Sub cmdDeleteSisu_Click()
    
    Dim sStr        As String
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    
    Dim ni          As Long
    
    Dim sTmp        As String
    Dim nLength     As Long
    Dim nExe        As Long
    
    Dim nRow        As Long

    If Trim(txtSisuCD.Text) = "" Then
        MsgBox "삭제할 내용이 없습니다." & vbCrLf & _
               "삭제대상을 선택하십시요.", vbExclamation + vbOKOnly, "강사 및 과목내역 삭제"
        Exit Sub
    End If
    If Trim(txtTcrNM.Text) = "" Then
        MsgBox "강사명이 없습니다.", vbExclamation + vbOKOnly, "강사 및 과목내역 삭제"
        Exit Sub
    End If
    If Trim(txtSubjNM.Text) = "" Then
        MsgBox "과목명이 없습니다.", vbExclamation + vbOKOnly, "강사 및 과목내역 삭제"
        Exit Sub
    End If
    
    
    If MsgBox(Trim(txtTcrNM.Text) & "강사의" & vbCrLf & _
              Trim(txtSubjNM.Text) & "과목내용을" & vbCrLf & _
              "삭제하시겠습니까?", vbQuestion + vbYesNo, "강사 및 과목내역 삭제") = vbNo Then
        Exit Sub
    End If
    
    On Error GoTo ErrStmt
    
    basDataBase.DBConn.BeginTrans
    
    Set DBCmd = New ADODB.Command
    Set DBParam = New ADODB.Parameter
    
    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
                
    '>> 기존 파라미터가 남아 있으면 메모리에서 삭제함.
    For ni = 0 To DBCmd.Parameters.Count - 1 Step 1
        DBCmd.Parameters.Delete (0)
    Next ni
    
    
    nExe = 0
    
    '<< DELETE : 시수코드 삭제 >>
    sStr = ""
    sStr = sStr & "  DELETE "
    sStr = sStr & "    FROM SDTCR01TB "
    sStr = sStr & "   WHERE ACID   = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "     AND SISUCD = '" & Trim(txtSisuCD.Text) & "'"
            
    '>> 기존 파라미터가 남아 있으면 메모리에서 삭제함.
    For ni = 0 To DBCmd.Parameters.Count - 1 Step 1
        DBCmd.Parameters.Delete (0)
    Next ni
            
'    '>> ACID
'    sTmp = Trim(basModule.SchCD)
'    nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'        Set DBParam = DBCmd.CreateParameter("LSNCD", adVarChar, adParamInput, nLength, Trim(sTmp)): DBCmd.Parameters.Append DBParam
'
'    '>> 시수코드
'    sTmp = Trim(sSisuCD)
'    nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'        Set DBParam = DBCmd.CreateParameter("LSNCD", adVarChar, adParamInput, nLength, Trim(sTmp)): DBCmd.Parameters.Append DBParam
            
    DBCmd.CommandText = sStr
    DBCmd.CommandType = adCmdText
    DBCmd.CommandTimeout = 30
    
    DBCmd.Execute nExe, , -1
            
    Do While basDataBase.DBConn.State And adStateExecuting
        DoEvents
    Loop
    
    If nExe = 1 Then
        basDataBase.DBConn.CommitTrans
        With sprSubj
            For nRow = 1 To .MaxRows Step 1
                .Row = nRow
                .Col = 1
                
                If StrComp(Trim(txtSisuCD.Text), Trim(.Text), vbTextCompare) = 0 Then
                    .Row = nRow
                    
                    .DeleteRows .Row, 1
                    .MaxRows = .MaxRows - 1
                    
                End If
            Next nRow
        End With
        
        MsgBox "삭제하였습니다.", vbInformation + vbOKOnly, "강사 및 과목내역 삭제"
        
        Call cmdNewSisu_Click
        
    Else
        basDataBase.DBConn.RollbackTrans
        MsgBox "에러가 발생하였습니다.", vbCritical + vbOKOnly, "강사 및 과목내역 삭제"
    End If
        
    Set DBCmd = Nothing
    Set DBParam = Nothing
    
    Exit Sub
    
ErrStmt:
    basDataBase.DBConn.RollbackTrans
    
    MsgBox "삭제중 에러가 발생하였습니다.", vbCritical + vbOKOnly, "강사 및 과목내역 삭제"
    On Error GoTo 0
    
    Set DBCmd = Nothing
    Set DBParam = Nothing
End Sub



'>> 강사 및 과목내역 저장
Private Sub cmdSaveSisu_Click()

    Dim sStr        As String
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    
    Dim ni          As Long
    
    Dim sTmp        As String
    Dim nLength     As Long
    Dim nExe        As Long
    
    Dim sSisuCD     As String
    Dim sExecute    As String
    
    Dim nRow        As Long
    
    If Trim(txtTcrNM.Text) = "" Then
        MsgBox "강사명이 없습니다.", vbExclamation + vbOKOnly, "강사 및 과목내역 저장"
        Exit Sub
    End If
    If Trim(txtSubjNM.Text) = "" Then
        MsgBox "과목명이 없습니다.", vbExclamation + vbOKOnly, "강사 및 과목내역 저장"
        Exit Sub
    End If
    
    
    On Error GoTo ErrStmt
    
    basDataBase.DBConn.BeginTrans
    
    Set DBCmd = New ADODB.Command
    Set DBParam = New ADODB.Parameter
    
    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
                
    '>> 기존 파라미터가 남아 있으면 메모리에서 삭제함.
    For ni = 0 To DBCmd.Parameters.Count - 1 Step 1
        DBCmd.Parameters.Delete (0)
    Next ni
    
    
    nExe = 0
    
    '<< INSERT : 시수코드 생성 >>
        If Trim(txtSisuCD.Text) = "" Then
            sStr = ""
            sStr = sStr & "  SELECT MAX(CD) AS CD"
            sStr = sStr & "    FROM (SELECT SISUCD + 1 AS CD"
            sStr = sStr & "            From SDTCR01TB"
            sStr = sStr & "           WHERE ACID = '" & Trim(basModule.SchCD) & "'"
            sStr = sStr & "          Union All"
            sStr = sStr & "          SELECT 1 AS CD"
            sStr = sStr & "            From DUAL"
            sStr = sStr & "          ) "
            
            
            Set DBCmd = New ADODB.Command
            Set DBRec = New ADODB.Recordset
            Set DBParam = New ADODB.Parameter
            
            DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
            DBCmd.CommandText = sStr
            DBCmd.CommandType = adCmdText
            DBCmd.CommandTimeout = 30
            
            '>> 기존 파라미터가 남아 있으면 메모리에서 삭제함.
            For ni = 0 To DBCmd.Parameters.Count - 1 Step 1
                DBCmd.Parameters.Delete (0)
            Next ni
            
            DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic형태로 열게되면 record count를 할 수 없음.
            Do While DBRec.State And adStateExecuting
                DoEvents
            Loop
            
            With DBRec
                If .RecordCount > 0 Then
                    .MoveFirst
                    
                    If IsNull(.Fields("CD")) = False Then
                        sSisuCD = Trim(.Fields("CD"))
                    Else
                        sSisuCD = "1"
                    End If
                End If
            End With
                
            Set DBRec = Nothing
    
            
            sStr = ""
            sStr = sStr & "  INSERT INTO SDTCR01TB ( ACID, SISUCD, TCRNM, TCRGBN, SUBJNM ) "
            sStr = sStr & "  VALUES ( ?, ?, ?, ?, ? )"
            
            
            '>> 기존 파라미터가 남아 있으면 메모리에서 삭제함.
            For ni = 0 To DBCmd.Parameters.Count - 1 Step 1
                DBCmd.Parameters.Delete (0)
            Next ni
            
            '>> ACID
            sTmp = Trim(basModule.SchCD)
            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
                Set DBParam = DBCmd.CreateParameter("LSNCD", adVarChar, adParamInput, nLength, Trim(sTmp)): DBCmd.Parameters.Append DBParam
                
            '>> 시수코드
            sTmp = Trim(sSisuCD)
            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
                Set DBParam = DBCmd.CreateParameter("LSNCD", adVarChar, adParamInput, nLength, Trim(sTmp)): DBCmd.Parameters.Append DBParam
                
            '>> 선생님 명
            sTmp = Trim(txtTcrNM.Text)
            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
                Set DBParam = DBCmd.CreateParameter("LSNNM", adVarChar, adParamInput, nLength, Trim(sTmp)): DBCmd.Parameters.Append DBParam
                
            '>> 과목구분
            sTmp = Trim(Right(cboTcrGbn.Text, 30))
            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
                Set DBParam = DBCmd.CreateParameter("KAEYOL", adChar, adParamInput, nLength, Trim(sTmp)): DBCmd.Parameters.Append DBParam
                
            '>> 과목명
            sTmp = Trim(txtSubjNM.Text)
            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
                Set DBParam = DBCmd.CreateParameter("KAEYOL", adChar, adParamInput, nLength, Trim(sTmp)): DBCmd.Parameters.Append DBParam
            
            DBCmd.CommandText = sStr
            DBCmd.CommandType = adCmdText
            DBCmd.CommandTimeout = 30
            
            DBCmd.Execute nExe, , -1
                    
            Do While basDataBase.DBConn.State And adStateExecuting
                DoEvents
            Loop
            
            If nExe = 1 Then
                sExecute = "INSERT"
                basDataBase.DBConn.CommitTrans
            Else
                sExecute = ""
                basDataBase.DBConn.RollbackTrans
            End If
            
    '<< UPDATE : 이미 조회된 코드로 등록 >>
        Else
            
            sSisuCD = Trim(txtSisuCD.Text)
            
            
            sStr = ""
            sStr = sStr & "  UPDATE SDTCR01TB "
            sStr = sStr & "     SET TCRNM  = '" & Trim(txtTcrNM.Text) & "',"
            sStr = sStr & "         TCRGBN = '" & Trim(Right(cboTcrGbn.Text, 30)) & "',"
            sStr = sStr & "         SUBJNM = '" & Trim(txtSubjNM.Text) & "'"
            sStr = sStr & "   WHERE ACID   = '" & Trim(basModule.SchCD) & "'"
            sStr = sStr & "     AND SISUCD = '" & Trim(txtSisuCD.Text) & "'"
            
            '>> 기존 파라미터가 남아 있으면 메모리에서 삭제함.
            For ni = 0 To DBCmd.Parameters.Count - 1 Step 1
                DBCmd.Parameters.Delete (0)
            Next ni
            
'            '>> 선생님 명
'            sTmp = Trim(txtTcrNM.Text)
'            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'                Set DBParam = DBCmd.CreateParameter("LSNNM", adVarChar, adParamInput, nLength, Trim(sTmp)): DBCmd.Parameters.Append DBParam
'            '>> 과목구분
'            sTmp = Trim(Right(cboTcrGbn.Text, 30))
'            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'                Set DBParam = DBCmd.CreateParameter("KAEYOL", adChar, adParamInput, nLength, Trim(sTmp)): DBCmd.Parameters.Append DBParam
'            '>> 과목명
'            sTmp = Trim(txtSubjNM.Text)
'            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'                Set DBParam = DBCmd.CreateParameter("KAEYOL", adChar, adParamInput, nLength, Trim(sTmp)): DBCmd.Parameters.Append DBParam
'            '>> ACID
'            sTmp = Trim(basModule.SchCD)
'            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'                Set DBParam = DBCmd.CreateParameter("LSNCD", adVarChar, adParamInput, nLength, Trim(sTmp)): DBCmd.Parameters.Append DBParam
'            '>> 시수코드
'            sTmp = Trim(sSisuCD)
'            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'                Set DBParam = DBCmd.CreateParameter("LSNCD", adVarChar, adParamInput, nLength, Trim(sTmp)): DBCmd.Parameters.Append DBParam
                
            DBCmd.CommandText = sStr
            DBCmd.CommandType = adCmdText
            DBCmd.CommandTimeout = 30
            
            DBCmd.Execute nExe, , -1
                    
            Do While basDataBase.DBConn.State And adStateExecuting
                DoEvents
            Loop
            
            If nExe = 1 Then
                sExecute = "UPDATE"
                basDataBase.DBConn.CommitTrans
            Else
                sExecute = ""
                basDataBase.DBConn.RollbackTrans
            End If
            
        End If
        
    
        With sprSubj
            Select Case sExecute
                Case "INSERT"
                    .MaxRows = .MaxRows + 1
                    .InsertRows 1, 1
                    .Row = 1
                Case "UPDATE"
                    For nRow = 1 To .MaxRows Step 1
                        .Row = nRow
                        .Col = 1
                        
                        If StrComp(sSisuCD, Trim(.Text), vbTextCompare) = 0 Then
                            .Row = nRow
                            Exit For
                        End If
                    Next nRow
            End Select
            
            .Col = 1:           sTmp = sSisuCD:                             Call basFunction.Set_SprType_Text(sprSubj, "center", "left", basFunction.LenKor(sTmp), sTmp)
            .Col = .Col + 1:    sTmp = Trim(txtTcrNM.Text):                 Call basFunction.Set_SprType_Text(sprSubj, "center", "left", basFunction.LenKor(sTmp), sTmp)
            .Col = .Col + 1:    sTmp = Trim(Right(cboTcrGbn.Text, 30)):     Call basFunction.Set_SprType_Text(sprSubj, "CENTER", "LEFT", basFunction.LenKor(sTmp), sTmp)
            
                Select Case sTmp
                    Case "10"
                        .Col = .Col + 1
                        sTmp = "언어":      Call basFunction.Set_SprType_Text(sprSubj, "CENTER", "LEFT", basFunction.LenKor(sTmp), sTmp)
                    Case "20"
                        .Col = .Col + 1
                        sTmp = "수리":      Call basFunction.Set_SprType_Text(sprSubj, "CENTER", "LEFT", basFunction.LenKor(sTmp), sTmp)
                    Case "30"
                        .Col = .Col + 1
                        sTmp = "Eng":       Call basFunction.Set_SprType_Text(sprSubj, "CENTER", "LEFT", basFunction.LenKor(sTmp), sTmp)
                    Case "40"
                        .Col = .Col + 1
                        sTmp = "사탐":      Call basFunction.Set_SprType_Text(sprSubj, "CENTER", "LEFT", basFunction.LenKor(sTmp), sTmp)
                    Case "50"
                        .Col = .Col + 1
                        sTmp = "과탐":      Call basFunction.Set_SprType_Text(sprSubj, "CENTER", "LEFT", basFunction.LenKor(sTmp), sTmp)
                End Select
                
            .Col = .Col + 1:    sTmp = Trim(txtSubjNM.Text):                Call basFunction.Set_SprType_Text(sprSubj, "CENTER", "LEFT", basFunction.LenKor(sTmp), sTmp)
            
            
            
        End With
    
    Call cmdNewSisu_Click
    
    Set DBCmd = Nothing
    Set DBParam = Nothing
    
    Exit Sub
    
ErrStmt:
    
    basDataBase.DBConn.RollbackTrans
    
    Set DBCmd = Nothing
    Set DBParam = Nothing
    
End Sub

Private Sub cmdNewSisu_Click()
    txtSisuCD.Text = ""
    txtTcrNM.Text = ""
    txtSubjNM.Text = ""

End Sub



'>> 강사 및 과목내역 조회
Private Sub cmdFindSisu_Click()
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    
    Dim nLength     As Long
    Dim sStr        As String
    Dim ni          As Integer
    Dim nRec        As Long
    
    Dim sTmp        As String
    
    sprSubj.MaxRows = 0
    
    On Error GoTo ErrStmt
    
    sStr = ""
    sStr = sStr & "  SELECT ACID, SISUCD, TCRNM,"
    sStr = sStr & "         TCRGBN,"
    sStr = sStr & "         DECODE(TCRGBN,10,'언어',"
    sStr = sStr & "                       20,'수리',"
    sStr = sStr & "                       30,'ENG' ,"
    sStr = sStr & "                       40,'사탐',"
    sStr = sStr & "                       50,'과탐') TCRGBN_NM,"
    sStr = sStr & "         SUBJNM"
    sStr = sStr & "    FROM SDTCR01TB"
    sStr = sStr & "   WHERE ACID = '" & Trim(basModule.SchCD) & "'"
    
    If Trim(txtTcrNM.Text) > " " Then
        sStr = sStr & " AND TCRNM LIKE '" & Trim(txtTcrNM.Text) & "'"
    End If
    If Trim(txtSubjNM.Text) > " " Then
        sStr = sStr & " AND SUBJNM LIKE '" & Trim(txtSubjNM.Text) & "'"
    End If
    sStr = sStr & "   ORDER BY TCRNM "
    
    Set DBCmd = New ADODB.Command
    Set DBRec = New ADODB.Recordset
    Set DBParam = New ADODB.Parameter
    
    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
    DBCmd.CommandText = sStr
    DBCmd.CommandType = adCmdText
    DBCmd.CommandTimeout = 30
    
    '>> 기존 파라미터가 남아 있으면 메모리에서 삭제함.
    For ni = 0 To DBCmd.Parameters.Count - 1 Step 1
        DBCmd.Parameters.Delete (0)
    Next ni
        
'    ' ACID
'        sTmp = Trim(basModule.SchCD)
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
'    ' TCRNM
'        If Trim(txtTcrNM.Text) > " " Then
'            sTmp = Trim(txtTcrNM.Text) & "%"
'            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'                Set DBParam = DBCmd.CreateParameter("LSNTYPE", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
'        End If
'
'        If Trim(txtSubjNM.Text) > " " Then
'            sTmp = Trim(txtSubjNM.Text) & "%"
'            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'                Set DBParam = DBCmd.CreateParameter("LSNTYPE", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
'        End If
    
    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic형태로 열게되면 record count를 할 수 없음.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop
    
    With DBRec
        
        If .RecordCount > 0 Then
            .MoveFirst
            
            For nRec = 1 To .RecordCount Step 1
                sprSubj.MaxRows = sprSubj.MaxRows + 1
                sprSubj.Row = sprSubj.MaxRows
                
                sprSubj.Col = 1
                    sTmp = " ":  If IsNull(.Fields("SISUCD")) = False Then sTmp = Trim(.Fields("SISUCD"))
                        Call basFunction.Set_SprType_Text(sprSubj, "CENTER", "LEFT", basFunction.LenKor(sTmp), Trim(sTmp))
                sprSubj.Col = sprSubj.Col + 1
                    sTmp = " ":  If IsNull(.Fields("TCRNM")) = False Then sTmp = Trim(.Fields("TCRNM"))
                        Call basFunction.Set_SprType_Text(sprSubj, "CENTER", "LEFT", basFunction.LenKor(sTmp), Trim(sTmp))
                sprSubj.Col = sprSubj.Col + 1
                    sTmp = " ":  If IsNull(.Fields("TCRGBN")) = False Then sTmp = Trim(.Fields("TCRGBN"))
                        Call basFunction.Set_SprType_Text(sprSubj, "CENTER", "LEFT", basFunction.LenKor(sTmp), Trim(sTmp))
                sprSubj.Col = sprSubj.Col + 1
                    sTmp = " ":  If IsNull(.Fields("TCRGBN_NM")) = False Then sTmp = Trim(.Fields("TCRGBN_NM"))
                        Call basFunction.Set_SprType_Text(sprSubj, "CENTER", "LEFT", basFunction.LenKor(sTmp), Trim(sTmp))
                sprSubj.Col = sprSubj.Col + 1
                    sTmp = " ":  If IsNull(.Fields("SUBJNM")) = False Then sTmp = Trim(.Fields("SUBJNM"))
                        Call basFunction.Set_SprType_Text(sprSubj, "CENTER", "LEFT", basFunction.LenKor(sTmp), Trim(sTmp))
                
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
    MsgBox "강사 및 과목내역 조회시 에러가 발생하였습니다.", vbCritical + vbOKOnly, "강사 및 과목내역 조회"
End Sub






'##########################################################################################################




'##########################################################################################################
Private Sub cmdFindTmr_Click()
    Dim nCol        As Long
    Dim nColChk     As Long
    
    
    sprTmr.MaxRows = 0
    sprTmr.MaxCols = 0
    
    sprTmr.Col = 0:   sprTmr.ColHidden = False
    sprTmr.Row = 0:   sprTmr.RowHidden = False
    
    sprTmr.RowHeaderCols = 1
    sprTmr.ColHeaderRows = 1
    
    ReDim uSisu_Data(0) As tSisu_Data                           '<< 초기화
    
    Call Display_SprTmr_Col_SpreadHeader
    If sprTmr.RowHeaderCols > 2 Then
        Call Display_SprTmr_Row_SpreadHeader
        
        If sprTmr.ColHeaderRows <= 2 Then
            sprTmr.MaxCols = 0
            sprTmr.MaxRows = 0
    
            sprTmr.ColHeaderRows = 1
            sprTmr.RowHeaderCols = 1
        Else
            Call Construct_Spread_Sisu_Data(sprTmr.MaxRows, sprTmr.MaxCols)
            
            If sprTmr.ColHeaderRows = 3 Then sprTmr.Row = SpreadHeader + 1:         sprTmr.RowHidden = True
            sprTmr.Col = sprTmr.MaxCols:                                            sprTmr.ColHidden = True
            
            
            sprTmr.Row = SpreadHeader
            sprTmr.Col = SpreadHeader
            
            If sprTmr.ColHidden = False Then
                sprTmr.ColHidden = True
            End If
            
            
            '## 데이터 넣기
            Call Find_input_SisuData
            
        End If
    End If
    
End Sub


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
    Dim sSisuCD     As String
    Dim sLsnCD      As String
    
    Dim nTmp        As Long
    
    On Error GoTo ErrStmt
    
    With sprTmr
        If .MaxRows = 0 Then Exit Sub
        If .MaxCols = 0 Then Exit Sub
        
            
        sStr = ""
        sStr = sStr & "  SELECT ACID, SISUCD, LSNCD, SISU "
        sStr = sStr & "    FROM SDTCR11TB "
        sStr = sStr & "   WHERE ACID = '" & Trim(basModule.SchCD) & "'"
        
        Set DBCmd = New ADODB.Command
        Set DBRec = New ADODB.Recordset
        Set DBParam = New ADODB.Parameter
        
        DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
        DBCmd.CommandText = sStr
        DBCmd.CommandType = adCmdText
        DBCmd.CommandTimeout = 30
        
        '>> 기존 파라미터가 남아 있으면 메모리에서 삭제함.
        For ni = 0 To DBCmd.Parameters.Count - 1 Step 1
            DBCmd.Parameters.Delete (0)
        Next ni
            
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
                    
                    If IsNull(.Fields("SISU")) = False Then
                        If Trim(.Fields("SISU")) <> "0" Then
                            uSisu_Data(nRec).ACID = Trim(.Fields("ACID"))
                            uSisu_Data(nRec).SISUCD = Trim(.Fields("SISUCD"))
                            uSisu_Data(nRec).LSNCD = Trim(.Fields("LSNCD"))
                            uSisu_Data(nRec).SISU = CLng(.Fields("SISU"))
                        End If
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
                .Row = nRow:    .Col = SpreadHeader:            sSisuCD = Trim(.Text)
                
                For nCol = 1 To (.MaxCols - 1) Step 1
                    .Col = nCol:    .Row = SpreadHeader + 1:    sLsnCD = Trim(.Text)
                    
                    For nRec = 1 To UBound(uSisu_Data) Step 1
                        If StrComp(uSisu_Data(nRec).SISUCD, sSisuCD, vbTextCompare) = 0 And _
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
    
    If Me.Tag = "LOAD" Then Exit Sub
    
    sprSubj.MaxRows = 0
    
    On Error GoTo ErrStmt
    
    sStr = ""
    sStr = sStr & "  SELECT DECODE(KAEYOL,'01','인문',"
    sStr = sStr & "                       '02','자연',"
    sStr = sStr & "                       '03','예체') KAEYOL,"
    sStr = sStr & "         LSNCD , LSNNM"
    sStr = sStr & "    From SDLSN01TB "
    sStr = sStr & "   WHERE ACID   = '" & Trim(basModule.SchCD) & "'"
    If Trim(Right(cboKaeyol.Text, 30)) <> "ALL" Then
        sStr = sStr & " AND KAEYOL = '" & Trim(Right(cboKaeyol.Text, 30)) & "'"
    End If
    sStr = sStr & "   ORDER BY KAEYOL, LSNNM"

    Set DBCmd = New ADODB.Command
    Set DBRec = New ADODB.Recordset
    Set DBParam = New ADODB.Parameter
    
    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
    DBCmd.CommandText = sStr
    DBCmd.CommandType = adCmdText
    DBCmd.CommandTimeout = 30
    
    '>> 기존 파라미터가 남아 있으면 메모리에서 삭제함.
    For ni = 0 To DBCmd.Parameters.Count - 1 Step 1
        DBCmd.Parameters.Delete (0)
    Next ni
        
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
            sprTmr.ColHeaderRows = 3
            
            .MoveFirst
            
            For nRec = 1 To .RecordCount Step 1
                sprTmr.Col = nRec
                
                sprTmr.Row = SpreadHeader:      sTmp = "":  If IsNull(.Fields("KAEYOL")) = False Then sTmp = Trim(.Fields("KAEYOL"))
                    sprTmr.Text = sTmp
                sprTmr.Row = SpreadHeader + 1:  sTmp = "":  If IsNull(.Fields("LSNCD")) = False Then sTmp = Trim(.Fields("LSNCD"))
                    sprTmr.Text = sTmp:     sprTmr.ColWidth(sprTmr.Col) = 6
                sprTmr.Row = SpreadHeader + 2:  sTmp = "":  If IsNull(.Fields("LSNNM")) = False Then sTmp = Trim(.Fields("LSNNM"))
                    sprTmr.Text = sTmp:     sprTmr.ColWidth(sprTmr.Col) = 6
                
                .MoveNext
            Next nRec
            
            sprTmr.Row = SpreadHeader + 2
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
    
    sprSubj.MaxRows = 0
    
    On Error GoTo ErrStmt
    
    sStr = ""
    sStr = sStr & "  SELECT SISUCD, TCRNM, SUBJNM"
    sStr = sStr & "    From SDTCR01TB"
    sStr = sStr & "   WHERE ACID   = '" & Trim(basModule.SchCD) & "'"
    If Trim(Right(cboFindTcrGbn.Text, 30)) <> "ALL" Then
        sStr = sStr & "     AND TCRGBN = '" & Trim(Right(cboFindTcrGbn.Text, 30)) & "'"
    End If
    sStr = sStr & "   ORDER BY TCRNM"
    
    Set DBCmd = New ADODB.Command
    Set DBRec = New ADODB.Recordset
    Set DBParam = New ADODB.Parameter
    
    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
    DBCmd.CommandText = sStr
    DBCmd.CommandType = adCmdText
    DBCmd.CommandTimeout = 30
    
    '>> 기존 파라미터가 남아 있으면 메모리에서 삭제함.
    For ni = 0 To DBCmd.Parameters.Count - 1 Step 1
        DBCmd.Parameters.Delete (0)
    Next ni
        
'    ' ACID
'        sTmp = Trim(basModule.SchCD)
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
'        If Trim(Right(cboFindTcrGbn.Text, 30)) <> "ALL" Then
'    ' TCRNM
'            sTmp = Trim(Right(cboFindTcrGbn.Text, 30))
'            nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'                Set DBParam = DBCmd.CreateParameter("LSNTYPE", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
'
'        End If
    
    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic형태로 열게되면 record count를 할 수 없음.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop
    
    With DBRec
        
        If .RecordCount > 0 Then
        
            sprTmr.MaxRows = .RecordCount + 1
            sprTmr.RowHeaderCols = 3
            
            .MoveFirst
            
            
            sprTmr.Row = 1
            sprTmr.Col = SpreadHeader + 2:  sTmp = "소 계"
                sprTmr.Text = sTmp
                sprTmr.RowHeight(sprTmr.Row) = 14             '<< 처음 행 : 합계처리
            
            
            For nRec = 1 To .RecordCount Step 1
                sprTmr.Row = nRec + 1
                
                sprTmr.Col = SpreadHeader:      sTmp = "":  If IsNull(.Fields("SISUCD")) = False Then sTmp = Trim(.Fields("SISUCD"))
                    sprTmr.Text = sTmp
                sprTmr.Col = SpreadHeader + 1:  sTmp = "":  If IsNull(.Fields("TCRNM")) = False Then sTmp = Trim(.Fields("TCRNM"))
                    sprTmr.Text = sTmp:     sprTmr.ColWidth(sprTmr.Col) = 7
                sprTmr.Col = SpreadHeader + 2:  sTmp = "":  If IsNull(.Fields("SUBJNM")) = False Then sTmp = Trim(.Fields("SUBJNM"))
                    sprTmr.Text = sTmp:     sprTmr.ColWidth(sprTmr.Col) = 9
                    sprTmr.TypeHAlign = TypeHAlignLeft
                    sprTmr.TypeVAlign = TypeVAlignCenter
                
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
        
    End With
    
End Sub









'>> 등록하기
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
            MsgBox "시수를 넣으신 후 등록버튼을 클릭하세요.", vbExclamation + vbOKOnly, "강사 및 과목내역 등록"
            Exit Sub
        End If
        
            
        '## 데이터 저장
        Call Save_Detail_Data
        
        
    End With
    
    Exit Sub
ErrStmt:
    On Error GoTo 0
    MsgBox "강사 및 과목내역 등록시 에러가 발생하였습니다.", vbCritical + vbOKOnly, "강사 및 과목내역 등록"
    
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
    
    Dim sSisuCD     As String           ' 시수코드 : header에 있음
    Dim sLsnCD      As String           ' 반코드 : header에 있음
    Dim nSisu       As Long             ' 시수
    
    Dim nTotExe     As Long             ' insert/update 되어질 것
    Dim nAddExe     As Long             '               처리된 결과 합
    Dim nExe        As Long             '               처리
    
    
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
            .Row = nRow:        .Col = SpreadHeader:        sSisuCD = Trim(.Text)                   '<< 시수코드 : header
            .Col = .MaxCols
            
            If .Value = 1 Then
                For nCol = 1 To (.MaxCols - 2) Step 1
                    .Col = nCol:                            .Row = SpreadHeader + 1:                sLsnCD = Trim(.Text)    '<< 반코드
                    
                    .Row = nRow
                    .Col = nCol
                        If Trim(.Text) > " " Then
                            'If Trim(.Text) <> "0" Then
                            
                                nTotExe = nTotExe + 1       '<< 작업
                                nSisu = .Value
                                
                                
                                '## SELECT
                                sStr = ""
                                sStr = sStr & " SELECT ACID, SISUCD, LSNCD, SISU "
                                sStr = sStr & "   FROM SDTCR11TB "
                                sStr = sStr & "  WHERE ACID   = '" & Trim(basModule.SchCD) & "'"
                                sStr = sStr & "    AND SISUCD =  " & sSisuCD
                                sStr = sStr & "    AND LSNCD  = '" & sLsnCD & "'"
                                
                                DBCmd.CommandText = sStr
                                DBCmd.CommandType = adCmdText
                                DBCmd.CommandTimeout = 30
                    
                                '>> 기존 파라미터가 남아 있으면 메모리에서 삭제함.
                                For ni = 0 To DBCmd.Parameters.Count - 1 Step 1
                                    DBCmd.Parameters.Delete (0)
                                Next ni
                                        
'                                ' ACID
'                                    sTmp = Trim(basModule.SchCD)
'                                    nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'                                        Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
'                                ' SISUCD
'                                    nTmp = CLng(sSisuCD)
'                                        Set DBParam = DBCmd.CreateParameter("SISUCD", adDouble, adParamInput, , nTmp):   DBCmd.Parameters.Append DBParam
'                                ' LSNCD
'                                    sTmp = Trim(sLsnCD)
'                                    nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'                                        Set DBParam = DBCmd.CreateParameter("LSNCD", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
                                
                                DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic형태로 열게되면 record count를 할 수 없음.
                                Do While DBRec.State And adStateExecuting
                                    DoEvents
                                Loop
                                
                                Select Case DBRec.RecordCount
                                    Case 0
                        '< insert >
                                        sStr = ""
                                        sStr = sStr & "  INSERT INTO SDTCR11TB (ACID, SISUCD, LSNCD, SISU)"
                                        sStr = sStr & "  VALUES ( "
                                        sStr = sStr & "     '" & Trim(basModule.SchCD) & "', "
                                        sStr = sStr & "      " & sSisuCD & ", "
                                        sStr = sStr & "     '" & sLsnCD & "', "
                                        sStr = sStr & "      " & Trim(CStr(nSisu))
                                        sStr = sStr & "  ) "
                                        
                                        
                                        '>> 기존 파라미터가 남아 있으면 메모리에서 삭제함.
                                        For ni = 0 To DBCmd.Parameters.Count - 1 Step 1
                                            DBCmd.Parameters.Delete (0)
                                        Next ni
                                        
'                                    ' ACID
'                                        sTmp = Trim(basModule.SchCD)
'                                        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'                                            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
'                                    ' SISUCD
'                                        nTmp = CLng(sSisuCD)
'                                            Set DBParam = DBCmd.CreateParameter("SISUCD", adDouble, adParamInput, , nTmp):   DBCmd.Parameters.Append DBParam
'                                    ' LSNCD
'                                        sTmp = Trim(sLsnCD)
'                                        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'                                            Set DBParam = DBCmd.CreateParameter("LSNCD", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
'                                    ' SISU
'                                        nTmp = nSisu
'                                            Set DBParam = DBCmd.CreateParameter("SISU", adDouble, adParamInput, , nTmp):   DBCmd.Parameters.Append DBParam
                                            
                                    Case Else
                        '< update >
                                        sStr = ""
                                        sStr = sStr & "  UPDATE SDTCR11TB"
                                        sStr = sStr & "     SET SISU   =  " & Trim(CStr(nSisu))
                                        sStr = sStr & "   WHERE ACID   = '" & Trim(basModule.SchCD) & "'"
                                        sStr = sStr & "     AND SISUCD =  " & sSisuCD
                                        sStr = sStr & "     AND LSNCD  = '" & sLsnCD & "'"
                                        
                                        '>> 기존 파라미터가 남아 있으면 메모리에서 삭제함.
                                        For ni = 0 To DBCmd.Parameters.Count - 1 Step 1
                                            DBCmd.Parameters.Delete (0)
                                        Next ni
                            
'                                    ' SISU
'                                        nTmp = nSisu
'                                            Set DBParam = DBCmd.CreateParameter("SISU", adDouble, adParamInput, , nTmp):   DBCmd.Parameters.Append DBParam
'                                    ' ACID
'                                        sTmp = Trim(basModule.SchCD)
'                                        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'                                            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
'                                    ' SISUCD
'                                        nTmp = CLng(sSisuCD)
'                                            Set DBParam = DBCmd.CreateParameter("SISUCD", adDouble, adParamInput, , nTmp):   DBCmd.Parameters.Append DBParam
'                                    ' LSNCD
'                                        sTmp = Trim(sLsnCD)
'                                        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'                                            Set DBParam = DBCmd.CreateParameter("LSNCD", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
                            
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
                                
                                DBRec.Close
                                
                            'End If
                        End If
                Next nCol
                
            End If
        Next nRow
    End With
    
    
    If nTotExe = nAddExe Then
        basDataBase.DBConn.CommitTrans
        MsgBox "강사 및 과목내역 등록하였습니다.", vbInformation + vbOKOnly, "강사 및 과목내역 등록"
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
    
    On Error GoTo 0
    
    MsgBox "강사 및 과목내역 등록시 에러가 발생하였습니다.", vbCritical + vbOKOnly, "강사 및 과목내역 등록"
    
End Sub



































