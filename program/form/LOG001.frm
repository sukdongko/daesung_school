VERSION 5.00
Begin VB.Form LOG001 
   BorderStyle     =   1  '단일 고정
   Caption         =   "대성학원 입학사정. 반배정. 시간표 프로그램"
   ClientHeight    =   2520
   ClientLeft      =   7620
   ClientTop       =   6375
   ClientWidth     =   3990
   Icon            =   "LOG001.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   3990
   Begin VB.CommandButton cmdExit 
      Caption         =   "나가기"
      Height          =   435
      Left            =   2010
      TabIndex        =   3
      Top             =   1560
      Width           =   1065
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "들어가기"
      Height          =   435
      Left            =   690
      TabIndex        =   2
      Top             =   1560
      Width           =   1065
   End
   Begin VB.CommandButton cmdSchool 
      Caption         =   "학원 선택하기"
      Height          =   300
      Left            =   3360
      TabIndex        =   8
      Top             =   270
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.ComboBox cboSchool 
      Height          =   300
      Left            =   1260
      Style           =   2  '드롭다운 목록
      TabIndex        =   4
      Top             =   180
      Width           =   1875
   End
   Begin VB.TextBox txtPass 
      BorderStyle     =   0  '없음
      Height          =   300
      IMEMode         =   3  '사용 못함
      Left            =   1290
      PasswordChar    =   "*"
      TabIndex        =   1
      Text            =   "txtPass"
      Top             =   1035
      Width           =   1485
   End
   Begin VB.TextBox txtNM 
      BorderStyle     =   0  '없음
      Height          =   300
      IMEMode         =   10  '한글 
      Left            =   1290
      MaxLength       =   50
      TabIndex        =   0
      Text            =   "txtNM"
      Top             =   600
      Width           =   1485
   End
   Begin VB.Label lblShow 
      Alignment       =   1  '오른쪽 맞춤
      BackStyle       =   0  '투명
      Caption         =   "."
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   3840
      TabIndex        =   9
      Top             =   2340
      Width           =   135
   End
   Begin VB.Label lblSchool 
      Alignment       =   1  '오른쪽 맞춤
      BackStyle       =   0  '투명
      Caption         =   "학원선택"
      Height          =   210
      Left            =   240
      TabIndex        =   7
      Top             =   270
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  '오른쪽 맞춤
      BackStyle       =   0  '투명
      Caption         =   "비밀번호"
      Height          =   210
      Left            =   240
      TabIndex        =   6
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label4 
      Alignment       =   1  '오른쪽 맞춤
      BackStyle       =   0  '투명
      Caption         =   "사원"
      Height          =   210
      Left            =   240
      TabIndex        =   5
      Top             =   645
      Width           =   975
   End
End
Attribute VB_Name = "LOG001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################################################
'   시 스 템  명 : 대성학원 입학사정, 반배정 & 시간표 프로그램
'   서브시스템명 :
'   모   듈   명 : LOG001
'   모 듈  목 적 : LOGIN 처리
'
'   작   성   일 : 2007/08/20
'   작   성   자 : 유하균
' --------------------------------------------------------------------------------------------------------------
' --------------------------------------------------------------------------------------------------------------
'                 수     정     내     용
' --------------------------------------------------------------------------------------------------------------
'   1. 수정일 :
'   2. 내  용 :
'################################################################################################################


Option Explicit
Private sini_Path      As String    '>> 대성학원
Private Const sDebug = 1


Private Sub cmdExit_Click()
    Unload Me
End Sub



Private Sub Form_Load()
    Dim sData               As String * 255
    Dim sGbn                As String
    Dim nRtn                As Long
    
    Dim sTmp                As String
    Dim sDB_Tns_Location    As String
    Dim bFirstLogin         As Boolean
    
    bFirstLogin = False
    
    '## 프로그램 여러개를 띄울 수 있도록 함. 단, update는 안됨.
    If App.PrevInstance = False Then
        Call Update
    End If
    
    Me.KeyPreview = True
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Tag = "LOAD"
    
    
    txtNM.Text = ""
    txtPass.Text = ""
    
    If sDebug = 1 Then
        txtNM.Text = "ADMIN"
        txtPass.Text = "1"
    End If
    
    If sDebug = 0 Then
        cboSchool.Visible = False
        lblSchool.Visible = False
    End If
    
    basFunction.RemoveContextMenu txtNM
    basFunction.RemoveContextMenu txtPass
    
    With cboSchool
        .Clear
        .AddItem "노량진" & Space(30) & "N"
        .AddItem "강남" & Space(30) & "K"
        .AddItem "송파" & Space(30) & "S"
        .AddItem "송파 M" & Space(30) & "P"
        
        .AddItem "강남 M" & Space(30) & "M"
        .AddItem "주말법의대" & Space(30) & "W"
        .AddItem "야간법의대" & Space(30) & "Q"
        
        .AddItem "양재" & Space(30) & "J"
        .AddItem "부산" & Space(30) & "B"
        
        .ListIndex = 0
    End With
    
    ' 폼을 먼저 띄우고 INI파일 생성하자.
    
    '>> 프로그램 INI 파일
    sini_Path = App.Path & "\DAESUNG.INI"
    If Dir(sini_Path) = "" Then                                     '<< 파일이 없으면 생성
        Call Create_School_ini_File
        '파일이없으면 첫로그인으로 처리한다. 학원선택
        bFirstLogin = True
    End If
    
    '>>>>>>>>>>>>>>>>>>>>>> 프로그램 전역정보 세팅
    sGbn = "SCHOOL"
    sData = ""
    nRtn = basModule.GetPrivateProfileString(sGbn, "SCHOOL", "", sData, 255, sini_Path)         '>> 학교코드
    basModule.schcd = Trim(Replace(Left(Trim(sData), nRtn), Chr(0), "", 1, -1, vbTextCompare))
    If nRtn = 0 Then
        Call Create_School_ini_File
    End If
    
    
    Select Case Trim(basModule.schcd)
        Case "N"
            cboSchool.ListIndex = 0
        Case "K"
            cboSchool.ListIndex = 1
        Case "S"
            cboSchool.ListIndex = 2
        Case "P"
            cboSchool.ListIndex = 3
        Case "M"
            cboSchool.ListIndex = 4
            
        Case "W"
            cboSchool.ListIndex = 5
        Case "Q"
            cboSchool.ListIndex = 6
            
        Case "J"
            cboSchool.ListIndex = 7
        Case "B"
            cboSchool.ListIndex = 8
        
        Case Else
            cboSchool.ListIndex = 0
    End Select
    
    
    sData = ""
    nRtn = basModule.GetPrivateProfileString(sGbn, "SCHOOL_NM", "", sData, 255, sini_Path)      '>> 보내는 사람 코드
    basModule.SchNM = Trim(Replace(Left(Trim(sData), nRtn), Chr(0), "", 1, -1, vbTextCompare))
    
    
    If nRtn = 0 Then
        Call Create_School_ini_File
    End If
    
    sData = ""
    nRtn = basModule.GetPrivateProfileString(sGbn, "DB", "", sData, 255, sini_Path)             '>> DB접속
    basModule.connDB = Trim(Replace(Left(Trim(sData), nRtn), Chr(0), "", 1, -1, vbTextCompare))
    If nRtn = 0 Then
        Call Create_School_ini_File
    End If
    
    
    '각 학원에 따른 과목코드 정보 세팅
    Call basGwamok.setConstant
        
    
    
    '## 접속데이터 처리
    '## DB 접속 : 접속이 완료되면 => DBConn 에 connection 정보가 넣어집니다.
    If basDataBase.DataBase_Connection() = False Then
        MsgBox "접속 에러 관리자에게 문의 바랍니다"
        Exit Sub
    End If
            
    Me.Tag = ""
    
    
    
    If True = bFirstLogin Then
        cboSchool.Visible = True
        lblSchool.Visible = True
    End If
    
End Sub

Private Sub Create_School_ini_File()
    Dim sGbn        As String
    Dim nRtn        As Long
    
    basModule.schcd = "N"
    basModule.SchNM = "노량진"
    basModule.connDB = "MIMAC"
        
    sGbn = "SCHOOL"
    'nRtn = basModule.WritePrivateProfileString(sGbn, "PATH_ORACLE_TNS", basDataBase.TNS_Path1, sini_Path)                  '<< oracle tns 경로 - 앞으로 변경될수있다.
    nRtn = basModule.WritePrivateProfileString(sGbn, "SCHOOL", schcd, sini_Path)                  '<< 학원
    nRtn = basModule.WritePrivateProfileString(sGbn, "SCHOOL_NM", SchNM, sini_Path)          '<< 학원명
    nRtn = basModule.WritePrivateProfileString(sGbn, "DB", connDB, sini_Path)                  '<< DB 접속 - mimac 실서버, dev 개발서버
        
        
End Sub


Private Sub cboSchool_Click()
    If StrComp(Trim(Me.Tag), "LOAD", vbTextCompare) = 0 Then Exit Sub
    
    Call cmdSchool_Click
End Sub

'>> 학원선택
Private Sub cmdSchool_Click()
    Dim sGbn        As String
    Dim nRtn        As Long
    
    If StrComp(Trim(Me.Tag), "LOAD", vbTextCompare) = 0 Then Exit Sub
    
    Select Case Trim(Right(cboSchool.Text, 30))
        Case "N"
            If MsgBox("노량진 학원입니다." & vbCrLf & "맞습니까?", vbQuestion + vbYesNo, "학원선택") = vbNo Then Exit Sub
        Case "K"
            If MsgBox("강남 학원입니다." & vbCrLf & "맞습니까?", vbQuestion + vbYesNo, "학원선택") = vbNo Then Exit Sub
        Case "S"
            If MsgBox("송파 학원입니다." & vbCrLf & "맞습니까?", vbQuestion + vbYesNo, "학원선택") = vbNo Then Exit Sub
        Case "P"
            If MsgBox("송파마이맥 학원입니다." & vbCrLf & "맞습니까?", vbQuestion + vbYesNo, "학원선택") = vbNo Then Exit Sub
        Case "M"
            If MsgBox("강남마이맥 학원입니다." & vbCrLf & "맞습니까?", vbQuestion + vbYesNo, "학원선택") = vbNo Then Exit Sub
            
        Case "W"
            If MsgBox("주말법의대 학원입니다." & vbCrLf & "맞습니까?", vbQuestion + vbYesNo, "학원선택") = vbNo Then Exit Sub
        Case "Q"
            If MsgBox("야간법의대 학원입니다." & vbCrLf & "맞습니까?", vbQuestion + vbYesNo, "학원선택") = vbNo Then Exit Sub
            
        Case "J"
            If MsgBox("양재 학원입니다." & vbCrLf & "맞습니까?", vbQuestion + vbYesNo, "학원선택") = vbNo Then Exit Sub
        Case "B"
            If MsgBox("부산 학원입니다." & vbCrLf & "맞습니까?", vbQuestion + vbYesNo, "학원선택") = vbNo Then Exit Sub
            
    End Select
    
    sGbn = "SCHOOL"
    Select Case Trim(Right(cboSchool.Text, 30))
        Case "N"
            nRtn = basModule.WritePrivateProfileString(sGbn, "SCHOOL", "N", sini_Path)                  '<< 학원
            nRtn = basModule.WritePrivateProfileString(sGbn, "SCHOOL_NM", "노량진", sini_Path)          '<< 학원명
            
            schcd = "N":    SchNM = "노량진"
        Case "K"
            nRtn = basModule.WritePrivateProfileString(sGbn, "SCHOOL", "K", sini_Path)                  '<< 학원
            nRtn = basModule.WritePrivateProfileString(sGbn, "SCHOOL_NM", "강남", sini_Path)            '<< 학원명
            
            schcd = "K":    SchNM = "강남"
        Case "S"
            nRtn = basModule.WritePrivateProfileString(sGbn, "SCHOOL", "S", sini_Path)                  '<< 학원
            nRtn = basModule.WritePrivateProfileString(sGbn, "SCHOOL_NM", "송파", sini_Path)            '<< 학원명
            
            schcd = "S":    SchNM = "송파"
        Case "P"
            nRtn = basModule.WritePrivateProfileString(sGbn, "SCHOOL", "P", sini_Path)                  '<< 학원
            nRtn = basModule.WritePrivateProfileString(sGbn, "SCHOOL_NM", "송파 M", sini_Path)          '<< 학원명
            
            schcd = "P":    SchNM = "송파 M"
        Case "M"
            nRtn = basModule.WritePrivateProfileString(sGbn, "SCHOOL", "M", sini_Path)                  '<< 학원
            nRtn = basModule.WritePrivateProfileString(sGbn, "SCHOOL_NM", "강남 M", sini_Path)          '<< 학원명
            
            schcd = "M":    SchNM = "강남 M"
            
        Case "W"
            nRtn = basModule.WritePrivateProfileString(sGbn, "SCHOOL", "W", sini_Path)                  '<< 학원
            nRtn = basModule.WritePrivateProfileString(sGbn, "SCHOOL_NM", "주말법의대", sini_Path)      '<< 학원명
            
            schcd = "W":    SchNM = "주말법의대"
        Case "Q"
            nRtn = basModule.WritePrivateProfileString(sGbn, "SCHOOL", "Q", sini_Path)                  '<< 학원
            nRtn = basModule.WritePrivateProfileString(sGbn, "SCHOOL_NM", "야간법의대", sini_Path)      '<< 학원명
            
            schcd = "Q":    SchNM = "야간법의대"
            
        Case "J"
            nRtn = basModule.WritePrivateProfileString(sGbn, "SCHOOL", "J", sini_Path)                  '<< 학원
            nRtn = basModule.WritePrivateProfileString(sGbn, "SCHOOL_NM", "양재", sini_Path)        '<< 학원명
            
            schcd = "J":    SchNM = "양재"
            
        Case "B"
            nRtn = basModule.WritePrivateProfileString(sGbn, "SCHOOL", "B", sini_Path)                  '<< 학원
            nRtn = basModule.WritePrivateProfileString(sGbn, "SCHOOL_NM", "부산", sini_Path)            '<< 학원명
            
            schcd = "B":    SchNM = "부산"
            
    End Select
    
    MsgBox "완료하였습니다.", vbInformation + vbOKOnly, "학원선택"
    
End Sub



'>> 프로그램 사용여부
Private Sub cmdOK_Click()
    Dim sSql        As String
    Dim sTmp        As String
    
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim ni          As Long
    
    Dim bChk        As Boolean
    
    If Trim(txtNM.Text) = "" Then
        MsgBox "사원명을 넣으세요.", vbExclamation + vbOKOnly, "확인"
        Exit Sub
    End If
    
    If Trim(txtPass.Text) = "" Then
        MsgBox "비밀번호를 넣으세요.", vbExclamation + vbOKOnly, "확인"
        Exit Sub
    End If
    
    bChk = False
    
    '>> 회원확인
    If Trim(txtNM.Text) = "ADMIN" Then
        sSql = ""
        sSql = sSql & " SELECT EMPNO, EMPNM, PASSWD "
        sSql = sSql & "   FROM CLEMP01TB "
        sSql = sSql & "  WHERE EMPNM  = '" & Trim(txtNM.Text) & "' "
        sSql = sSql & "    AND PASSWD = '" & Trim(txtPass.Text) & "' "
    Else
        sSql = ""
        sSql = sSql & " SELECT EMPNO, EMPNM, PASSWD "
        sSql = sSql & "   FROM CLEMP01TB "
        sSql = sSql & "  WHERE ACID   = '" & Trim(basModule.schcd) & "' "
        sSql = sSql & "    AND EMPNM  = '" & Trim(txtNM.Text) & "' "
        sSql = sSql & "    AND PASSWD = '" & Trim(txtPass.Text) & "' "
    End If
    
    On Error GoTo ErrStmt
    
    Set DBCmd = New ADODB.Command
    Set DBRec = New ADODB.Recordset
    
    With DBCmd
        .ActiveConnection = basDataBase.DBConn
        .CommandTimeout = 60
        .CommandText = sSql
        .CommandType = adCmdText
        
    End With
    
    With DBRec
        .Open DBCmd, , adOpenStatic, adLockReadOnly, -1         '<< dynamic 형태로 열개되면 record count를 할 수 없음.
        Do While .State And adStateExecuting
            DoEvents
        Loop
        
        If .RecordCount = 1 Then
            If StrComp(Trim(txtNM.Text), .Fields("EMPNM"), vbTextCompare) = 0 And _
               StrComp(Trim(txtPass.Text), .Fields("PASSWD"), vbTextCompare) = 0 Then
                        
                bChk = True         '<< 확인 OK
                
                basModule.RegID = .Fields("EMPNO")
                
            End If
        Else
            MsgBox "담당자가 없습니다.", vbExclamation + vbOKOnly, "LOGIN"
        End If
    End With
    
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    If bChk = True Then
        Load MDI001
        MDI001.WindowState = 2
        MDI001.Show
        
        Unload LOG001
    End If
    
    Exit Sub
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    'MsgBox "프로그램 사용체크시 오류가 발생하였습니다." & vbCrLf & _
           Trim(CStr(Err.Number)) & ":" & Err.Description & vbCrLf & _
           basDataBase.DBConn, vbCritical + vbOKOnly, "LOGIN"
    
    MsgBox "프로그램 사용체크시 오류가 발생하였습니다.  " & vbCrLf & _
            "환경변수 Path에 오라클 경로를 확인하세요." & vbCrLf & _
           Trim(CStr(Err.Number)) & ":" & Err.Description & vbCrLf _
           , vbCritical + vbOKOnly, "LOGIN"
           
    On Error GoTo 0
End Sub



'>> liveupdate 처리
Private Sub Update()
    If Command = "no_update" Then
        Call basModule.SleepEx(1000, 0)
        Call RenameUpdateExe
    Else
        Call RunUpdateExe
    End If
End Sub

Private Sub RunUpdateExe()
    On Error GoTo EH
    Call Shell(App.Path & "\update.exe " & App.EXEName & ",대성학원 반배정 프로그램을", vbNormalFocus)
    End
EH:
    MsgBox "라이브업데이트가 누락되었습니다.", vbExclamation, "대성학원 반배정 프로그램"
End Sub

Sub RenameUpdateExe()
    On Error Resume Next

    
    Call Shell(Environ("COMSPEC") & " /c " & _
            "move /y " & Chr(34) & App.Path & "\update" & "\update_x.exe" & Chr(34) & " " & Chr(34) & App.Path & "\update.exe" & Chr(34), vbNormalFocus)

    Call Kill(App.Path & "\update_x.exe")
End Sub



Private Sub lblShow_Click()
    cboSchool.Visible = True
    lblSchool.Visible = True
End Sub

Private Sub txtPass_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            Call cmdOK_Click
            
    End Select
End Sub
