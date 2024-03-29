VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form TMR028 
   BackColor       =   &H00C0FFC0&
   Caption         =   "시간표 만들기 >> 이동수업 시간표 과목등록 CP"
   ClientHeight    =   6780
   ClientLeft      =   4200
   ClientTop       =   3540
   ClientWidth     =   13620
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   6780
   ScaleWidth      =   13620
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  '없음
      Caption         =   "Frame2"
      Height          =   6675
      Left            =   30
      TabIndex        =   17
      Top             =   30
      Width           =   13395
      Begin FPSpread.vaSpread sprExcel 
         Height          =   3765
         Left            =   1230
         TabIndex        =   21
         Top             =   2040
         Visible         =   0   'False
         Width           =   10455
         _Version        =   393216
         _ExtentX        =   18441
         _ExtentY        =   6641
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
         SpreadDesigner  =   "TMR028.frx":0000
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  '없음
         Caption         =   "Frame1"
         Height          =   435
         Left            =   30
         TabIndex        =   18
         Top             =   540
         Width           =   13335
         Begin VB.OptionButton optTamgu 
            BackColor       =   &H00FFFFFF&
            Caption         =   "과목내역"
            Height          =   240
            Index           =   0
            Left            =   0
            TabIndex        =   3
            Top             =   90
            Width           =   1125
         End
         Begin VB.OptionButton optTamgu 
            BackColor       =   &H00C0C0FF&
            Caption         =   "Option1"
            Height          =   240
            Index           =   1
            Left            =   1170
            TabIndex        =   4
            Top             =   90
            Width           =   1065
         End
         Begin VB.OptionButton optTamgu 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Option1"
            Height          =   240
            Index           =   2
            Left            =   2280
            TabIndex        =   5
            Top             =   90
            Width           =   1065
         End
         Begin VB.OptionButton optTamgu 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Option1"
            Height          =   240
            Index           =   3
            Left            =   3390
            TabIndex        =   6
            Top             =   90
            Width           =   1065
         End
         Begin VB.OptionButton optTamgu 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Option1"
            Height          =   240
            Index           =   4
            Left            =   4500
            TabIndex        =   7
            Top             =   90
            Width           =   1065
         End
         Begin VB.OptionButton optTamgu 
            BackColor       =   &H0000C0C0&
            Caption         =   "Option1"
            Height          =   240
            Index           =   5
            Left            =   5610
            TabIndex        =   8
            Top             =   90
            Width           =   1065
         End
         Begin VB.OptionButton optTamgu 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Option1"
            Height          =   240
            Index           =   6
            Left            =   6720
            TabIndex        =   9
            Top             =   90
            Width           =   1065
         End
         Begin VB.OptionButton optTamgu 
            BackColor       =   &H00FF80FF&
            Caption         =   "Option1"
            Height          =   240
            Index           =   7
            Left            =   7830
            TabIndex        =   10
            Top             =   90
            Width           =   1065
         End
         Begin VB.OptionButton optTamgu 
            BackColor       =   &H00FFFF00&
            Caption         =   "Option1"
            Height          =   240
            Index           =   8
            Left            =   8940
            TabIndex        =   11
            Top             =   90
            Width           =   1065
         End
         Begin VB.OptionButton optTamgu 
            BackColor       =   &H0000C000&
            Caption         =   "Option1"
            Height          =   240
            Index           =   9
            Left            =   10050
            TabIndex        =   12
            Top             =   90
            Width           =   1065
         End
         Begin VB.OptionButton optTamgu 
            BackColor       =   &H000000FF&
            Caption         =   "Option1"
            Height          =   240
            Index           =   10
            Left            =   11160
            TabIndex        =   13
            Top             =   90
            Width           =   1065
         End
         Begin VB.OptionButton optTamgu 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Option1"
            Height          =   240
            Index           =   11
            Left            =   12270
            TabIndex        =   14
            Top             =   90
            Width           =   1065
         End
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "반별 과목내역 조회하기 (&F)"
         Height          =   435
         Left            =   2940
         TabIndex        =   2
         Top             =   60
         Width           =   2775
      End
      Begin VB.ComboBox cboLsnType 
         Height          =   300
         Left            =   1500
         Style           =   2  '드롭다운 목록
         TabIndex        =   1
         Top             =   90
         Width           =   975
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "반별 과목내역 등록하기 (&S)"
         Height          =   525
         Left            =   10440
         TabIndex        =   16
         Top             =   5880
         Width           =   2655
      End
      Begin VB.ComboBox cboKaeyol 
         Height          =   300
         Left            =   180
         Style           =   2  '드롭다운 목록
         TabIndex        =   0
         Top             =   90
         Width           =   975
      End
      Begin FPSpread.vaSpread sprGwamok 
         Height          =   4635
         Left            =   30
         TabIndex        =   15
         Top             =   990
         Width           =   13335
         _Version        =   393216
         _ExtentX        =   23521
         _ExtentY        =   8176
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
         MaxRows         =   8
         ProcessTab      =   -1  'True
         SpreadDesigner  =   "TMR028.frx":01D4
      End
      Begin MSComDlg.CommonDialog dlgExcel 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Image imgExcel 
         Height          =   420
         Left            =   12960
         Picture         =   "TMR028.frx":32DD
         Stretch         =   -1  'True
         Top             =   60
         Width           =   390
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '투명
         Caption         =   "▷ 작업후 반드시【 반별 과목내역 등록하기 】    를 클릭하여 저장합니다."
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   8730
         TabIndex        =   20
         Top             =   90
         Width           =   4185
      End
      Begin VB.Label Label2 
         BackStyle       =   0  '투명
         Caption         =   $"TMR028.frx":371E
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   5880
         TabIndex        =   19
         Top             =   90
         Width           =   2805
      End
   End
End
Attribute VB_Name = "TMR028"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################################################
'   시 스 템  명 : 대성학원 입학사정, 반배정 & 시간표 프로그램
'   서브시스템명 :
'   모   듈   명 : TRM028
'   모 듈  목 적 :
'
'   작   성   일 : 2008/02/11
'   작   성   자 : 유하균
' --------------------------------------------------------------------------------------------------------------
' --------------------------------------------------------------------------------------------------------------
'                 수     정     내     용
' --------------------------------------------------------------------------------------------------------------
'   1. 수정일 :
'   2. 내  용 :
'################################################################################################################

Option Explicit

Private Sub Form_Activate()
    sprGwamok.SetFocus
    If sprGwamok.MaxCols > 1 Then sprGwamok.SetActiveCell 1, 1
    
End Sub

Private Sub Form_Load()
    Dim ni      As Long
    
    Me.Move 200, 900, 13600, 7100
    
    With sprGwamok
        .ShadowColor = basModule.ShadowColor2
        .ShadowDark = basModule.ShadowDark2
        .ShadowText = basModule.ShadowText2
        .GridColor = basModule.GridColor2
        .GrayAreaBackColor = basModule.GrayAreaBackColor2
        
        .MaxCols = 0
        .Row = SpreadHeader:        .RowHidden = True
        .Row = SpreadHeader + 1:    .RowHeight(.Row) = 16
        .Row = SpreadHeader + 2:    .RowHeight(.Row) = 16
        
        For ni = 1 To 8 Step 1
            
            Select Case ni
                Case 1, 3, 5, 7
                    .Row = ni:      .RowHeight(.Row) = 25
                Case 2, 4, 6, 8
                    .Row = ni:      .RowHeight(.Row) = 16
            End Select
        Next ni
        
    End With
        
    With cboKaeyol
        .Clear
        .AddItem "인문" & Space(30) & "01"
        .AddItem "자연" & Space(30) & "02"
        
        .ListIndex = 0
    End With
    
    With cboLsnType
        .Clear
        .AddItem "A type" & Space(30) & "A"
        .AddItem "B type" & Space(30) & "B"
        .AddItem "C type" & Space(30) & "C"
        
        .ListIndex = 0
    End With
    
    cmdFind.Tag = "LOAD"
    
    Call Find_LsnCD         '< 반 조회
    Call cmdFind_Click
    
    cmdFind.Tag = ""
    
End Sub

Public Sub init_Data(ByVal aKaeyol As String, ByVal aLsnType As String)
    
    Me.Tag = "LOAD"
    
    sprGwamok.MaxCols = 0
    
    With cboKaeyol
        Select Case aKaeyol
            Case "01"
                .ListIndex = 0
                
                optTamgu(0).Caption = "선택/삭제"
                optTamgu(1).Caption = "국사":           optTamgu(1).Tag = "01"
                optTamgu(2).Caption = "윤리":           optTamgu(2).Tag = "02"
                optTamgu(3).Caption = "경제":           optTamgu(3).Tag = "03"
                optTamgu(4).Caption = "한근":           optTamgu(4).Tag = "04"
                optTamgu(5).Caption = "세계사":         optTamgu(5).Tag = "05"
                optTamgu(6).Caption = "경지":           optTamgu(6).Tag = "06"
                optTamgu(7).Caption = "한지":           optTamgu(7).Tag = "07"
                optTamgu(8).Caption = "정치":           optTamgu(8).Tag = "08"
                optTamgu(9).Caption = "사문":           optTamgu(9).Tag = "09":             optTamgu(9).Visible = True
                optTamgu(10).Caption = "법사":          optTamgu(10).Tag = "10":            optTamgu(10).Visible = True
                optTamgu(11).Caption = "세지":          optTamgu(11).Tag = "11":            optTamgu(11).Visible = True
                
                optTamgu(0).value = True            '기본선택
                
            Case "02"
                .ListIndex = 1
                
                optTamgu(0).Caption = "선택/삭제"
                optTamgu(1).Caption = "물1":            optTamgu(1).Tag = "51"
                optTamgu(2).Caption = "화1":            optTamgu(2).Tag = "52"
                optTamgu(3).Caption = "생1":            optTamgu(3).Tag = "53"
                optTamgu(4).Caption = "지1":            optTamgu(4).Tag = "54"
                optTamgu(5).Caption = "물2":            optTamgu(5).Tag = "55"
                optTamgu(6).Caption = "화2":            optTamgu(6).Tag = "56"
                optTamgu(7).Caption = "생2":            optTamgu(7).Tag = "57"
                optTamgu(8).Caption = "지2":            optTamgu(8).Tag = "58"
                
                optTamgu(9).Caption = "":               optTamgu(9).Tag = "00":             optTamgu(9).Visible = False
                optTamgu(10).Caption = "":              optTamgu(10).Tag = "00":            optTamgu(10).Visible = False
                optTamgu(11).Caption = "":              optTamgu(11).Tag = "00":            optTamgu(11).Visible = False
    
                optTamgu(0).value = True            '기본선택
                
        End Select
    End With
    
    With cboLsnType
        Select Case aLsnType
            Case "A"
                .ListIndex = 0
            Case "B"
                .ListIndex = 1
            Case "C"
                .ListIndex = 2
        End Select
    End With

    
    cmdFind.Tag = "FIRST"
        Call cmdFind_Click
        
    cmdFind.Tag = ""

    Me.Tag = ""

End Sub

'## 반 내역 조회
Private Sub Find_LsnCD()
    
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    
    Dim nLength     As Long
    Dim sStr        As String
    Dim ni          As Integer
    Dim nRec        As Long
    
    Dim sTmp        As String
    Dim nColor      As Long
    Dim nRow        As Long
    
    sprGwamok.MaxCols = 0
    
    On Error GoTo ErrStmt
    
    sStr = ""
    sStr = sStr & "    SELECT A.LSNCD, LSNNM, LSNCDNM, LSN_CL "
    sStr = sStr & "      FROM (SELECT ACID, LSNCD, LSNNM, LSNCDNM, LSN_CL"
    sStr = sStr & "              FROM (SELECT *"
    sStr = sStr & "                      From SDLSN01TB"
    sStr = sStr & "                     WHERE ACID   = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "                       AND KAEYOL = '" & Trim(Right(cboKaeyol.Text, 30)) & "'"
    sStr = sStr & "                     ORDER BY LSNCDNM"
    sStr = sStr & "                    )"
    sStr = sStr & "            Union All"
    sStr = sStr & "            SELECT ACID, LSNCD, LSNNM, LSNCDNM, LSN_CL"
    sStr = sStr & "              FROM (SELECT *"
    sStr = sStr & "                      From SDLSN02TB"
    sStr = sStr & "                     WHERE ACID   = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "                       AND KAEYOL = '" & Trim(Right(cboKaeyol.Text, 30)) & "'"
    sStr = sStr & "                     ORDER BY LSNCDNM"
    sStr = sStr & "                    )"
    sStr = sStr & "            ) A, "
    sStr = sStr & "            SDLSN05TB B"
    sStr = sStr & "      WHERE A.ACID    = B.ACID"
    sStr = sStr & "        AND A.LSNCD   = B.LSNCD "
    sStr = sStr & "        AND A.ACID    = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "        AND B.LSNTYPE = '" & Trim(Right(cboLsnType.Text, 30)) & "'"
    
    
    Set DBCmd = New ADODB.Command
    Set DBRec = New ADODB.Recordset
    Set DBParam = New ADODB.Parameter
    
    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
    DBCmd.CommandText = sStr
    DBCmd.CommandType = adCmdText
    DBCmd.CommandTimeout = 30
    


'    '>> 분원
'        sTmp = Trim(basModule.SchCD)
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
       
        
    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic형태로 열게되면 record count를 할 수 없음.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop
    
    With DBRec
        If .RecordCount > 0 Then
            .MoveFirst
            
            For nRec = 1 To .RecordCount Step 1
                sprGwamok.MaxCols = sprGwamok.MaxCols + 1
                sprGwamok.Col = sprGwamok.MaxCols
                
                sprGwamok.Row = SpreadHeader
                    sTmp = "":      If IsNull(.Fields("LSNCD")) = False Then sTmp = Trim(.Fields("LSNCD")):         sprGwamok.Text = sTmp
                sprGwamok.Row = SpreadHeader + 1
                    sTmp = "":      If IsNull(.Fields("LSNCDNM")) = False Then sTmp = Trim(.Fields("LSNCDNM")):     sprGwamok.Text = sTmp
                sprGwamok.Row = SpreadHeader + 2
                    sTmp = "":      If IsNull(.Fields("LSNNM")) = False Then sTmp = Trim(.Fields("LSNNM")):         sprGwamok.Text = sTmp
                    
                For nRow = 1 To 8 Step 1
                    sprGwamok.Row = nRow:   Call basFunction.Set_SprType_Text(sprGwamok, "center", "center", 30, "")
                Next nRow
                
                .MoveNext       '<< 다음항목
                
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
    MsgBox "반 조회시 오류가 발생하였습니다.", vbCritical + vbOKOnly, "반 조회"
End Sub

Private Sub cmdFind_Click()
    cmdFind.Enabled = False
        
        Call Find_Gwamok_Detail
        
    cmdFind.Enabled = True
End Sub


Private Sub Find_Gwamok_Detail()
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    
    Dim nLength     As Long
    Dim sStr        As String
    Dim ni          As Integer
    Dim nRec        As Long
    
    Dim sTmp        As String
    Dim nTmp        As Long
    
    Dim nCol        As Long
    Dim sLsnCD      As String
    Dim nColor      As Long
    
    Call Find_LsnCD         '< 반 조회
    
    On Error GoTo ErrStmt
    
    sStr = ""
    sStr = sStr & "  SELECT LSNCD, ORD, GET_TCRNM(ACID, TCRCD) AS TCRNM, "

    sStr = sStr & "         CASE WHEN      TRIM(SUBJCD) = '01' THEN '국사'"
    sStr = sStr & "         ELSE CASE WHEN TRIM(SUBJCD) = '02' THEN '윤리'"
    sStr = sStr & "         ELSE CASE WHEN TRIM(SUBJCD) = '03' THEN '경제'"
    sStr = sStr & "         ELSE CASE WHEN TRIM(SUBJCD) = '04' THEN '한근'"
    sStr = sStr & "         ELSE CASE WHEN TRIM(SUBJCD) = '05' THEN '세계사'"
    sStr = sStr & "         ELSE CASE WHEN TRIM(SUBJCD) = '06' THEN '경지'"
    sStr = sStr & "         ELSE CASE WHEN TRIM(SUBJCD) = '07' THEN '한지'"
    sStr = sStr & "         ELSE CASE WHEN TRIM(SUBJCD) = '08' THEN '정치'"
    sStr = sStr & "         ELSE CASE WHEN TRIM(SUBJCD) = '09' THEN '사문'"
    sStr = sStr & "         ELSE CASE WHEN TRIM(SUBJCD) = '10' THEN '법사'"
    sStr = sStr & "         ELSE CASE WHEN TRIM(SUBJCD) = '11' THEN '세지'"
    
    sStr = sStr & "         ELSE CASE WHEN TRIM(SUBJCD) = '51' THEN '물1'"
    sStr = sStr & "         ELSE CASE WHEN TRIM(SUBJCD) = '52' THEN '화1'"
    sStr = sStr & "         ELSE CASE WHEN TRIM(SUBJCD) = '53' THEN '생1'"
    sStr = sStr & "         ELSE CASE WHEN TRIM(SUBJCD) = '54' THEN '지1'"
    sStr = sStr & "         ELSE CASE WHEN TRIM(SUBJCD) = '55' THEN '물2'"
    sStr = sStr & "         ELSE CASE WHEN TRIM(SUBJCD) = '56' THEN '화2'"
    sStr = sStr & "         ELSE CASE WHEN TRIM(SUBJCD) = '57' THEN '생2'"
    sStr = sStr & "         ELSE CASE WHEN TRIM(SUBJCD) = '58' THEN '지2'"
    
    sStr = sStr & "         END END END END END END END END END END END"
    sStr = sStr & "         END END END END END END END END AS SUBJCD,"
    
    sStr = sStr & "         CASE WHEN      TRIM(SUBJCD) = '01' THEN 1"
    sStr = sStr & "         ELSE CASE WHEN TRIM(SUBJCD) = '02' THEN 2"
    sStr = sStr & "         ELSE CASE WHEN TRIM(SUBJCD) = '03' THEN 3"
    sStr = sStr & "         ELSE CASE WHEN TRIM(SUBJCD) = '04' THEN 4"
    sStr = sStr & "         ELSE CASE WHEN TRIM(SUBJCD) = '05' THEN 5"
    sStr = sStr & "         ELSE CASE WHEN TRIM(SUBJCD) = '06' THEN 6"
    sStr = sStr & "         ELSE CASE WHEN TRIM(SUBJCD) = '07' THEN 7"
    sStr = sStr & "         ELSE CASE WHEN TRIM(SUBJCD) = '08' THEN 8"
    sStr = sStr & "         ELSE CASE WHEN TRIM(SUBJCD) = '09' THEN 9"
    sStr = sStr & "         ELSE CASE WHEN TRIM(SUBJCD) = '10' THEN 10"
    sStr = sStr & "         ELSE CASE WHEN TRIM(SUBJCD) = '11' THEN 11"
    
    sStr = sStr & "         ELSE CASE WHEN TRIM(SUBJCD) = '51' THEN 1"
    sStr = sStr & "         ELSE CASE WHEN TRIM(SUBJCD) = '52' THEN 2"
    sStr = sStr & "         ELSE CASE WHEN TRIM(SUBJCD) = '53' THEN 3"
    sStr = sStr & "         ELSE CASE WHEN TRIM(SUBJCD) = '54' THEN 4"
    sStr = sStr & "         ELSE CASE WHEN TRIM(SUBJCD) = '55' THEN 5"
    sStr = sStr & "         ELSE CASE WHEN TRIM(SUBJCD) = '56' THEN 6"
    sStr = sStr & "         ELSE CASE WHEN TRIM(SUBJCD) = '57' THEN 7"
    sStr = sStr & "         ELSE CASE WHEN TRIM(SUBJCD) = '58' THEN 8"
    sStr = sStr & "         ELSE 0"
    sStr = sStr & "         END END END END END END END END END END END"
    sStr = sStr & "         END END END END END END END END AS COLORS"
    
    sStr = sStr & "    FROM SDLSN06TB"
    sStr = sStr & "   WHERE ACID    = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "     AND KAEYOL  = '" & Trim(Right(cboKaeyol.Text, 30)) & "'"
    sStr = sStr & "     AND LSNTYPE = '" & Trim(Right(cboLsnType.Text, 30)) & "'"
    
    Set DBCmd = New ADODB.Command
    Set DBRec = New ADODB.Recordset
    Set DBParam = New ADODB.Parameter
    
    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
    DBCmd.CommandText = sStr
    DBCmd.CommandType = adCmdText
    DBCmd.CommandTimeout = 30
    


    
'    '>> 분원
'        sTmp = Trim(basModule.SchCD)
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
       
        
    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic형태로 열게되면 record count를 할 수 없음.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop
    
    With DBRec
        If .RecordCount > 0 Then
            .MoveFirst
            
            For nRec = 1 To .RecordCount Step 1
                
                
                sLsnCD = "":      If IsNull(.Fields("LSNCD")) = False Then sLsnCD = Trim(.Fields("LSNCD"))
                If sLsnCD <> "" Then
                        
                    sprGwamok.Row = SpreadHeader
                    For nCol = 1 To sprGwamok.MaxCols Step 1
                        sprGwamok.Col = nCol
                        
                        If StrComp(Trim(sprGwamok.Text), sLsnCD, vbTextCompare) = 0 Then            '< LSNCD 비교
                            If IsNumeric(.Fields("ORD")) = True Then                                '< ORD : 행
                                
                                Select Case CLng(.Fields("ORD"))
                                    Case 1
                                        sprGwamok.Row = 1
                                            sTmp = "":      If IsNull(.Fields("SUBJCD")) = False Then sTmp = Trim(.Fields("SUBJCD"))
                                            If sTmp = "" Then
                                                Call basFunction.Set_SprType_Text(sprGwamok, "CENTER", "CENTER", 10, "")
                                            Else
                                                Call basFunction.Set_SprType_Text(sprGwamok, "CENTER", "CENTER", 10, sTmp)
                                            End If
                                            
                                            sprGwamok.Row2 = sprGwamok.Row
                                            sprGwamok.Col2 = sprGwamok.Col
                                            sprGwamok.BlockMode = True
                                                nTmp = &HFFFFFF:        If IsNumeric(.Fields("COLORS")) = True Then nTmp = CLng(.Fields("COLORS"))
                                                If nTmp = 0 Or nTmp = &HFFFFFF Then
                                                    nTmp = 0
                                                    sprGwamok.BackColor = optTamgu(nTmp).BackColor
                                                Else
                                                    sprGwamok.BackColor = optTamgu(nTmp).BackColor
                                                End If
                                                sprGwamok.BackColorStyle = BackColorStyleUnderGrid
                                            sprGwamok.BlockMode = False
                                                                                        
                                        sprGwamok.Row = 2
                                            sTmp = "":      If IsNull(.Fields("TCRNM")) = False Then sTmp = Trim(.Fields("TCRNM"))
                                            If sTmp = "" Then
                                                Call basFunction.Set_SprType_Text(sprGwamok, "CENTER", "CENTER", 10, "")
                                            Else
                                                Call basFunction.Set_SprType_Text(sprGwamok, "CENTER", "CENTER", 10, sTmp)
                                            End If
                                    Case 2
                                        sprGwamok.Row = 3
                                            sTmp = "":      If IsNull(.Fields("SUBJCD")) = False Then sTmp = Trim(.Fields("SUBJCD"))
                                            If sTmp = "" Then
                                                Call basFunction.Set_SprType_Text(sprGwamok, "CENTER", "CENTER", 10, "")
                                            Else
                                                Call basFunction.Set_SprType_Text(sprGwamok, "CENTER", "CENTER", 10, sTmp)
                                            End If
                                            
                                            sprGwamok.Row2 = sprGwamok.Row
                                            sprGwamok.Col2 = sprGwamok.Col
                                            sprGwamok.BlockMode = True
                                                nTmp = &HFFFFFF:        If IsNumeric(.Fields("COLORS")) = True Then nTmp = CLng(.Fields("COLORS"))
                                                If nTmp = 0 Or nTmp = &HFFFFFF Then
                                                    nTmp = 0
                                                    sprGwamok.BackColor = optTamgu(nTmp).BackColor
                                                Else
                                                    sprGwamok.BackColor = optTamgu(nTmp).BackColor
                                                End If
                                                sprGwamok.BackColorStyle = BackColorStyleUnderGrid
                                            sprGwamok.BlockMode = False
                                            
                                        sprGwamok.Row = 4
                                            sTmp = "":      If IsNull(.Fields("TCRNM")) = False Then sTmp = Trim(.Fields("TCRNM"))
                                            If sTmp = "" Then
                                                Call basFunction.Set_SprType_Text(sprGwamok, "CENTER", "CENTER", 10, "")
                                            Else
                                                Call basFunction.Set_SprType_Text(sprGwamok, "CENTER", "CENTER", 10, sTmp)
                                            End If
                                    Case 3
                                        sprGwamok.Row = 5
                                            sTmp = "":      If IsNull(.Fields("SUBJCD")) = False Then sTmp = Trim(.Fields("SUBJCD"))
                                            If sTmp = "" Then
                                                Call basFunction.Set_SprType_Text(sprGwamok, "CENTER", "CENTER", 10, "")
                                            Else
                                                Call basFunction.Set_SprType_Text(sprGwamok, "CENTER", "CENTER", 10, sTmp)
                                            End If
                                            
                                            sprGwamok.Row2 = sprGwamok.Row
                                            sprGwamok.Col2 = sprGwamok.Col
                                            sprGwamok.BlockMode = True
                                                nTmp = &HFFFFFF:        If IsNumeric(.Fields("COLORS")) = True Then nTmp = CLng(.Fields("COLORS"))
                                                If nTmp = 0 Or nTmp = &HFFFFFF Then
                                                    nTmp = 0
                                                    sprGwamok.BackColor = optTamgu(nTmp).BackColor
                                                Else
                                                    sprGwamok.BackColor = optTamgu(nTmp).BackColor
                                                End If
                                                sprGwamok.BackColorStyle = BackColorStyleUnderGrid
                                            sprGwamok.BlockMode = False
                                            
                                        sprGwamok.Row = 6
                                            sTmp = "":      If IsNull(.Fields("TCRNM")) = False Then sTmp = Trim(.Fields("TCRNM"))
                                            If sTmp = "" Then
                                                Call basFunction.Set_SprType_Text(sprGwamok, "CENTER", "CENTER", 10, "")
                                            Else
                                                Call basFunction.Set_SprType_Text(sprGwamok, "CENTER", "CENTER", 10, sTmp)
                                            End If
                                    Case 4
                                        sprGwamok.Row = 7
                                            sTmp = "":      If IsNull(.Fields("SUBJCD")) = False Then sTmp = Trim(.Fields("SUBJCD"))
                                            If sTmp = "" Then
                                                Call basFunction.Set_SprType_Text(sprGwamok, "CENTER", "CENTER", 10, "")
                                            Else
                                                Call basFunction.Set_SprType_Text(sprGwamok, "CENTER", "CENTER", 10, sTmp)
                                            End If
                                            
                                            sprGwamok.Row2 = sprGwamok.Row
                                            sprGwamok.Col2 = sprGwamok.Col
                                            sprGwamok.BlockMode = True
                                                nTmp = &HFFFFFF:        If IsNumeric(.Fields("COLORS")) = True Then nTmp = CLng(.Fields("COLORS"))
                                                If nTmp = 0 Or nTmp = &HFFFFFF Then
                                                    nTmp = 0
                                                    sprGwamok.BackColor = optTamgu(nTmp).BackColor
                                                Else
                                                    sprGwamok.BackColor = optTamgu(nTmp).BackColor
                                                End If
                                                sprGwamok.BackColorStyle = BackColorStyleUnderGrid
                                            sprGwamok.BlockMode = False
                                            
                                        sprGwamok.Row = 8
                                            sTmp = "":      If IsNull(.Fields("TCRNM")) = False Then sTmp = Trim(.Fields("TCRNM"))
                                            If sTmp = "" Then
                                                Call basFunction.Set_SprType_Text(sprGwamok, "CENTER", "CENTER", 10, "")
                                            Else
                                                Call basFunction.Set_SprType_Text(sprGwamok, "CENTER", "CENTER", 10, sTmp)
                                            End If
                                End Select
                                
                                
                                
                            End If
                        End If
                    Next nCol
                End If
                
                .MoveNext       '<< 다음항목
                
            Next nRec
        End If
    End With
    
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    
    If cmdFind.Tag = "" Then
        MsgBox "조회하였습니다.", vbInformation + vbOKOnly, "과목 등록내역 조회"
    End If
    
    
    Exit Sub
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    On Error GoTo 0
    MsgBox "과목 등록내역 조회시 오류가 발생하였습니다.", vbCritical + vbOKOnly, "과목 등록내역 조회"
End Sub


'## 과목 선택함.
Private Sub cboKaeyol_Click()
    If Me.Tag = "LOAD" Then Exit Sub
    
    Select Case Trim(Right(cboKaeyol.Text, 30))
        Case "01", "03"         '<< 인문
            
            optTamgu(0).Caption = "선택/삭제":      optTamgu(0).Tag = "00"
            optTamgu(1).Caption = "국사":           optTamgu(1).Tag = "01"
            optTamgu(2).Caption = "윤리":           optTamgu(2).Tag = "02"
            optTamgu(3).Caption = "경제":           optTamgu(3).Tag = "03"
            optTamgu(4).Caption = "한근":           optTamgu(4).Tag = "04"
            optTamgu(5).Caption = "세계사":         optTamgu(5).Tag = "05"
            optTamgu(6).Caption = "경지":           optTamgu(6).Tag = "06"
            optTamgu(7).Caption = "한지":           optTamgu(7).Tag = "07"
            optTamgu(8).Caption = "정치":           optTamgu(8).Tag = "08"
            optTamgu(9).Caption = "사문":           optTamgu(9).Tag = "09":             optTamgu(9).Visible = True:     optTamgu(9).BackColor = &HC000&
            optTamgu(10).Caption = "법사":          optTamgu(10).Tag = "10":            optTamgu(10).Visible = True:    optTamgu(10).BackColor = &HFF&
            optTamgu(11).Caption = "세지":          optTamgu(11).Tag = "11":            optTamgu(11).Visible = True:    optTamgu(11).BackColor = &HC0C0C0
            
            optTamgu(0).value = True            '기본선택
            
        Case "02"       '<< 자연
            
            optTamgu(0).Caption = "선택/삭제":      optTamgu(0).Tag = "00"
            optTamgu(1).Caption = "물1":            optTamgu(1).Tag = "51"
            optTamgu(2).Caption = "화1":            optTamgu(2).Tag = "52"
            optTamgu(3).Caption = "생1":            optTamgu(3).Tag = "53"
            optTamgu(4).Caption = "지1":            optTamgu(4).Tag = "54"
            optTamgu(5).Caption = "물2":            optTamgu(5).Tag = "55"
            optTamgu(6).Caption = "화2":            optTamgu(6).Tag = "56"
            optTamgu(7).Caption = "생2":            optTamgu(7).Tag = "57"
            optTamgu(8).Caption = "지2":            optTamgu(8).Tag = "58"
            
            optTamgu(9).Caption = "":               optTamgu(9).Tag = "00":             optTamgu(9).Visible = False:    optTamgu(9).BackColor = basModule.WhiteColor
            optTamgu(10).Caption = "":              optTamgu(10).Tag = "00":            optTamgu(10).Visible = False:   optTamgu(10).BackColor = basModule.WhiteColor
            optTamgu(11).Caption = "":              optTamgu(11).Tag = "00":            optTamgu(11).Visible = False:   optTamgu(11).BackColor = basModule.WhiteColor

            optTamgu(0).value = True            '기본선택
            
    End Select
    
    Call Find_LsnCD         '< 반 조회
    
End Sub

'// 과목선택
Private Sub sprGwamok_DblClick(ByVal Col As Long, ByVal Row As Long)

    Dim ni          As Integer
    Dim ninDex      As Integer
    Dim sTmp        As String

    If Row < 1 Then Exit Sub
    If Col < 1 Then Exit Sub

    With sprGwamok      '<< COLUMN값은 고정됨.
        If .MaxCols = 0 Then Exit Sub

        For ni = 0 To optTamgu.UBound Step 1
            If optTamgu(ni).value = True Then
                ninDex = ni
                Exit For
            End If
        Next ni

        Select Case Row
            Case 1, 3, 5, 7
                If optTamgu(ninDex).value = True Then
                    .Row = Row:     .Row2 = Row
                    .Col = Col:     .Col2 = Col
                    .BlockMode = True
                        .BackColor = optTamgu(ninDex).BackColor
                        .BackColorStyle = BackColorStyleUnderGrid
                    .BlockMode = False

                    Select Case optTamgu(ninDex).Tag
                        Case "00"
                            .Text = ""
                        Case Else
                            sTmp = optTamgu(ninDex).Caption
                            Call basFunction.Set_SprType_Text(sprGwamok, "center", "center", basFunction.LenKor(sTmp), sTmp)
                    End Select
                End If
        End Select
    End With
End Sub

Private Sub sprGwamok_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Button
        Case vbRightButton
            With sprGwamok
                .Row = .ActiveRow
                .Col = .ActiveCol

                    .Text = ""

                .Row2 = .Row
                .Col2 = .Col
                .BlockMode = True
                    .BackColor = basModule.WhiteColor
                    .BackColorStyle = BackColorStyleUnderGrid
                .BlockMode = False
            End With
    End Select

End Sub

Private Sub sprGwamok_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    
    Dim ninDex  As Integer

    With sprGwamok
        
        .Row = Row
        .Col = Col

        Select Case Trim(.Text)
            Case "국사":     ninDex = 1
            Case "윤리":     ninDex = 2
            Case "경제":     ninDex = 3
            Case "한근":     ninDex = 4
            Case "세계사", "세사":   ninDex = 5
            Case "경지":     ninDex = 6
            Case "한지":     ninDex = 7
            Case "정치":     ninDex = 8
            Case "사문":     ninDex = 9
            Case "법사":     ninDex = 10
            Case "세지":     ninDex = 11

            Case "물1":     ninDex = 1
            Case "화1":     ninDex = 2
            Case "생1":     ninDex = 3
            Case "지1":     ninDex = 4
            Case "물2":     ninDex = 5
            Case "화2":     ninDex = 6
            Case "생2":     ninDex = 7
            Case "지2":     ninDex = 8

            Case Else:      ninDex = 0
        End Select

        If ninDex = 0 Then
            .Row2 = .Row
            .Col2 = .Col
            .BlockMode = True
                .BackColor = basModule.WhiteColor
                .BackColorStyle = BackColorStyleUnderGrid
            .BlockMode = False
        Else
            .Row2 = .Row
            .Col2 = .Col
            .BlockMode = True
                .BackColor = optTamgu(ninDex).BackColor
                .BackColorStyle = BackColorStyleUnderGrid
            .BlockMode = False
        End If
            
        If NewCol < 1 Then Exit Sub
        If NewRow < 1 Then Exit Sub
        
    End With
End Sub

Private Sub sprGwamok_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim ninDex  As Integer

    With sprGwamok
        If .ActiveCol < 1 Then Exit Sub
        If .ActiveRow < 1 Then Exit Sub

        .Row = .ActiveRow
        .Col = .ActiveCol

        If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then
            .Text = ""
            .Row2 = .Row
            .Col2 = .Col

            .BlockMode = True
                .BackColor = basModule.WhiteColor
                .BackColorStyle = BackColorStyleUnderGrid
            .BlockMode = False
            Exit Sub
        End If
        
    End With
End Sub



















'## 과목내역 등록
Private Sub cmdSave_Click()
    Dim sTmp        As String
    
    cmdSave.Enabled = False
    
        With sprGwamok
            If .MaxCols = 0 Then
                MsgBox "등록할 내역이 없습니다.", vbExclamation + vbOKOnly, "과목등록"
                cmdSave.Enabled = True
                Exit Sub
            End If
            
            sTmp = ""
            sTmp = "【 " & Trim(Left(cboKaeyol.Text, 30)) & " 】계열 "
            sTmp = sTmp & "【 " & Trim(Left(cboLsnType.Text, 30))
            sTmp = sTmp & " 】타입으로 현 선택과목 내역을 등록하시겠습니까?"
            If MsgBox(sTmp, vbQuestion + vbYesNo, "선택과목 등록") = vbNo Then
                cmdSave.Enabled = True
                Exit Sub
            End If
            
            Call Save_inPutData
        
        End With
        
    
    cmdSave.Enabled = True
    
End Sub

Private Sub Save_inPutData()
    
    Dim DBCmd       As ADODB.Command        '<< 학생 반 내역 등록하기
    Dim DBParam     As ADODB.Parameter
    
    Dim sStr        As String
    Dim nTot        As Long
    Dim nExeTot     As Long
    Dim nExe        As Long
    Dim nLength     As Long
    
    Dim nRow        As Long
    Dim nCol        As Integer
    Dim ni          As Integer
    
    Dim sTmp        As String
    Dim nTmp        As Long
    Dim sGwamok     As String
    
    Dim sTcrCD      As String
    Dim nC          As Long
    
'>> 등록방법 : 기존의 등록된 type 에 해당하는 내역을 모두 삭제 후 처리함.
    
    On Error GoTo ErrStmt
    
    basDataBase.DBConn.BeginTrans

    Set DBCmd = New ADODB.Command
    Set DBParam = New ADODB.Parameter

    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection


    '<< TYPE 에 해당하는 내역을 모두 삭제 >>
    sStr = ""
    sStr = sStr & " DELETE "
    sStr = sStr & "   FROM SDLSN06TB "
    sStr = sStr & "  WHERE ACID    = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "    AND KAEYOL  = '" & Trim(Right(cboKaeyol.Text, 30)) & "'"
    sStr = sStr & "    AND LSNTYPE = '" & Trim(Right(cboLsnType.Text, 30)) & "'"
    
'    '>> ACID
'        sTmp = Trim(basModule.SchCD)
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("SEL_CLASS", adVarChar, adParamInput, nLength, Trim(sTmp)): DBCmd.Parameters.Append DBParam

    DBCmd.CommandText = sStr
    DBCmd.CommandType = adCmdText
    DBCmd.CommandTimeout = 30
    
    nExe = 0
    DBCmd.Execute nExe, , -1
    
    Do While basDataBase.DBConn.State And adStateExecuting
        DoEvents
    Loop
    
    
    '<< 행의 값을 모두 저장 >>
    With sprGwamok
        nTot = 0
        nExeTot = 0
        nExe = 0
        
        For nCol = 1 To .MaxCols Step 1
            nC = 1          '< ord 값 처리
            
            For nRow = 1 To .MaxRows Step 2
                nTot = nTot + 1
                
                sStr = ""
                sStr = sStr & " INSERT INTO SDLSN06TB ( ACID       , KAEYOL     , LSNTYPE    , LSNCD      , ORD        , SUBJCD     , TCRCD     ) "
                sStr = sStr & " VALUES ( "
                sStr = sStr & "       '" & Trim(basModule.SchCD) & "', "                '< ACID
                sStr = sStr & "       '" & Trim(Right(cboKaeyol.Text, 30)) & "', "      '< KAEYOL
                sStr = sStr & "       '" & Trim(Right(cboLsnType.Text, 30)) & "', "     '< LSNTYPE
                
                .Row = SpreadHeader
                .Col = nCol
                    sTmp = Trim(.Text)
                        sStr = sStr & " '" & sTmp & "', "                               '< LSNCD
                
                
                .Row = nRow
                        sStr = sStr & "  " & Trim(CStr(1 + (nRow - nC))) & ", "         '< ORD
                                                                   nC = nC + 1              '< ORD 값 처리 count  1, 3, 5, 7 -> 1, 2, 3, 4 로 바꿈.
                        
                        .Col = nCol
                            Select Case Trim(.Text)                     '< 과목체크
                                Case "국사":     sGwamok = "01"
                                Case "윤리":     sGwamok = "02"
                                Case "경제":     sGwamok = "03"
                                Case "한근":     sGwamok = "04"
                                Case "세계사", "세사":     sGwamok = "05"
                                Case "경지":     sGwamok = "06"
                                Case "한지":     sGwamok = "07"
                                Case "정치":     sGwamok = "08"
                                Case "사문":     sGwamok = "09"
                                Case "법사":     sGwamok = "10"
                                Case "세지":     sGwamok = "11"
                                
                                Case "물1":     sGwamok = "51"
                                Case "화1":     sGwamok = "52"
                                Case "생1":     sGwamok = "53"
                                Case "지1":     sGwamok = "54"
                                Case "물2":     sGwamok = "55"
                                Case "화2":     sGwamok = "56"
                                Case "생2":     sGwamok = "57"
                                Case "지2":     sGwamok = "58"
                                Case "":     sGwamok = ""
                            End Select
                        sStr = sStr & " '" & sGwamok & "', "                            '< SUBJCD
                        
                .Row = nRow + 1
                    If Trim(.Text) = "" Then
                        sTmp = ""
                    Else
                        sTmp = Get_TcrCD(Trim(.Text))
                    End If
                        sStr = sStr & " '" & sTmp & "' "                                '< TCRCD
                        
                sStr = sStr & " )"
                
                
                DBCmd.CommandText = sStr
                DBCmd.CommandType = adCmdText
                DBCmd.CommandTimeout = 30
        
                nExe = 0
                DBCmd.Execute nExe, , -1
        
                Do While basDataBase.DBConn.State And adStateExecuting
                    DoEvents
                Loop
        
                If nExe = 1 Then
                    nExeTot = nExeTot + 1
                End If
            
            Next nRow
        Next nCol
        
    End With
    
    '>> 처리수가 동일해야 함.
    If nTot = nExeTot Then
        basDataBase.DBConn.CommitTrans
        MsgBox "과목 등록하였습니다.", vbInformation + vbOKOnly, "과목등록"
    Else
        basDataBase.DBConn.RollbackTrans
        MsgBox "등록 중 에러가 발생하였습니다.", vbCritical + vbOKOnly, "과목등록"
    End If
    
    ' NO ERROR
    Set DBCmd = Nothing
    Set DBParam = Nothing
    
    Exit Sub
    
ErrStmt:
    basDataBase.DBConn.RollbackTrans
    Set DBCmd = Nothing
    Set DBParam = Nothing
    
    MsgBox "과목 등록 중 에러가 발생하였습니다." & vbCrLf & _
           Trim(CStr(Err.Number)) & " " & Err.Description, vbCritical + vbOKOnly, "과목등록"
    
    On Error GoTo 0
End Sub

'## 강사코드 가져오기
Private Function Get_TcrCD(ByVal aTcrNM As String) As String
    
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    
    Dim nLength     As Long
    Dim sStr        As String
    Dim ni          As Integer
    Dim nRec        As Long
    
    Dim sTmp        As String
    Dim nColor      As Long
    Dim sRet        As String
    
    On Error GoTo ErrStmt
    
    sStr = ""
    sStr = sStr & " SELECT TCRCD "
    sStr = sStr & "   FROM SDTCR01TB "
    sStr = sStr & "  WHERE ACID  = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "    AND TCRNM = '" & Trim(aTcrNM) & "'"
    
    Set DBCmd = New ADODB.Command
    Set DBRec = New ADODB.Recordset
    Set DBParam = New ADODB.Parameter
    
    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
    DBCmd.CommandText = sStr
    DBCmd.CommandType = adCmdText
    DBCmd.CommandTimeout = 30
    


    
'    '>> 분원
'        sTmp = Trim(basModule.SchCD)
'        nLength = LenB(StrConv(sTmp, vbFromUnicode)):   If nLength < 1 Then nLength = 1
'            Set DBParam = DBCmd.CreateParameter("ACID", adChar, adParamInput, nLength, Trim(sTmp)):   DBCmd.Parameters.Append DBParam
       
        
    DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic형태로 열게되면 record count를 할 수 없음.
    Do While DBRec.State And adStateExecuting
        DoEvents
    Loop
    
    sRet = ""
    With DBRec
        If .RecordCount > 0 Then
            .MoveFirst
            
            sRet = "":      If IsNull(.Fields("TCRCD")) = False Then sRet = Trim(.Fields("TCRCD"))
        End If
    End With
    
ErrStmt:
    Set DBParam = Nothing
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    On Error GoTo 0
    
    Get_TcrCD = sRet
    
End Function



'## 선택항목만 받기
Private Sub imgExcel_Click()
    
    Dim nRow        As Long
    Dim nCol        As Long
    Dim nRec        As Long
    
    Dim sTmp        As String
    Dim sComp       As String
    
    Dim sFileName   As String
    Dim sFilePath   As String
    Dim sLogFile    As String
    
    Dim nWeekSrt    As Long
    Dim nColor      As Long
    
    Dim nRet        As Long
    Dim nRow2       As Long
    
    
    Dim sTcrTmp     As String
    Dim sTcrComp    As String
    Dim nChkRow     As Long
    
    If sprGwamok.MaxCols = 0 Then Exit Sub
    
    On Error GoTo ErrDlg
    
    If Dir(App.Path & "\TMR_EXCEL", vbDirectory) = "" Then MkDir App.Path & "\TMR_EXCEL"

    'TEXT파일을 생성 처리합니다.
    With dlgExcel
        .CancelError = True
        .fileName = ""
        .InitDir = App.Path & "\TMR_EXCEL"
        .Filter = "DAT FILE(*.XLS)|*.XLS"
        .DefaultExt = "*.XLS"
        .ShowSave

        '파일명을 처리합니다.
        If (.fileName) = "" Then Exit Sub
        
        sFileName = Mid$(dlgExcel.FileTitle, 1, InStr(1, dlgExcel.FileTitle, ".", vbTextCompare) - 1)
        sFilePath = Mid$(dlgExcel.fileName, 1, Len(dlgExcel.fileName) - InStrB(1, dlgExcel.fileName, "\", vbTextCompare) - 1)
        sLogFile = sFilePath & sFileName & ".LOG"
        
    End With

    On Error GoTo 0
    On Error GoTo ErrExcel
    
    sprExcel.ColHeadersShow = True
    sprExcel.RowHeadersShow = True
    
    sprExcel.MaxRows = 0
    sprExcel.MaxCols = 0
    
    For nRow = 1 To sprGwamok.ColHeaderRows Step 1
        sprGwamok.Row = SpreadHeader + nRow - 1
            '< 데이타 복사 >
            sprExcel.MaxRows = sprExcel.MaxRows + 1
            sprExcel.Row = sprExcel.MaxRows                                         '< header row
        
            sprExcel.MaxCols = sprGwamok.RowHeaderCols + sprGwamok.MaxCols        '< 전체 cols
        
            '< Row Header 생성 >
            For nCol = 1 To sprGwamok.RowHeaderCols Step 1
                sprGwamok.Col = SpreadHeader + nCol - 1
                    sTmp = sprGwamok.Text
                    
                    sprExcel.Col = nCol                                                 '< 데이터 넣음
                    Call basFunction.Set_SprType_Text(sprExcel, "center", "center", basFunction.LenKor(sTmp), sTmp)
                    
                    With sprExcel
                        .Row2 = .Row
                        .Col2 = .Col
                        .BlockMode = True
                            .BackColor = basModule.ShadowColor1
                            .BackColorStyle = BackColorStyleUnderGrid
                        .BlockMode = False
                    End With
            Next nCol
            
            '< Data >
            For nCol = 1 To sprGwamok.MaxCols Step 1
                sprGwamok.Col = nCol
                    sTmp = Trim(sprGwamok.Text)
                
                    sprExcel.Col = sprGwamok.RowHeaderCols + nCol
                    Call basFunction.Set_SprType_Text(sprExcel, "center", "center", basFunction.LenKor(sTmp), sTmp)
            Next nCol
            
            sprExcel.SetCellBorder 1, sprExcel.Row, sprExcel.MaxCols, sprExcel.Row, 8, basModule.SectionColor1, CellBorderStyleSolid
            
            With sprExcel
                .Row2 = .Row
                .Col = 1:       .Col2 = .MaxCols
                .BlockMode = True
                    .BackColor = basModule.ShadowColor1
                    .BackColorStyle = BackColorStyleUnderGrid
                .BlockMode = False
            End With
    Next nRow
    
    '< Data 부분 >
    For nRow = 1 To sprGwamok.MaxRows Step 1
        sprGwamok.Row = nRow
        sprGwamok.Col = SpreadHeader:      sTcrComp = Trim(sprGwamok.Text)

        '< 데이타 복사 >
        sprExcel.MaxRows = sprExcel.MaxRows + 1
        sprExcel.Row = sprExcel.MaxRows                                         '< header row

        '< Row Header 생성 >
        For nCol = 1 To sprGwamok.RowHeaderCols Step 1
            sprGwamok.Col = SpreadHeader + nCol - 1
                sTmp = sprGwamok.Text

                sprExcel.Col = nCol                                                 '< 데이터 넣음
                Call basFunction.Set_SprType_Text(sprExcel, "center", "left", basFunction.LenKor(sTmp), sTmp)
                sprExcel.ColWidth(sprExcel.Col) = 5
                sprExcel.RowHeight(sprExcel.Row) = sprGwamok.RowHeight(sprGwamok.Row)
        Next nCol

        '< Data >
        For nCol = 1 To sprGwamok.MaxCols Step 1
            sprGwamok.Col = nCol:               nColor = sprGwamok.BackColor
                sTmp = Trim(sprGwamok.Text)

                sprExcel.Col = sprGwamok.RowHeaderCols + nCol
                    Call basFunction.Set_SprType_Text(sprExcel, "center", "left", basFunction.LenKor(sTmp), sTmp)
                sprExcel.ColWidth(sprExcel.Col) = 7
                sprExcel.RowHeight(sprExcel.Row) = sprGwamok.RowHeight(sprGwamok.Row)
                
                sprExcel.Row2 = sprExcel.Row
                sprExcel.Col2 = sprExcel.Col
                sprExcel.BlockMode = True
                    sprExcel.BackColor = nColor
                    sprExcel.BackColorStyle = BackColorStyleUnderGrid
                sprExcel.BlockMode = False
        Next nCol

        If (sprExcel.Row Mod 2) = 0 Then
            sprExcel.SetCellBorder 1, sprExcel.Row, sprExcel.MaxCols, sprExcel.Row, 8, basModule.SectionColor2, CellBorderStyleSolid
        Else
            sprExcel.SetCellBorder 1, sprExcel.Row, sprExcel.MaxCols, sprExcel.Row, 8, &H0, CellBorderStyleSolid
        End If

        With sprExcel
            .Row = 1:       .Row2 = .MaxRows
            .Col = 1:       .Col2 = sprGwamok.RowHeaderCols
            .BlockMode = True
                .BackColor = basModule.ShadowColor1
                .BackColorStyle = BackColorStyleUnderGrid
            .BlockMode = False
        End With

    Next nRow
    
    '< 데이터 정렬 및 맞춤 >
    With sprExcel
        If .MaxCols > 1 Then
            .Row = 1:       .RowHidden = True
            
            For nCol = 1 To .MaxCols Step 1
            .SetCellBorder nCol, 1, nCol, .MaxRows, 2, &H80000008, CellBorderStyleSolid
            Next nCol
        End If
    End With
    
    nRet = sprExcel.ExportToExcel(dlgExcel.fileName, "Time_Schedule", sLogFile)
    
    MsgBox "엑셀작성하였습니다." & vbCrLf & _
           "확인하십시요.", vbInformation + vbOKOnly, "시간표 엑셀자료 만들기"

    Exit Sub
ErrExcel:
    On Error GoTo 0
    
    MsgBox "엑셀자료 생성시 에러가 발생하였습니다.", vbCritical + vbOKOnly, "시간표 엑셀자료 만들기"
    Exit Sub
ErrDlg:
    On Error GoTo 0
    
    MsgBox "엑셀자료 생성을 취소하였습니다.", vbCritical + vbOKOnly, "엑셀자료 생성"
End Sub



