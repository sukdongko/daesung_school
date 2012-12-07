VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form TMR060 
   Caption         =   "시간표 출력 >> 강사 출석부"
   ClientHeight    =   9780
   ClientLeft      =   2160
   ClientTop       =   2535
   ClientWidth     =   15810
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9780
   ScaleWidth      =   15810
   Begin FPSpread.vaSpread sprNot 
      Height          =   8835
      Left            =   13290
      TabIndex        =   4
      Top             =   660
      Width           =   2235
      _Version        =   393216
      _ExtentX        =   3942
      _ExtentY        =   15584
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
      MaxCols         =   4
      SpreadDesigner  =   "TMR060.frx":0000
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0082C8E8&
      BorderStyle     =   0  '없음
      Caption         =   "Frame1"
      Height          =   585
      Left            =   60
      TabIndex        =   5
      Top             =   30
      Width           =   15435
      Begin VB.Frame Frame2 
         BackColor       =   &H00D2EAF5&
         BorderStyle     =   0  '없음
         Caption         =   "Frame2"
         Height          =   525
         Left            =   30
         TabIndex        =   6
         Top             =   30
         Width           =   15375
         Begin VB.ComboBox cboWeek 
            Height          =   300
            Left            =   2400
            Style           =   2  '드롭다운 목록
            TabIndex        =   1
            Top             =   105
            Width           =   1065
         End
         Begin VB.CommandButton cmdFindTmr 
            Caption         =   "조 회 (&F)"
            Height          =   375
            Left            =   3990
            TabIndex        =   2
            Top             =   60
            Width           =   1515
         End
         Begin EditLib.fpMask fpYM 
            Height          =   285
            Left            =   540
            TabIndex        =   0
            Top             =   120
            Width           =   1005
            _Version        =   196608
            _ExtentX        =   1773
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
            BackColor       =   16777215
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
         Begin VB.Label Label1 
            BackStyle       =   0  '투명
            Caption         =   "수업없는 강사"
            Height          =   210
            Left            =   13740
            TabIndex        =   8
            Top             =   300
            Width           =   1305
         End
         Begin VB.Label Label4 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "요일"
            Height          =   210
            Left            =   1380
            TabIndex        =   7
            Top             =   150
            Width           =   975
         End
      End
   End
   Begin FPSpread.vaSpread sprTmr 
      Height          =   8865
      Left            =   30
      TabIndex        =   3
      Top             =   660
      Width           =   13245
      _Version        =   393216
      _ExtentX        =   23363
      _ExtentY        =   15637
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
      SpreadDesigner  =   "TMR060.frx":18F9
   End
End
Attribute VB_Name = "TMR060"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################################################
'   시 스 템  명 : 대성학원 입학사정, 반배정 & 시간표 프로그램
'   서브시스템명 :
'   모   듈   명 : TMR060
'   모 듈  목 적 : 강사 출석부
'
'   작   성   일 : 2008/02/20
'   작   성   자 : 유하균
' --------------------------------------------------------------------------------------------------------------
' --------------------------------------------------------------------------------------------------------------
'                 수     정     내     용
' --------------------------------------------------------------------------------------------------------------
'   1. 수정일 :
'   2. 내  용 :
'################################################################################################################

Option Explicit


Private Type tAttend
    TCRCD       As String
    TCRNM       As String
    
    SUBJCD      As String
    SUBJNM      As String
    
    LSNCD       As String
    
    WEEKS       As String
    LESSON      As String
End Type
Private uAttend()       As tAttend


Private Sub Form_Load()
    Me.Move 0, 0, 15700, 9980
    
    fpYM.Text = Format(Now, "yyyymm")
    
    With sprTmr
        .ShadowColor = basModule.ShadowColor2
        .ShadowDark = basModule.ShadowDark2
        .ShadowText = basModule.ShadowText2
        .GridColor = basModule.GridColor2
        .GrayAreaBackColor = basModule.GrayAreaBackColor2
        
        .MaxRows = 0
        .MaxCols = 0
    End With
    
    With sprNot
        .ShadowColor = basModule.ShadowColor1
        .ShadowDark = basModule.ShadowDark1
        .ShadowText = basModule.ShadowText1
        .GridColor = basModule.GridColor1
        .GrayAreaBackColor = basModule.GrayAreaBackColor1
        
        .MaxRows = 0
        
    End With
    
    
    With cboWeek
        .Clear
        
        .AddItem "전체" & Space(30) & "ALL"
        .AddItem "월" & Space(30) & "1"
        .AddItem "화" & Space(30) & "2"
        .AddItem "수" & Space(30) & "3"
        .AddItem "목" & Space(30) & "4"
        .AddItem "금" & Space(30) & "5"
        .AddItem "토" & Space(30) & "6"
        .AddItem "일" & Space(30) & "7"
        
        .ListIndex = 1
    End With
    
    
End Sub





'#######################################################################################################################################################################
' 출석부 조회
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
    
    sprNot.MaxRows = 0
    
    
    ReDim uAttend(0) As tAttend     '< 초기화
    
    sprTmr.Visible = False
    Call Display_SprTmr_Row_SpreadHeader                    '<< COL 로 진행하는 ROW의 헤더 작성
    
    sprTmr.Visible = True
    
    If sprTmr.MaxCols > 1 Then
        
        '>> 강의시간 조회
        Call Find_TeachingTime
        
        If UBound(uAttend) > 0 Then
            '> 내역 보여주기
                        
            Call Show_First_TeachingTime
            
        End If
        
    End If
End Sub

'>> 내역 보여주기
Private Sub Show_First_TeachingTime()

    Dim sTcrCD      As String
    Dim sTmpTcrCD   As String
    
    Dim sSubjCD     As String
    Dim sTmpSubjCD  As String
    
    Dim sLsnCD      As String
    Dim sTmpLsnCD   As String
    
    Dim sWeek       As String
    Dim sTmpWeek    As String
    Dim sLesson     As String
    Dim sTmpLesson  As String
    
    Dim nAtt        As Long
    Dim nRow        As Long
    Dim nRowChk     As Long
    Dim nCol        As Long
    Dim nColChk     As Long
    
    Dim sTmp        As String
    
    For nAtt = 1 To UBound(uAttend) Step 1
        
        sTcrCD = uAttend(nAtt).TCRCD                    '< 강사
        sSubjCD = uAttend(nAtt).SUBJCD                  '< 과목
        sLsnCD = uAttend(nAtt).LSNCD                    '< 반
        sWeek = uAttend(nAtt).WEEKS                     '< 요일
        sLesson = uAttend(nAtt).LESSON                  '< 교시
        
        
        If StrComp(sLsnCD, "XXXXX", vbTextCompare) = 0 And _
           StrComp(sWeek, "0", vbTextCompare) = 0 And _
           StrComp(sLesson, "0", vbTextCompare) = 0 Then
           
            'sprNot
            sprNot.MaxRows = sprNot.MaxRows + 1
            sprNot.Row = sprNot.MaxRows
            
            sprNot.Col = 1:                 Call basFunction.Set_SprType_Text(sprNot, "center", "left", 100, sTcrCD)
            sprNot.Col = sprNot.Col + 1:    Call basFunction.Set_SprType_Text(sprNot, "center", "left", 100, sSubjCD)
            sprNot.Col = sprNot.Col + 1:    Call basFunction.Set_SprType_Text(sprNot, "center", "left", 100, uAttend(nAtt).TCRNM)
            sprNot.Col = sprNot.Col + 1:    Call basFunction.Set_SprType_Text(sprNot, "center", "left", 100, uAttend(nAtt).SUBJNM)
            
        Else
            For nRow = 1 To sprTmr.MaxRows Step 1
                
                sTmpLesson = Trim(CLng(nRow))
                
                If StrComp(sLesson, sTmpLesson, vbTextCompare) = 0 Then             '< 1. 교시가 맞음.
                    
                    nRowChk = nRow                                                      '< 교시의 행
                    
                    For nCol = 1 To sprTmr.MaxCols Step 1
                        sprTmr.Col = nCol:      nColChk = sprTmr.Col
                        sprTmr.Row = SpreadHeader + 1
                            sTmpWeek = Trim(sprTmr.Text)
                        sprTmr.Row = SpreadHeader + 3
                            sTmpLsnCD = Trim(sprTmr.Text)
                            
                        If StrComp(sWeek, sTmpWeek, vbTextCompare) = 0 And _
                           StrComp(sLsnCD, sTmpLsnCD, vbTextCompare) = 0 Then       '< 2. 요일 & 반
                            
                            sprTmr.Row = nRowChk
                            sprTmr.Col = nColChk
                            
                                sTmp = uAttend(nAtt).SUBJNM & vbCrLf & uAttend(nAtt).TCRNM
                                Call basFunction.Set_SprType_Text(sprTmr, "top", "left", 100, sTmp)
                                sprTmr.TypeEditMultiLine = True
                            
                            Exit For
                        End If
                        
                    Next nCol
                End If
            Next nRow
        End If
        
    Next nAtt
    
    With sprTmr
        .Row = 1:   .Row2 = .MaxRows
        .Col = 1:   .Col2 = .MaxCols
        .BlockMode = True
            .Lock = True
            .Protect = True
        .BlockMode = False
    End With
    
    With sprNot
        .Row = 1:   .Row2 = .MaxRows
        .Col = 1:   .Col2 = .MaxCols
        .BlockMode = True
            .Lock = True
            .Protect = True
        .BlockMode = False
    End With
    
End Sub







'>> 강의시간
Private Sub Find_TeachingTime()

    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    
    Dim nLength     As Long
    Dim sStr        As String
    Dim ni          As Integer
    Dim nRec        As Long
    
    If Me.Tag = "LOAD" Then Exit Sub
    
    On Error GoTo ErrStmt
    
    sStr = ""
    sStr = sStr & "         SELECT A.TCRCD , GET_TCRNM('" & Trim(basModule.SchCD) & "', A.TCRCD) AS TCRNM, "
    sStr = sStr & "                NVL(B.SUBJCD,'XX') AS SUBJCD, NVL(GET_SUBJNM('" & Trim(basModule.SchCD) & "', A.TCRCD, B.SUBJCD),'--') AS SUBJNM, "
    sStr = sStr & "                NVL(B.LSNCD,'XXXXX') AS LSNCD,"
    sStr = sStr & "                NVL(B.WEEKS,0) AS WEEKS,"
    sStr = sStr & "                NVL(B.LESSON,0) AS LESSON"
    sStr = sStr & "           FROM (SELECT A.TCRCD "
    sStr = sStr & "                   FROM SDTCR01TB A, SDTRX50TB B"
    sStr = sStr & "                  WHERE A.ACID  = B.ACID "
    sStr = sStr & "                    AND A.TCRCD = B.TCRCD"
    sStr = sStr & "                    AND A.ACID  = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "                  GROUP BY A.TCRCD "
    sStr = sStr & "                 ) A,"
    sStr = sStr & "                (SELECT A.TCRCD, A.SUBJCD, A.LSNCD, A.WEEKS, A.LESSON"
    sStr = sStr & "                   FROM (SELECT TCRCD, SUBJCD, LSNCD, WEEKS, LESSON"
    sStr = sStr & "                           FROM SDTRX50TB"
    sStr = sStr & "                          WHERE YM    = '" & Trim(fpYM.UnFmtText) & "'"
    sStr = sStr & "                            AND ACID  = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "                         ) A,"
    sStr = sStr & "                        (SELECT TCRCD, WEEKS, MIN(LESSON) AS LESSON"
    sStr = sStr & "                           FROM SDTRX50TB"
    sStr = sStr & "                          WHERE YM    = '" & Trim(fpYM.UnFmtText) & "'"
    sStr = sStr & "                            AND ACID  = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "                            AND (TCRCD, WEEKS)"
    sStr = sStr & "                             IN (SELECT TCRCD, WEEKS"
    sStr = sStr & "                                   FROM SDTRX50TB"
    sStr = sStr & "                                  WHERE YM    = '" & Trim(fpYM.UnFmtText) & "'"
    sStr = sStr & "                                    AND ACID  = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "                                  GROUP BY TCRCD, WEEKS"
    sStr = sStr & "                                 )"
    sStr = sStr & "                          GROUP BY TCRCD, WEEKS"
    sStr = sStr & "                         ) B"
    sStr = sStr & "                  WHERE A.TCRCD  = B.TCRCD"
    sStr = sStr & "                    AND A.WEEKS  = B.WEEKS"
    sStr = sStr & "                    AND A.LESSON = B.LESSON"
    
    Select Case Trim(Right(cboWeek.Text, 30))
            Case "ALL"
                'NO ACTION
            Case "1"
                sStr = sStr & "        AND B.WEEKS = 2 "
            Case "2"
                sStr = sStr & "        AND B.WEEKS = 3 "
            Case "3"
                sStr = sStr & "        AND B.WEEKS = 4 "
            Case "4"
                sStr = sStr & "        AND B.WEEKS = 5 "
            Case "5"
                sStr = sStr & "        AND B.WEEKS = 6 "
            Case "6"
                sStr = sStr & "        AND B.WEEKS = 7 "
            Case "7"
                sStr = sStr & "        AND B.WEEKS = 1 "
        End Select
    
    sStr = sStr & "                 ) B"
    sStr = sStr & "          WHERE A.TCRCD  = B.TCRCD (+)"
    sStr = sStr & "          ORDER BY A.TCRCD "
    
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
        .MoveFirst
        If .RecordCount > 0 Then
           
            ReDim uAttend(.RecordCount) As tAttend
            
            For nRec = 1 To .RecordCount Step 1
                uAttend(nRec).TCRCD = Trim(.Fields("TCRCD"))
                uAttend(nRec).TCRNM = Trim(.Fields("TCRNM"))
                
                uAttend(nRec).SUBJCD = Trim(.Fields("SUBJCD"))
                uAttend(nRec).SUBJNM = Trim(.Fields("SUBJNM"))
                
                uAttend(nRec).LSNCD = Trim(.Fields("LSNCD"))
                
                uAttend(nRec).WEEKS = Trim(.Fields("WEEKS"))
                uAttend(nRec).LESSON = Trim(.Fields("LESSON"))
                
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
    MsgBox "시간표 강사내역 조회시 에러가 발생하였습니다.", vbCritical + vbOKOnly, "시간표 강사내역 조회"

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
    
    Dim nMaxWeek    As Integer          '< 요일처리
    Dim nCol        As Long
    Dim sWeek       As String
    Dim sWeekChk    As String
    
    Dim nRow        As Long
    
    If Me.Tag = "LOAD" Then Exit Sub
    
    On Error GoTo ErrStmt
    
    sStr = ""
    sStr = sStr & "    SELECT ACID, LSNCD, LSNNM, LSNCDNM, "
    sStr = sStr & "           DECODE(KAEYOL,'01','인문',"
    sStr = sStr & "                         '02','자연',"
    sStr = sStr & "                         '03','예체') KAEYOL"
    sStr = sStr & "      FROM (SELECT ACID, LSNCD, LSNNM, LSNCDNM, KAEYOL "
    sStr = sStr & "              FROM SDLSN01TB "
    sStr = sStr & "             WHERE ACID = '" & Trim(basModule.SchCD) & "'"
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
        
            If Trim(Right(cboWeek.Text, 30)) <> "ALL" Then
                sprTmr.MaxCols = .RecordCount
                nMaxWeek = 1
            Else
                sprTmr.MaxCols = .RecordCount * 7
                nMaxWeek = 7
            End If
            sprTmr.ColHeaderRows = 6
                        
            For nCol = 1 To nMaxWeek
                .MoveFirst
            
                For nRec = 1 To .RecordCount Step 1
                    sprTmr.Col = .RecordCount * (nCol - 1) + nRec
                    
                    sWeekChk = ""
                    If Trim(Right(cboWeek.Text, 30)) = "ALL" Then
                        sWeekChk = Trim(CStr(nCol))
                    Else
                        sWeekChk = Trim(Right(cboWeek.Text, 30))
                    End If
                    
                    Select Case sWeekChk
                        Case "1"
                            sprTmr.Row = SpreadHeader:      sTmp = "월"
                                sprTmr.Text = sTmp:     sprTmr.ColWidth(sprTmr.Col) = 7
                                sprTmr.FontSize = 8
                                sprTmr.FontBold = False:    sprTmr.RowHeight(sprTmr.Row) = 15
                            sprTmr.Row = SpreadHeader + 1:  sTmp = "2"
                                sprTmr.Text = sTmp:     sprTmr.ColWidth(sprTmr.Col) = 7
                                sprTmr.FontSize = 8
                                sprTmr.FontBold = False:    sprTmr.RowHeight(sprTmr.Row) = 15
                        Case "2"
                            sprTmr.Row = SpreadHeader:      sTmp = "화"
                                sprTmr.Text = sTmp:     sprTmr.ColWidth(sprTmr.Col) = 7
                                sprTmr.FontSize = 8
                                sprTmr.FontBold = False:    sprTmr.RowHeight(sprTmr.Row) = 15
                            sprTmr.Row = SpreadHeader + 1:  sTmp = "3"
                                sprTmr.Text = sTmp:     sprTmr.ColWidth(sprTmr.Col) = 7
                                sprTmr.FontSize = 8
                                sprTmr.FontBold = False:    sprTmr.RowHeight(sprTmr.Row) = 15
                        Case "3"
                            sprTmr.Row = SpreadHeader:      sTmp = "수"
                                sprTmr.Text = sTmp:     sprTmr.ColWidth(sprTmr.Col) = 7
                                sprTmr.FontSize = 8
                                sprTmr.FontBold = False:    sprTmr.RowHeight(sprTmr.Row) = 15
                            sprTmr.Row = SpreadHeader + 1:  sTmp = "4"
                                sprTmr.Text = sTmp:     sprTmr.ColWidth(sprTmr.Col) = 7
                                sprTmr.FontSize = 8
                                sprTmr.FontBold = False:    sprTmr.RowHeight(sprTmr.Row) = 15
                        Case "4"
                            sprTmr.Row = SpreadHeader:      sTmp = "목"
                                sprTmr.Text = sTmp:     sprTmr.ColWidth(sprTmr.Col) = 7
                                sprTmr.FontSize = 8
                                sprTmr.FontBold = False:    sprTmr.RowHeight(sprTmr.Row) = 15
                            sprTmr.Row = SpreadHeader + 1:  sTmp = "5"
                                sprTmr.Text = sTmp:     sprTmr.ColWidth(sprTmr.Col) = 7
                                sprTmr.FontSize = 8
                                sprTmr.FontBold = False:    sprTmr.RowHeight(sprTmr.Row) = 15
                        Case "5"
                            sprTmr.Row = SpreadHeader:      sTmp = "금"
                                sprTmr.Text = sTmp:     sprTmr.ColWidth(sprTmr.Col) = 7
                                sprTmr.FontSize = 8
                                sprTmr.FontBold = False:    sprTmr.RowHeight(sprTmr.Row) = 15
                            sprTmr.Row = SpreadHeader + 1:  sTmp = "6"
                                sprTmr.Text = sTmp:     sprTmr.ColWidth(sprTmr.Col) = 7
                                sprTmr.FontSize = 8
                                sprTmr.FontBold = False:    sprTmr.RowHeight(sprTmr.Row) = 15
                        Case "6"
                            sprTmr.Row = SpreadHeader:      sTmp = "토"
                                sprTmr.Text = sTmp:     sprTmr.ColWidth(sprTmr.Col) = 7
                                sprTmr.FontSize = 8
                                sprTmr.FontBold = False:    sprTmr.RowHeight(sprTmr.Row) = 15
                            sprTmr.Row = SpreadHeader + 1:  sTmp = "7"
                                sprTmr.Text = sTmp:     sprTmr.ColWidth(sprTmr.Col) = 7
                                sprTmr.FontSize = 8
                                sprTmr.FontBold = False:    sprTmr.RowHeight(sprTmr.Row) = 15
                        Case "7"
                            sprTmr.Row = SpreadHeader:      sTmp = "일"
                                sprTmr.Text = sTmp:     sprTmr.ColWidth(sprTmr.Col) = 7
                                sprTmr.FontSize = 8
                                sprTmr.FontBold = False:    sprTmr.RowHeight(sprTmr.Row) = 15
                            sprTmr.Row = SpreadHeader + 1:  sTmp = "1"
                                sprTmr.Text = sTmp:     sprTmr.ColWidth(sprTmr.Col) = 7
                                sprTmr.FontSize = 8
                                sprTmr.FontBold = False:    sprTmr.RowHeight(sprTmr.Row) = 15
                    End Select
                    
                    sprTmr.Row = SpreadHeader + 2:  sTmp = "":  If IsNull(.Fields("KAEYOL")) = False Then sTmp = Trim(.Fields("KAEYOL"))
                        sprTmr.Text = sTmp
                        sprTmr.FontSize = 8
                        sprTmr.FontBold = False
                        
                        If nRec = 1 Then sKaeyol = sTmp
                        If StrComp(sKaeyol, sTmp, vbTextCompare) <> 0 Then
                            sprTmr.SetCellBorder sprTmr.Col, 1, sprTmr.Col, sprTmr.MaxRows, 1, basModule.SectionColor1, CellBorderStyleSolid
                            sKaeyol = sTmp
                        End If
                    
                    sprTmr.Row = SpreadHeader + 3:  sTmp = "":  If IsNull(.Fields("LSNCD")) = False Then sTmp = Trim(.Fields("LSNCD"))
                        sprTmr.Text = sTmp:     sprTmr.ColWidth(sprTmr.Col) = 7
                        sprTmr.FontSize = 8
                        sprTmr.FontBold = False
                    sprTmr.Row = SpreadHeader + 4:  sTmp = "":  If IsNull(.Fields("LSNCDNM")) = False Then sTmp = Trim(.Fields("LSNCDNM"))
                        sprTmr.Text = sTmp:     sprTmr.ColWidth(sprTmr.Col) = 7
                        sprTmr.FontSize = 8
                        sprTmr.FontBold = False
                    sprTmr.Row = SpreadHeader + 5:  sTmp = "":  If IsNull(.Fields("LSNNM")) = False Then sTmp = Trim(.Fields("LSNNM"))
                        sprTmr.Text = sTmp:     sprTmr.ColWidth(sprTmr.Col) = 7
                        sprTmr.FontSize = 12
                        sprTmr.FontBold = True
                    
                    .MoveNext
                Next nRec
            
            Next nCol
            
        End If
    End With
    
    With sprTmr
        If .MaxCols > 1 Then
            .Row = SpreadHeader
                .RowMerge = MergeAlways
            
            .AddCellSpan SpreadHeader, SpreadHeader, 1, 6
            
            .Row = SpreadHeader + 1:    .RowHidden = True
            .Row = SpreadHeader + 2:    .RowHidden = True
            .Row = SpreadHeader + 3:    .RowHidden = True
            
            .MaxRows = 10
            
            For nRow = 1 To .MaxRows Step 1
                .Row = nRow:    .RowHeight(.Row) = 30
            Next nRow
            
        End If
    End With

    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    Exit Sub
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    On Error GoTo 0
    MsgBox "시간표 반 조회시 에러가 발생하였습니다.", vbCritical + vbOKOnly, "ROW 헤더처리"

End Sub



























