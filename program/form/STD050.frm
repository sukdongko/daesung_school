VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form STD090 
   Caption         =   "입학사정 >> 등록취소 학생 엑셀로 삭제하기"
   ClientHeight    =   10065
   ClientLeft      =   6600
   ClientTop       =   4095
   ClientWidth     =   10545
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10065
   ScaleWidth      =   10545
   Begin VB.Frame Frame20 
      BackColor       =   &H0082C8E8&
      BorderStyle     =   0  '없음
      Caption         =   "Frame20"
      Height          =   9435
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   10395
      Begin VB.Frame Frame21 
         BackColor       =   &H00D2EAF5&
         BorderStyle     =   0  '없음
         Caption         =   "Frame21"
         Height          =   9375
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   10335
         Begin FPSpread.vaSpread sprData 
            Height          =   6495
            Left            =   6660
            TabIndex        =   7
            Top             =   2790
            Width           =   3645
            _Version        =   393216
            _ExtentX        =   6429
            _ExtentY        =   11456
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
            GrayAreaBackColor=   16777215
            MaxCols         =   3
            SpreadDesigner  =   "STD050.frx":0000
         End
         Begin VB.CommandButton cmdGetExcel 
            Caption         =   "엑셀자료 가져오기"
            Height          =   510
            Left            =   480
            TabIndex        =   3
            Top             =   30
            Width           =   1875
         End
         Begin VB.CommandButton cmdExcelSave 
            Caption         =   "등록취소 학생 삭제하기"
            Height          =   1110
            Left            =   7140
            TabIndex        =   2
            Top             =   1230
            Width           =   2625
         End
         Begin MSComDlg.CommonDialog dlgFile 
            Left            =   0
            Top             =   420
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin FPSpread.vaSpread sprExcel_STD_Data 
            Height          =   8115
            Left            =   60
            TabIndex        =   4
            Top             =   1200
            Width           =   6525
            _Version        =   393216
            _ExtentX        =   11509
            _ExtentY        =   14314
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
            GrayAreaBackColor=   16777215
            MaxCols         =   6
            ProcessTab      =   -1  'True
            SpreadDesigner  =   "STD050.frx":18EF
         End
         Begin EditLib.fpLongInteger fpTotCnt 
            Height          =   345
            Left            =   5460
            TabIndex        =   9
            Top             =   113
            Width           =   825
            _Version        =   196608
            _ExtentX        =   1455
            _ExtentY        =   609
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
            ButtonDefaultAction=   -1  'True
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483633
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            AlignTextH      =   2
            AlignTextV      =   1
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
            Text            =   "0"
            MaxValue        =   "2147483647"
            MinValue        =   "0"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ""
            UseSeparator    =   0   'False
            IncInt          =   1
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483633
            Appearance      =   0
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            ButtonColor     =   -2147483633
            AutoMenu        =   0   'False
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin EditLib.fpLongInteger fpProcCnt 
            Height          =   345
            Left            =   9420
            TabIndex        =   11
            Top             =   2400
            Width           =   825
            _Version        =   196608
            _ExtentX        =   1455
            _ExtentY        =   609
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
            ButtonDefaultAction=   -1  'True
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483633
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            AlignTextH      =   2
            AlignTextV      =   1
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
            Text            =   "0"
            MaxValue        =   "2147483647"
            MinValue        =   "0"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ""
            UseSeparator    =   0   'False
            IncInt          =   1
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483633
            Appearance      =   0
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            ButtonColor     =   -2147483633
            AutoMenu        =   0   'False
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '투명
            Caption         =   "처리인원"
            ForeColor       =   &H000040C0&
            Height          =   210
            Left            =   8670
            TabIndex        =   12
            Top             =   2490
            Width           =   975
         End
         Begin VB.Label Label5 
            BackStyle       =   0  '투명
            Caption         =   "조회인원"
            ForeColor       =   &H000040C0&
            Height          =   210
            Left            =   4530
            TabIndex        =   10
            Top             =   180
            Width           =   975
         End
         Begin VB.Label Label1 
            BackStyle       =   0  '투명
            Caption         =   "> 작업처리 현황"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H001E5A75&
            Height          =   285
            Left            =   6690
            TabIndex        =   8
            Top             =   2520
            Width           =   2625
         End
         Begin VB.Label Label29 
            BackStyle       =   0  '투명
            Caption         =   ">> 조회기본항목"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H001E5A75&
            Height          =   285
            Left            =   120
            TabIndex        =   6
            Top             =   120
            Width           =   2625
         End
         Begin VB.Label Label30 
            BackStyle       =   0  '투명
            Caption         =   $"STD050.frx":32BA
            Height          =   615
            Left            =   240
            TabIndex        =   5
            Top             =   630
            Width           =   5475
         End
      End
   End
End
Attribute VB_Name = "STD090"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################################################
'   시 스 템  명 : 대성학원 입학사정, 반배정 & 시간표 프로그램
'   서브시스템명 :
'   모   듈   명 : STD090
'   모 듈  목 적 :
'
'   작   성   일 : 2007/08/22
'   작   성   자 : 유하균
' --------------------------------------------------------------------------------------------------------------
' --------------------------------------------------------------------------------------------------------------
'                 수     정     내     용
' --------------------------------------------------------------------------------------------------------------
'   1. 수정일 :
'   2. 내  용 :
'################################################################################################################

Option Explicit


Private Type tExcel_StdData
    ACID        As String
    EXMID       As String
    STDNM       As String
    Birth_ymd       As String
    EXMTYPE     As String
    kaeyol      As String
    
End Type
Private uExcel_StdData      As tExcel_StdData



Private Sub Form_Load()
    Me.Move 0, 0, 10665, 9980
    
    sprExcel_STD_Data.MaxRows = 0
    sprData.MaxRows = 0
    
    fpTotCnt.value = 0
    fpProcCnt.value = 0
    
    With sprExcel_STD_Data
        .ShadowColor = basModule.ShadowColor1
        .ShadowDark = basModule.ShadowDark1
        .ShadowText = basModule.ShadowText1
        .GridColor = basModule.GridColor1
        .GrayAreaBackColor = basModule.GrayAreaBackColor1
    End With
    
    With sprData
        .ShadowColor = basModule.ShadowColor2
        .ShadowDark = basModule.ShadowDark2
        .ShadowText = basModule.ShadowText2
        .GridColor = basModule.GridColor2
        .GrayAreaBackColor = basModule.GrayAreaBackColor2
    End With
    
    
End Sub


'## EXCEL 자료조회
Private Sub cmdGetExcel_Click()
    
    On Error GoTo ErrStmt
    
    cmdGetExcel.Enabled = False
        Call Get_Excel_Data
        
    cmdGetExcel.Enabled = True
    
    Exit Sub
ErrStmt:
    MsgBox "엑셀자료 가져오는 중 오류가 발생하였습니다.", vbCritical + vbOKOnly, "학생 엑셀자료 가져오기"
    On Error GoTo 0
    
End Sub

Private Sub Get_Excel_Data()

    Dim sPath       As String
    
    ' Excel Data 처리
    Dim xlsDBConn   As ADODB.Connection
    Dim DBExCmd     As ADODB.Command
    Dim DBExRec     As ADODB.Recordset
    
    Dim sConn       As String
    Dim sSql        As String
    
    Dim nRow        As Long
    Dim sTmp        As String
    Dim nTmp        As Long
    
    Dim nJumsu      As Long
    Dim ni          As Long
    Dim nC          As Long
    
    On Error GoTo ErrStmt1
    
    With dlgFile
        .CancelError = True
        .fileName = ""
        .InitDir = App.Path
        .Filter = "EXCEL FILE(*.XLS)|*.XLS"
        .DefaultExt = "*.XLS"
        .ShowOpen
        
        If (.fileName) = "" Then
            MsgBox "선택한 파일이 없습니다.", vbExclamation + vbOKOnly, Me.Caption
            Exit Sub
        End If
        
        sPath = .fileName
        
    End With
    
    On Error GoTo 0
    
    On Error GoTo ErrStmt2                          '>> error 처리
    
    Set xlsDBConn = New ADODB.Connection
    sConn = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
            "Data Source=" & sPath & ";" & _
            "Extended Properties=""Excel 8.0;HDR=no;"";"
    
    With xlsDBConn
        .ConnectionString = sConn                   ' 데이터베이스와 연결을 시도합니다.
        .ConnectionTimeout = 30                     ' 제한 시간내에 연결이 되지 않으면 자동으로 끊습니다.
        .Properties("Prompt") = adPromptNever       ' 이것은 ADO에서 기본 프롬프트 모드입니다.
        .CursorLocation = adUseClient               ' 커서위치를 Client 쪽에 넣습니다.
        
        .Open                                       ' 데이터베이스를 엽니다.
        
        Do While .State And adStateConnecting
            DoEvents
        Loop
    End With
       
       
    fpTotCnt.value = 0
    
'>> 엑셀 DB Open
    sSql = ""
    sSql = sSql & " SELECT * "
    sSql = sSql & "   FROM [Sheet1$] "
    
    Set DBExCmd = New ADODB.Command
    Set DBExRec = New ADODB.Recordset
    
    DBExCmd.ActiveConnection = xlsDBConn
    DBExCmd.CommandText = sSql
    DBExCmd.CommandType = adCmdText
    DBExCmd.CommandTimeout = 30
    
    DBExRec.Open DBExCmd, , adOpenStatic, adLockReadOnly, -1
    Do While xlsDBConn.State And adStateExecuting
        DoEvents
    Loop
    
    If DBExRec.RecordCount = 0 Then
        Set DBExCmd = Nothing
        Set DBExRec = Nothing
        Set xlsDBConn = Nothing
        
        MsgBox "Excel Data가 없습니다.", vbExclamation + vbOKOnly, "IT2007"
        Exit Sub
    End If
        
    
    sprExcel_STD_Data.MaxRows = 0       ' 초기화
    
    
    DBExRec.MoveFirst
        
    '## header 1 line skip
    DBExRec.MoveNext
    
    
    For nRow = 2 To DBExRec.RecordCount Step 1
    '학원코드
        sTmp = "":  If IsNull(DBExRec.Fields(0)) = False Then sTmp = UCase(Trim(DBExRec.Fields(0)))
        uExcel_StdData.ACID = sTmp
    '수험번호
        sTmp = "":  If IsNull(DBExRec.Fields(1)) = False Then sTmp = Trim(DBExRec.Fields(1))
        uExcel_StdData.EXMID = sTmp
    '학생명
        sTmp = "":  If IsNull(DBExRec.Fields(2)) = False Then sTmp = Trim(DBExRec.Fields(2))
        uExcel_StdData.STDNM = sTmp
    '생년월일
        sTmp = "":  If IsNull(DBExRec.Fields(3)) = False Then sTmp = Trim(DBExRec.Fields(3))
        sTmp = Replace(sTmp, "-", "", 1, -1, vbTextCompare)
        If basFunction.LenKor(sTmp) > 6 Then
            sTmp = Left(sTmp, 4) & "-" & Mid(sTmp, 5, 2) & "-" & Mid(sTmp, 7, 2)
        End If
        uExcel_StdData.Birth_ymd = sTmp
    '유.무시험
        sTmp = "1"
        If IsNull(DBExRec.Fields(4)) = False Then
            sTmp = UCase(Trim(DBExRec.Fields(4)))
            Select Case sTmp
                Case "0", "1"
                    'no action
                Case Else
                    sTmp = "1"
                    
            End Select
        End If
        uExcel_StdData.EXMTYPE = sTmp
    '계열
        sTmp = "01"
        If IsNull(DBExRec.Fields(5)) = False Then
            sTmp = UCase(Trim(DBExRec.Fields(5)))
            Select Case sTmp
                Case "1" To "9"
                    sTmp = Format(sTmp, "00")
                Case "인문", "인"
                    sTmp = "01"
                Case "자연", "자"
                    sTmp = "02"
                Case "특인", "특별인문"
                    sTmp = "03"
                Case "특자", "특별자연"
                    sTmp = "04"
                    
                Case "수능인문"
                    sTmp = "05"
                Case "수능자연"
                    sTmp = "06"
                Case "수리나형"
                    sTmp = "08"
                    
                Case Else
                    sTmp = "01"
            End Select
        End If
        uExcel_StdData.kaeyol = sTmp
        
    
        
        
    '## 스프레드에 데이터 넣기 --------------------------------------------------------------------
        With sprExcel_STD_Data
        
            fpTotCnt.value = fpTotCnt.value + 1
            
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows:            .RowHeight(.Row) = 13
            
            '>> 학원
                .Col = 1
                    sTmp = uExcel_StdData.ACID
                    Call basFunction.Set_SprType_Text(sprExcel_STD_Data, "center", "left", basFunction.LenKor(sTmp), sTmp)
                    
            '>> 수험번호
                .Col = .Col + 1
                    sTmp = uExcel_StdData.EXMID
                    Call basFunction.Set_SprType_Text(sprExcel_STD_Data, "center", "left", basFunction.LenKor(sTmp), sTmp)
            '>> 학생명
                .Col = .Col + 1
                    sTmp = uExcel_StdData.STDNM
                    Call basFunction.Set_SprType_Text(sprExcel_STD_Data, "center", "left", basFunction.LenKor(sTmp), sTmp)
            '>> 생년월일
                .Col = .Col + 1
                    sTmp = Replace(uExcel_StdData.Birth_ymd, "-", "", 1, -1, vbTextCompare)
                    Call basFunction.Set_SprType_Text(sprExcel_STD_Data, "center", "left", basFunction.LenKor(sTmp), sTmp)
            '>> 유.무시험
                .Col = .Col + 1
                    sTmp = uExcel_StdData.EXMTYPE
                    Call basFunction.Set_SprType_Text(sprExcel_STD_Data, "center", "left", basFunction.LenKor(sTmp), sTmp)
            '>> 계열
                .Col = .Col + 1
                    sTmp = uExcel_StdData.kaeyol
                    Call basFunction.Set_SprType_Text(sprExcel_STD_Data, "center", "left", basFunction.LenKor(sTmp), sTmp)
                    
            
        End With
        
        DBExRec.MoveNext
        
    Next nRow
    
    
    
    With sprExcel_STD_Data
        If .MaxRows > 0 Then
            .Row = 1:   .Row2 = .MaxRows
            .Col = 1:   .Col2 = .MaxCols
            .BlockMode = True
                .BackColor = basModule.WhiteColor
                .BackColorStyle = BackColorStyleUnderGrid
            .BlockMode = False
            
            '.ColsFrozen = 3
            '.SetCellBorder 3, 1, 3, .MaxRows, 2, basModule.SectionColor1, CellBorderStyleSolid
            
        End If
    End With

    
    Set DBExRec = Nothing
    Set DBExCmd = Nothing
    Set xlsDBConn = Nothing
    
    MsgBox "학생 엑셀자료를 가지고 왔습니다.", vbInformation + vbOKOnly, Me.Caption
    
    On Error GoTo 0
    Exit Sub
ErrStmt1:
    MsgBox "엑셀 파일선택을 하십시요.", vbExclamation + vbOKOnly, Me.Caption
    Exit Sub
ErrStmt2:
    Set DBExRec = Nothing
    Set DBExCmd = Nothing
    xlsDBConn.Close
    Set xlsDBConn = Nothing
    
    MsgBox "EXCEL 자료 Open시 에러가 발생하였습니다.", vbCritical + vbOKOnly, Me.Caption
    On Error GoTo 0
    Exit Sub
End Sub








Private Sub sprExcel_STD_Data_Click(ByVal Col As Long, ByVal Row As Long)
    If Row < 1 Then Exit Sub
    
    With sprExcel_STD_Data
        If Trim(.Tag) = "" Then .Tag = "1"
        
        .Row = CLng(.Tag):      .Row2 = .Row
        .Col = 1:               .Col2 = .MaxCols
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

Private Sub sprExcel_STD_Data_KeyUp(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        Case vbKeyDelete
            With sprExcel_STD_Data
                .Row = .ActiveRow
                
                .DeleteRows .Row, 1
                .MaxRows = .MaxRows - 1
            End With
    End Select
End Sub







'>> 등록취소 학생 삭제하기
Private Sub cmdExcelSave_Click()
    Dim bRet    As Boolean
    
    
    bRet = False
    
    If sprExcel_STD_Data.MaxRows = 0 Then
        MsgBox "처리할 학생목록이 없습니다.", vbExclamation + vbOKOnly, "등록취소 학생 삭제하기"
        Exit Sub
    End If
    
    If MsgBox("학생데이터가 삭제됩니다." & vbCrLf & _
              "계속진행하시겠습니까?", vbQuestion + vbYesNo, "등록취소 학생 삭제하기") = vbNo Then
         Exit Sub
    End If
              
              
    sprData.MaxRows = 0
    
    
    '1. 현재 스프레드 학생등록하기
    Me.MousePointer = vbHourglass
    
    bRet = Save_sprExcel_STD_Data
    If bRet = False Then
        MsgBox "삭제학생 등록시 에러가 발생하였습니다.", vbCritical + vbOKOnly, "등록취소 학생 삭제하기"
        
        Me.MousePointer = vbDefault
        Exit Sub
    End If
    
        
    If sprData.MaxRows > 0 Then
        '2. 학생 COPY
        '    CLSTD90TB에 해당하는 학생의 자료 CLSTD01TB -> CLSTD91TB 로 복사
        bRet = Copy_Std01_to_Std91
        If bRet = False Then
            MsgBox "학생 복사만 등록시 에러가 발생하였습니다.", vbCritical + vbOKOnly, "등록취소 학생 삭제하기"
            
            Me.MousePointer = vbDefault
            Exit Sub
        End If
        
        '3. 대상학생 삭제
        '    CLSTD01TB <- WHERE CLSTD90TB - SCHNO : 해당내용 삭제
        
        bRet = Delete_Std01
        If bRet = False Then
            MsgBox "대상학생 삭제시 에러가 발생하였습니다.", vbCritical + vbOKOnly, "등록취소 학생 삭제하기"
            
            Me.MousePointer = vbDefault
            Exit Sub
        End If
        
        '4. 시간표 대상자 삭제
        '    CLTTL01TB <- WHERE CLSTD90TB - SCHNO : 해당내용 삭제
        bRet = Delete_Ttl01
        If bRet = False Then
            MsgBox "등록취소는 완료하였습니다." & vbCrLf & _
                   "다만, 시간표 대상자 삭제시 처리되지 않았으니 확인하십시요.", vbExclamation + vbOKOnly, "등록취소 학생 삭제하기"
        
            Me.MousePointer = vbDefault
            Exit Sub
        End If
    End If
    
    sprExcel_STD_Data.MaxRows = 0
    
    MsgBox "완료하였습니다.", vbInformation + vbOKOnly, "등록취소 학생 삭제하기"
    Me.MousePointer = vbDefault
    
End Sub


'<< 4. 대상학생 삭제
'    CLTTL01TB <- WHERE CLSTD90TB - SCHNO : 해당내용 삭제
Private Function Delete_Ttl01() As Boolean
    
    Dim bRet        As Boolean
    
    Dim sStr        As String
    Dim DBCmd       As ADODB.Command
    
    Dim nRow        As Long
    
    Dim nE          As Long
    Dim nExe        As Long
    Dim nTot        As Long
    Dim sTmp        As String
    
    bRet = False
    
    On Error GoTo ErrStmt
    
    basDataBase.DBConn.BeginTrans
    
    Set DBCmd = New ADODB.Command
    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
    
    With sprData
        
        bRet = False
        nTot = 0
        nExe = 0
        
        For nRow = 1 To .MaxRows Step 1
            
            nTot = nTot + 1
        
        '> 시간표 학생 테이블 ------------------------------------------------------------------
            .Row = nRow
            sStr = ""
            sStr = sStr & "      DELETE "
            sStr = sStr & "        FROM CLTTL01TB "
            .Col = 1
                sStr = sStr & "   WHERE SCHNO = '" & Trim(.Text) & "'"
            
            DBCmd.CommandText = sStr
            DBCmd.CommandType = adCmdText
            DBCmd.CommandTimeout = 30
            
            DBCmd.Execute nE, , -1
            
            Do While basDataBase.DBConn.State And adStateExecuting
                DoEvents
            Loop
                
            If nE = 1 Then
                nExe = nExe + 1
            End If
        '---------------------------------------------------------------------------------------
        Next nRow
    End With
    
    Set DBCmd = Nothing
    
    If nTot = nExe Then
        basDataBase.DBConn.CommitTrans
        Delete_Ttl01 = True
    Else
        basDataBase.DBConn.RollbackTrans
        Delete_Ttl01 = False
    End If
    
    Exit Function
    On Error GoTo 0
    
ErrStmt:
    basDataBase.DBConn.RollbackTrans

    Set DBCmd = Nothing
    Delete_Ttl01 = bRet
    
End Function


'<< 3. 대상학생 삭제
'    CLSTD01TB <- WHERE CLSTD90TB - SCHNO : 해당내용 삭제
Private Function Delete_Std01() As Boolean
    
    Dim bRet        As Boolean
    
    Dim sStr        As String
    Dim DBCmd       As ADODB.Command
    
    Dim nRow        As Long
    
    Dim nE          As Long
    Dim nExe        As Long
    Dim nTot        As Long
    Dim sTmp        As String
    
    bRet = False
    
    On Error GoTo ErrStmt
    
    basDataBase.DBConn.BeginTrans
    
    Set DBCmd = New ADODB.Command
    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
    
    With sprData
        
        bRet = False
        nTot = 0
        nExe = 0
        
        For nRow = 1 To .MaxRows Step 1
            
            nTot = nTot + 1
            
        '> 학생 테이블 -------------------------------------------------------------------------
            .Row = nRow
            sStr = ""
            sStr = sStr & "      DELETE "
            sStr = sStr & "        FROM CLSTD01TB "
            .Col = 1
                sStr = sStr & "   WHERE SCHNO = '" & Trim(.Text) & "'"
            
            DBCmd.CommandText = sStr
            DBCmd.CommandType = adCmdText
            DBCmd.CommandTimeout = 30
            
            DBCmd.Execute nE, , -1
            
            Do While basDataBase.DBConn.State And adStateExecuting
                DoEvents
            Loop
                
            
            If nE = 1 Then
                nExe = nExe + 1
            End If
        '---------------------------------------------------------------------------------------
        
        Next nRow
    End With
    
    Set DBCmd = Nothing
    
    If nTot = nExe Then
        basDataBase.DBConn.CommitTrans
        Delete_Std01 = True
    Else
        basDataBase.DBConn.RollbackTrans
        Delete_Std01 = False
    End If
    
    Exit Function
    On Error GoTo 0
    
ErrStmt:
    basDataBase.DBConn.RollbackTrans

    Set DBCmd = Nothing
    Delete_Std01 = bRet
    
End Function




'<< 2. 학생 copy
'    CLSTD90TB에 해당하는 학생의 자료 CLSTD01TB -> CLSTD91TB 로 복사
Private Function Copy_Std01_to_Std91() As Boolean
    
    Dim bRet        As Boolean
    
    Dim sStr        As String
    Dim DBCmd       As ADODB.Command
    
    Dim nRow        As Long
    
    Dim nE          As Long
    Dim nExe        As Long
    Dim nTot        As Long
    Dim sTmp        As String
    
    bRet = False
    
    On Error GoTo ErrStmt
    
    basDataBase.DBConn.BeginTrans
    
    Set DBCmd = New ADODB.Command
    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
    
    With sprData
        
        bRet = False
        nTot = 0
        nExe = 0
        
        For nRow = 1 To .MaxRows Step 1
            nE = 0
            
            nTot = nTot + 1
            
            .Row = nRow
            sStr = ""
            sStr = sStr & "      INSERT INTO CLSTD91TB "
            sStr = sStr & "      SELECT * "
            sStr = sStr & "        FROM CLSTD01TB "
            .Col = 1
                sStr = sStr & "   WHERE SCHNO = '" & Trim(.Text) & "'"
            
            
            DBCmd.CommandText = sStr
            DBCmd.CommandType = adCmdText
            DBCmd.CommandTimeout = 30
            
            DBCmd.Execute nE, , -1
            
            Do While basDataBase.DBConn.State And adStateExecuting
                DoEvents
            Loop
                
            If nE = 1 Then
                nExe = nExe + 1

            End If
        Next nRow
    End With
    
    Set DBCmd = Nothing
    
    If nTot = nExe Then
        basDataBase.DBConn.CommitTrans
        Copy_Std01_to_Std91 = True
    Else
        basDataBase.DBConn.RollbackTrans
        Copy_Std01_to_Std91 = False
    End If
    
    Exit Function
    On Error GoTo 0
    
ErrStmt:
    basDataBase.DBConn.RollbackTrans

    Set DBCmd = Nothing
    Copy_Std01_to_Std91 = bRet
    
End Function




'>> 1. 현재 스프레드 학생등록하기
Private Function Save_sprExcel_STD_Data() As Boolean
    Dim bRet        As Boolean
    
    Dim sStr        As String
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    
    Dim nRow        As Long
    Dim sSchNO      As String
    
    Dim nE          As Long
    Dim nExe        As Long
    Dim nTot        As Long
    Dim bSaveChk    As Boolean
    Dim sTmp        As String
    
    bRet = False
    
    On Error Resume Next
    
    basDataBase.DBConn.BeginTrans
    
    Set DBCmd = New ADODB.Command
    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
    
    With sprExcel_STD_Data
        
        nTot = 0
        nExe = 0
        
        fpProcCnt.value = 0
        
        For nRow = 1 To .MaxRows Step 1
            nE = 0
            
            .Row = nRow
            
            sStr = ""
            sStr = sStr & "      SELECT SCHNO"
            sStr = sStr & "        FROM CLSTD01TB "
            .Col = 2
                sStr = sStr & "   WHERE EXMID = '" & Trim(.Text) & "'"
            
            Set DBRec = New ADODB.Recordset
            
            DBCmd.CommandText = sStr
            DBCmd.CommandType = adCmdText
            DBCmd.CommandTimeout = 30
            
            DBRec.Open DBCmd, , adOpenStatic, adLockReadOnly, -1            '<< dynamic형태로 열게되면 record count를 할 수 없음.
            Do While DBRec.State And adStateExecuting
                DoEvents
            Loop
            
            bSaveChk = False
            
            
            If DBRec.RecordCount > 0 Then
                
                DBRec.MoveFirst
                
                If IsNull(DBRec.Fields("SCHNO")) = False Then
                    sSchNO = Trim(DBRec.Fields("SCHNO"))
                    
                    Set DBRec = Nothing
                    
                    sprData.MaxRows = sprData.MaxRows + 1
                    sprData.Row = sprData.MaxRows
                    
                    
                    sTmp = sSchNO
                            sprData.Col = 1
                                Call basFunction.Set_SprType_Text(sprData, "center", "left", basFunction.LenKor(sTmp), sTmp)
                            
                    .Col = 2
                        sTmp = Trim(.Text)
                            sprData.Col = 2
                                Call basFunction.Set_SprType_Text(sprData, "center", "left", basFunction.LenKor(sTmp), sTmp)
                            
                    .Col = 3
                        sTmp = Trim(.Text)
                            sprData.Col = 3
                                Call basFunction.Set_SprType_Text(sprData, "center", "left", basFunction.LenKor(sTmp), sTmp)
                    
                    nTot = nTot + 1
                    bSaveChk = True
                    
                Else
                    sTmp = " "
                            sprData.Col = 1
                                Call basFunction.Set_SprType_Text(sprData, "center", "left", basFunction.LenKor(sTmp), sTmp)
                            
                    .Col = 2
                        sTmp = Trim(.Text)
                            sprData.Col = 2
                                Call basFunction.Set_SprType_Text(sprData, "center", "left", basFunction.LenKor(sTmp), sTmp)
                            
                    .Col = 3
                        sTmp = "작업대상학생 없음"
                            sprData.Col = 3
                                Call basFunction.Set_SprType_Text(sprData, "center", "left", basFunction.LenKor(sTmp), sTmp)
                                
                    bSaveChk = False
                    
                End If
            End If
            
            
            If bSaveChk = True Then
                '<< INSERT
                sStr = ""
                sStr = sStr & "  INSERT INTO CLSTD90TB ("
                sStr = sStr & "         SCHNO  ,"
                sStr = sStr & "         ACID   ,"
                sStr = sStr & "         EXMID  ,"
                sStr = sStr & "         STDNM  ,"
                sStr = sStr & "         Birth_ymd  , EXMTYPE, KAEYOL"
                sStr = sStr & "  )"
                sStr = sStr & "  VALUES ( "
                
                    sStr = sStr & " '" & sSchNO & "', "             '< 학번
                .Col = 1
                    sStr = sStr & " '" & Trim(.Text) & "', "        '< 학원
                .Col = 2
                    sStr = sStr & " '" & Trim(.Text) & "', "        '< 수험번호
                .Col = 3
                    sStr = sStr & " '" & Trim(.Text) & "', "        '< 학생명
                .Col = 4
                    sStr = sStr & " '" & Replace(Trim(.Text), "-", "", 1, -1, vbTextCompare) & "', "        '< 주민번호
                .Col = 5
                    sStr = sStr & " '" & Trim(.Text) & "', "        '< 유/무시험
                .Col = 6
                    sStr = sStr & " '" & Trim(.Text) & "' "         '< 계열
                
                sStr = sStr & "  ) "

                DBCmd.CommandText = sStr
                DBCmd.CommandType = adCmdText
                DBCmd.CommandTimeout = 30
                
                DBCmd.Execute nE, , -1
                
                        
                Do While basDataBase.DBConn.State And adStateExecuting
                    DoEvents
                Loop
                
                If nE = 1 Then
                    nExe = nExe + 1
                    fpProcCnt.value = fpProcCnt.value + 1
                
            'update
                Else
                    
                    '<< update
                    sStr = ""
                    sStr = sStr & "  UPDATE CLSTD90TB "
                    .Col = 1
                    sStr = sStr & "     SET ACID    = '" & Trim(.Text) & "', "
                    .Col = 2
                    sStr = sStr & "         EXMID   = '" & Trim(.Text) & "', "
                    .Col = 3
                    sStr = sStr & "         STDNM   = '" & Trim(.Text) & "', "
                    .Col = 4
                    sStr = sStr & "         Birth_ymd   = '" & Trim(.Text) & "', "
                    .Col = 5
                    sStr = sStr & "         EXMTYPE = '" & Trim(.Text) & "', "
                    .Col = 6
                    sStr = sStr & "         KAEYOL  = '" & Trim(.Text) & "'"
                    sStr = sStr & "   WHERE SCHNO   = '" & sSchNO & "'"
                    
                    DBCmd.CommandText = sStr
                    DBCmd.CommandType = adCmdText
                    DBCmd.CommandTimeout = 30
                    
                    DBCmd.Execute nE, , -1
                    
                            
                    Do While basDataBase.DBConn.State And adStateExecuting
                        DoEvents
                    Loop
                    
                    If nE = 1 Then
                        nExe = nExe + 1
                        fpProcCnt.value = fpProcCnt.value + 1
                    End If
                End If
            End If
        Next nRow
    End With
    
    Set DBCmd = Nothing
    
    If nTot = nExe Then
        basDataBase.DBConn.CommitTrans
        Save_sprExcel_STD_Data = True
    Else
        basDataBase.DBConn.RollbackTrans
        Save_sprExcel_STD_Data = False
        fpProcCnt.value = 0
    End If
    
    Exit Function
    
    On Error GoTo 0
    
'ErrStmt:
'    basDataBase.DBConn.RollbackTrans
'
'    Set DBCmd = Nothing
'
'    Save_sprExcel_STD_Data = False
    
End Function


















































