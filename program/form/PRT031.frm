VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form PRT031 
   Caption         =   "시간표 출력 >> 강사 및 반별 시간표 배정내역 조회"
   ClientHeight    =   10125
   ClientLeft      =   1815
   ClientTop       =   3105
   ClientWidth     =   15900
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10125
   ScaleWidth      =   15900
   Begin VB.Frame Frame1 
      BackColor       =   &H0082C8E8&
      BorderStyle     =   0  '없음
      Caption         =   "Frame1"
      Height          =   555
      Left            =   60
      TabIndex        =   6
      Top             =   30
      Width           =   15435
      Begin VB.Frame Frame2 
         BackColor       =   &H00D2EAF5&
         BorderStyle     =   0  '없음
         Caption         =   "Frame2"
         Height          =   495
         Left            =   30
         TabIndex        =   7
         Top             =   30
         Width           =   15375
         Begin VB.ComboBox cboKaeyol 
            Height          =   300
            Left            =   3990
            Style           =   2  '드롭다운 목록
            TabIndex        =   2
            Top             =   75
            Width           =   1065
         End
         Begin VB.CommandButton cmdFindTmr 
            Caption         =   "조 회 (&F)"
            Height          =   405
            Left            =   1860
            TabIndex        =   0
            Top             =   30
            Width           =   1515
         End
         Begin VB.TextBox txtTcrNM 
            Height          =   300
            IMEMode         =   10  '한글 
            Left            =   6390
            TabIndex        =   4
            Text            =   "txtTcrNM"
            Top             =   75
            Width           =   1455
         End
         Begin EditLib.fpMask fpTcrCD 
            Height          =   300
            Left            =   5760
            TabIndex        =   3
            Top             =   75
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
         Begin EditLib.fpMask fpYM 
            Height          =   285
            Left            =   540
            TabIndex        =   1
            Top             =   90
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
         Begin VB.Label Label2 
            BackStyle       =   0  '투명
            Caption         =   "배정되지 않은 시수은 (-) 이며, 추가배정된 내용은 (+) 로 표시됩니다."
            ForeColor       =   &H000000C0&
            Height          =   210
            Left            =   8220
            TabIndex        =   12
            Top             =   60
            Width           =   6645
         End
         Begin VB.Label Label1 
            BackStyle       =   0  '투명
            Caption         =   "배정이 되지 않았거나, 초과배정된 내용만 조회합니다."
            Height          =   210
            Left            =   10950
            TabIndex        =   11
            Top             =   270
            Width           =   4425
         End
         Begin VB.Label Label4 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "계열"
            Height          =   210
            Left            =   2970
            TabIndex        =   9
            Top             =   120
            Width           =   975
         End
         Begin VB.Label Label26 
            Alignment       =   1  '오른쪽 맞춤
            BackStyle       =   0  '투명
            Caption         =   "강사명"
            Height          =   210
            Left            =   4920
            TabIndex        =   8
            Top             =   120
            Width           =   765
         End
      End
   End
   Begin FPSpread.vaSpread sprTcr 
      Height          =   2655
      Left            =   90
      TabIndex        =   10
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
      SpreadDesigner  =   "PRT031.frx":0000
   End
   Begin FPSpread.vaSpread sprTmr 
      Height          =   9435
      Left            =   30
      TabIndex        =   5
      Top             =   630
      Width           =   15465
      _Version        =   393216
      _ExtentX        =   27279
      _ExtentY        =   16642
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
      SpreadDesigner  =   "PRT031.frx":17EE
   End
End
Attribute VB_Name = "PRT031"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################################################
'   시 스 템  명 : 대성학원 입학사정, 반배정 & 시간표 프로그램
'   서브시스템명 :
'   모   듈   명 : PRT031
'   모 듈  목 적 : 강사 및 반별 시간표 배정내역 조회
'
'   작   성   일 : 2008/02/21
'   작   성   자 : 유하균
' --------------------------------------------------------------------------------------------------------------
' --------------------------------------------------------------------------------------------------------------
'                 수     정     내     용
' --------------------------------------------------------------------------------------------------------------
'   1. 수정일 :
'   2. 내  용 :
'################################################################################################################

Option Explicit

Private Type tTCR
    TCRCD       As String
    TCRNM       As String
    
    SUBJCD      As String
    SUBJNM      As String
    
    LSNCD       As String
    
    CHKS        As Long
End Type
Private uTCR()       As tTCR

Private Sub Form_Load()
    
    Me.Move 0, 0, 15700, 10550
    
    fpYM.Text = Format(Now, "yyyymm")
    
    basFunction.RemoveContextMenu txtTcrNM
    
    With sprTmr
        .ShadowColor = basModule.ShadowColor2
        .ShadowDark = basModule.ShadowDark2
        .ShadowText = basModule.ShadowText2
        .GridColor = basModule.GridColor2
        .GrayAreaBackColor = basModule.GrayAreaBackColor2
        
        .MaxRows = 0
        .MaxCols = 0
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
        
        .Visible = False
    End With
    
    With cboKaeyol
        .Clear
        .AddItem "전체" & Space(30) & "ALL"
        .AddItem "인문" & Space(30) & "01"
        .AddItem "자연" & Space(30) & "02"
        '.AddItem "예체" & Space(30) & "03"
        
        .ListIndex = 0
    End With
    
    fpTcrCD.Text = ""
    txtTcrNM.Text = ""
    
    ReDim uTCR(0) As tTCR
    
End Sub




'>> 조회 guide line
Private Sub sprTmr_Click(ByVal Col As Long, ByVal Row As Long)
    
    With sprTmr
        If Row < 2 Then Exit Sub
        If Col < 1 Then Exit Sub
    
        .Row = 2:       .Row2 = .MaxRows
        .Col = 1:       .Col2 = .MaxCols
        .BlockMode = True
            .BackColor = basModule.WhiteColor
            .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
        
        '>> 첫번째 행 색
        .Row = 1:       .Row2 = 1
        .Col = 1:       .Col2 = .MaxCols
        .BlockMode = True
            .Lock = True
            .Protect = True
            
            .BackColor = basModule.SelectColor1
            .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
        
        
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
        
        .Row = 2:       .Row2 = .MaxRows
        .Col = 1:       .Col2 = .MaxCols
        .BlockMode = True
            .BackColor = basModule.WhiteColor
            .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
        
        '>> 첫번째 행 색
        .Row = 1:       .Row2 = 1
        .Col = 1:       .Col2 = .MaxCols
        .BlockMode = True
            .Lock = True
            .Protect = True
            
            .BackColor = basModule.SelectColor1
            .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
        
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
            If .value > 0 Then
                .Row = .ActiveRow
                .Col = .MaxCols
                    .value = 1
            End If
        End If
        
    End With
    
End Sub

'######################################################################################################################
' 시수조회
'######################################################################################################################
Private Sub cmdFindTmr_Click()
    Dim nCol        As Long
    Dim nColChk     As Long
    
    sprTmr.MaxRows = 0
    sprTmr.MaxCols = 0
    
    sprTmr.Col = 0:   sprTmr.ColHidden = False
    sprTmr.Row = 0:   sprTmr.RowHidden = False
    
    sprTmr.RowHeaderCols = 1
    sprTmr.ColHeaderRows = 1
    
    ReDim uTCR(0) As tTCR           '< 초기화
    
    'sprTmr.Visible = False
    Call Display_SprTmr_Row_SpreadHeader                    '<< COL 로 진행하는 ROW의 헤더 작성
    Call Display_SprTmr_Col_SpreadHeader                    '<< ROW 로 진행하는 COL의 헤더 작성
    
    'sprTmr.Visible = True
    
    If sprTmr.MaxCols > 1 Then
        
        '>> 시수내역 조회 : 정상내역 제외한 내역
        Call Find_SisuTime
        
        If UBound(uTCR) > 0 Then
            '> 내역 보여주기
            Call Show_SisuTime
            
            '> 전체 시수보기
            Call Show_SisuTot
            
        End If
        
    End If
End Sub


'>> 전체 시수보기
Private Sub Show_SisuTot()
    
    Dim DBCmd       As ADODB.Command
    Dim DBRec       As ADODB.Recordset
    Dim DBParam     As ADODB.Parameter
    
    Dim nLength     As Long
    Dim sStr        As String
    Dim ni          As Integer
    Dim nRec        As Long
    
    Dim sLsnCD      As String
    Dim sTmpLsnCD   As String
    Dim nSisu       As Long
    
    Dim nRow        As Long
    Dim nCol        As Long
    
    Dim sTmp        As String
    Dim sKaeyol     As String
    
    On Error GoTo ErrStmt
    
    sStr = ""
    sStr = sStr & "     SELECT LSNCD, SUM(NVL(SISU,0))  AS SISU "
    sStr = sStr & "       FROM SDTRX50TB"
    sStr = sStr & "      WHERE ACID = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "      GROUP BY LSNCD "
      
    Set DBCmd = New ADODB.Command
    Set DBRec = New ADODB.Recordset
    Set DBParam = New ADODB.Parameter
    
    DBCmd.ActiveConnection = basDataBase.DBConn             '<< DB connection
    DBCmd.CommandText = sStr
    DBCmd.CommandType = adCmdText
    DBCmd.CommandTimeout = 30
    
    '
    
        
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
           
            For nRec = 1 To .RecordCount Step 1
                sLsnCD = Trim(.Fields("LSNCD"))
                nSisu = CLng(.Fields("SISU"))
                
                For nCol = 1 To sprTmr.MaxCols Step 1
                    sprTmr.Row = SpreadHeader + 1
                    sprTmr.Col = nCol
                        sTmpLsnCD = Trim(sprTmr.Text)
                    
                    If StrComp(sLsnCD, sTmpLsnCD, vbTextCompare) = 0 Then       '< 2. 반
                        
                        sprTmr.Row = 1
                        sprTmr.Col = nCol
                        
                            Call basFunction.Set_SprType_Numeric(sprTmr, 0, -9999, 9999, "", nSisu)
                            
                        Exit For
                    End If
                Next nCol
                
                .MoveNext
            Next nRec
        End If
    End With
    
    With sprTmr
        .Row = 1:   .Row2 = .MaxRows
        .Col = 1:   .Col2 = .MaxCols
        .BlockMode = True
            .Lock = True
            .Protect = True
        .BlockMode = False
        
        '>> 첫번째 행 색
        .Row = 1:       .Row2 = 1
        .Col = 1:       .Col2 = .MaxCols
        .BlockMode = True
            .Lock = True
            .Protect = True
            
            .BackColor = basModule.SelectColor1
            .BackColorStyle = BackColorStyleUnderGrid
        .BlockMode = False
        
        sKaeyol = ""
        For nCol = 1 To .MaxCols Step 1
            .Row = SpreadHeader
            .Col = nCol
                
            If .Col Mod 5 = 0 Then
                .SetCellBorder .Col, 1, .Col, .MaxRows, 2, basModule.SectionColor2, CellBorderStyleSolid
            End If
                
            If sKaeyol = "" Then sKaeyol = Trim(.Text)
            
            If StrComp(sKaeyol, Trim(.Text), vbTextCompare) <> 0 Then
                .SetCellBorder .Col, 1, .Col, .MaxRows, 1, basModule.SectionColor1, CellBorderStyleSolid
                sKaeyol = Trim(.Text)
            End If
            
        Next nCol
        
        .SetCellBorder 1, 1, .MaxCols, 1, 8, basModule.SectionColor1, CellBorderStyleSolid
        
        For nRow = 1 To .MaxRows Step 1
            .Row = nRow
            If (.Row Mod 5) + 1 = 1 Then
                .SetCellBorder 1, .Row, .MaxCols, .Row, 8, basModule.SectionColor1, CellBorderStyleSolid
            End If
        Next nRow
        
    End With
    
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    Exit Sub
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    On Error GoTo 0
    MsgBox "전체시수 조회시 에러가 발생하였습니다.", vbCritical + vbOKOnly, "전체시수 조회"

End Sub



'>> 내역 보여주기
Private Sub Show_SisuTime()

    Dim sTcrCD      As String
    Dim sTmpTcrCD   As String
    
    Dim sSubjCD     As String
    Dim sTmpSubjCD  As String
    
    Dim sLsnCD      As String
    Dim sTmpLsnCD   As String
    
    Dim nSisu       As Long
    
    Dim nS          As Long
    Dim nRow        As Long
    Dim nRowChk     As Long
    Dim nCol        As Long
    Dim nColChk     As Long
    
    Dim sTmp        As String
    
    For nS = 1 To UBound(uTCR) Step 1
        sTcrCD = uTCR(nS).TCRCD                     '< 강사
        sSubjCD = uTCR(nS).SUBJCD                   '< 과목
        sLsnCD = uTCR(nS).LSNCD                     '< 반
        nSisu = uTCR(nS).CHKS                       '< 시수
        
        For nRow = 1 To sprTmr.MaxRows Step 1
        
            sprTmr.Row = nRow:      nRowChk = sprTmr.Row
            sprTmr.Col = SpreadHeader
                sTmpTcrCD = Trim(sprTmr.Text)
            sprTmr.Col = SpreadHeader + 1
                sTmpSubjCD = Trim(sprTmr.Text)
            
            If StrComp(sTcrCD, sTmpTcrCD, vbTextCompare) = 0 And _
               StrComp(sSubjCD, sTmpSubjCD, vbTextCompare) = 0 Then             '< 1. 강사 및 과목내역이 맞는 경우
               
                For nCol = 1 To sprTmr.MaxCols Step 1
                    sprTmr.Col = nCol:      nColChk = sprTmr.Col
                    sprTmr.Row = SpreadHeader + 1
                        sTmpLsnCD = Trim(sprTmr.Text)
                        
                    If StrComp(sLsnCD, sTmpLsnCD, vbTextCompare) = 0 Then       '< 2. 반
                        
                        sprTmr.Row = nRowChk
                        sprTmr.Col = nColChk
                        
                            Call basFunction.Set_SprType_Numeric(sprTmr, 0, -9999, 9999, "", nSisu)
                            
                        Exit For
                    End If
                    
                Next nCol
                
            End If
            
        Next nRow
    Next nS
    
    With sprTmr
        .Row = 1:   .Row2 = .MaxRows
        .Col = 1:   .Col2 = .MaxCols
        .BlockMode = True
            .Lock = True
            .Protect = True
        .BlockMode = False
    End With
    
End Sub


'>> 강의시간
Private Sub Find_SisuTime()

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
    sStr = sStr & "     SELECT TCRCD , GET_TCRNM('" & Trim(basModule.SchCD) & "',TCRCD) AS TCRNM, "
    sStr = sStr & "            SUBJCD, GET_SUBJNM('" & Trim(basModule.SchCD) & "',TCRCD,SUBJCD) AS SUBJNM, "
    sStr = sStr & "            LSNCD , CHKS"
    sStr = sStr & "       FROM (SELECT TCRCD, SUBJCD, LSNCD,"
    sStr = sStr & "                    SUM(NVL(TSISU,0)) AS TSISU, SUM(NVL(SSISU,0)) AS SSISU,"
    sStr = sStr & "                    SUM(NVL(SSISU,0))-SUM(NVL(TSISU,0)) AS CHKS"
    sStr = sStr & "               FROM (/* 강사시수 */"
    sStr = sStr & "                     SELECT TCRCD, SUBJCD, LSNCD, SISU AS TSISU, 0 AS SSISU"
    sStr = sStr & "                       FROM SDTCR11TB"
    sStr = sStr & "                      WHERE ACID  = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "                     UNION ALL"
    sStr = sStr & "                     /* 시간표 내역 */"
    sStr = sStr & "                     SELECT TCRCD, SUBJCD, LSNCD, 0 AS TSISU, SUM(SISU) AS SSISU"
    sStr = sStr & "                       FROM SDTRX50TB"
    sStr = sStr & "                      WHERE YM   = '" & Trim(fpYM.UnFmtText) & "'"
    sStr = sStr & "                        AND ACID = '" & Trim(basModule.SchCD) & "'"
    sStr = sStr & "                      GROUP BY TCRCD, SUBJCD, LSNCD"
    sStr = sStr & "                     )"
    sStr = sStr & "              GROUP BY TCRCD, SUBJCD, LSNCD"
    sStr = sStr & "             )"
    sStr = sStr & "      WHERE CHKS <> 0"

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
           
            ReDim uTCR(.RecordCount) As tTCR
            
            For nRec = 1 To .RecordCount Step 1
                uTCR(nRec).TCRCD = Trim(.Fields("TCRCD"))
                uTCR(nRec).TCRNM = Trim(.Fields("TCRNM"))
                
                uTCR(nRec).SUBJCD = Trim(.Fields("SUBJCD"))
                uTCR(nRec).SUBJNM = Trim(.Fields("SUBJNM"))
                
                uTCR(nRec).LSNCD = Trim(.Fields("LSNCD"))
                
                uTCR(nRec).CHKS = CLng(.Fields("CHKS"))
                
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
    MsgBox "강사시수 조회시 에러가 발생하였습니다.", vbCritical + vbOKOnly, "강사시수 조회"

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
'    If Trim(Right(cboSubjGbn.Text, 30)) <> "ALL" Then
'        sStr = sStr & " AND SUBJGBN = '" & Trim(Right(cboSubjGbn.Text, 30)) & "'"
'    End If
'    If Trim(Right(cboTcrGbn.Text, 30)) <> "ALL" Then
'        sStr = sStr & " AND TCRGBN  = '" & Trim(Right(cboTcrGbn.Text, 30)) & "'"
'    End If
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
    
    Dim nCol        As Long
    
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
        
            sprTmr.MaxCols = .RecordCount
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
            
        End If
    End With
    
    With sprTmr
        If .MaxCols > 1 Then
            
            '.AddCellSpan SpreadHeader, SpreadHeader, 1, 4
            '.Row = SpreadHeader + 1:    .RowHidden = True
            
        End If
    End With

    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    Exit Sub
ErrStmt:
    Set DBCmd = Nothing
    Set DBRec = Nothing
    
    On Error GoTo 0
    MsgBox "반 조회시 에러가 발생하였습니다.", vbCritical + vbOKOnly, "ROW 헤더처리"

End Sub











































'######################################################################################################################
'>> 강사조회
'######################################################################################################################
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
            sStr = sStr & "   ORDER BY TCRCD"
                
        Case vbKeyF10
            sStr = ""
            sStr = sStr & "  SELECT ACID, TCRCD, SUBJCD, SUBJGBN, TCRGBN, TCRNM, SUBJNM, TCR_CL"
            sStr = sStr & "    From SDTCR01TB"
            sStr = sStr & "   WHERE ACID   = '" & Trim(basModule.SchCD) & "'"
            If Trim(fpTcrCD.UnFmtText) > " " Then
                sStr = sStr & "     AND TCRNM  LIKE '" & Trim(txtTcrNM.Text) & "%'"
            End If
            sStr = sStr & "   ORDER BY TCRCD "
            
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
            sStr = sStr & "   ORDER BY TCRCD "
            
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
            sStr = sStr & "   ORDER BY TCRCD"
        Case vbKeyF10
            sStr = ""
            sStr = sStr & "  SELECT ACID, TCRCD, SUBJCD, SUBJGBN, TCRGBN, TCRNM, SUBJNM, TCR_CL"
            sStr = sStr & "    From SDTCR01TB"
            sStr = sStr & "   WHERE ACID   = '" & Trim(basModule.SchCD) & "'"
            If Trim(txtTcrNM.Text) > " " Then
                sStr = sStr & "     AND TCRNM  LIKE '" & Trim(txtTcrNM.Text) & "%'"
            End If
            sStr = sStr & "   ORDER BY TCRCD "
        
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
            sStr = sStr & "   ORDER BY TCRCD "
            
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



