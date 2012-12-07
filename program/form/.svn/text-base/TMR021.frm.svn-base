VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form TMR021 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  '단일 고정
   Caption         =   "시간표 만들기 >> 이동수업 시간표 등록 >> 수강가능 처리내역 보기"
   ClientHeight    =   5250
   ClientLeft      =   2370
   ClientTop       =   3270
   ClientWidth     =   10950
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   10950
   Begin FPSpread.vaSpread sprData 
      Height          =   5115
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   10815
      _Version        =   393216
      _ExtentX        =   19076
      _ExtentY        =   9022
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
      SpreadDesigner  =   "TMR021.frx":0000
   End
End
Attribute VB_Name = "TMR021"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################################################
'   시 스 템  명 : 대성학원 입학사정, 반배정 & 시간표 프로그램
'   서브시스템명 :
'   모   듈   명 : TRM021
'   모 듈  목 적 :
'
'   작   성   일 : 2007/11/06
'   작   성   자 : 유하균
' --------------------------------------------------------------------------------------------------------------
' --------------------------------------------------------------------------------------------------------------
'                 수     정     내     용
' --------------------------------------------------------------------------------------------------------------
'   1. 수정일 :
'   2. 내  용 :
'################################################################################################################

Option Explicit

Private Const nRowHeight = 14


Private Sub Form_Load()
    
    Me.Tag = "LOAD"
        With sprData
            .ShadowColor = basModule.ShadowColor1
            .ShadowDark = basModule.ShadowDark1
            .ShadowText = basModule.ShadowText1
            .GridColor = basModule.GridColor1
            .GrayAreaBackColor = basModule.GrayAreaBackColor1
        End With
    Me.Tag = ""
    
End Sub

Public Sub Show_TMR_WorkSheet_Data(ByRef aSpread As Object, ByVal aKaeyol As String)
    
    Dim nRow        As Long
    Dim nCol        As Long
    
    Dim sTmp        As String
    Dim nTmp        As Long
    
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2, 11070, 5655
    
    sprData.MaxRows = aSpread.MaxRows
    sprData.MaxCols = aSpread.MaxCols
    
    ' 헤더복사
    aSpread.Row = SpreadHeader
    sprData.Row = SpreadHeader:     sprData.RowHeight(sprData.Row) = nRowHeight
    
    For nCol = 1 To aSpread.MaxCols Step 1
        aSpread.Col = nCol
        sprData.Col = nCol
        
        Select Case sprData.Col
            Case 1
                sprData.ColWidth(sprData.Col) = 7
                sprData.Text = aSpread.Text
                sprData.ColHidden = True
                
            Case 2
                sprData.ColWidth(sprData.Col) = 8
                sprData.Text = aSpread.Text
                
            Case 3
                sprData.ColWidth(sprData.Col) = 8
                sprData.Text = aSpread.Text
                sprData.ColHidden = True
                
            Case Else
                Select Case aKaeyol
                    Case "01", "02"
                        sprData.ColWidth(sprData.Col) = 6
                        
                        sTmp = aSpread.Text
                        Select Case sTmp
                            Case "01":  sTmp = "국사"
                            Case "02":  sTmp = "윤리"
                            Case "03":  sTmp = "경제"
                            Case "04":  sTmp = "한근"
                            Case "05":  sTmp = "세계사"
                            Case "06":  sTmp = "경지"
                            Case "07":  sTmp = "한지"
                            Case "08":  sTmp = "정치"
                            Case "09":  sTmp = "사문"
                            Case "10":  sTmp = "법사"
                            Case "11":  sTmp = "세지"
                        End Select
                    Case "03"
                        sprData.ColWidth(sprData.Col) = 6
                        
                        sTmp = aSpread.Text
                        Select Case sTmp
                            Case "01":  sTmp = "물1"
                            Case "02":  sTmp = "화1"
                            Case "03":  sTmp = "생1"
                            Case "04":  sTmp = "지1"
                            Case "05":  sTmp = "물2"
                            Case "06":  sTmp = "화2"
                            Case "07":  sTmp = "생2"
                            Case "08":  sTmp = "지2"
                        End Select
                End Select
                sprData.Text = sTmp
                
        End Select
    Next nCol
    
    
    For nRow = 1 To aSpread.MaxRows Step 1
        For nCol = 1 To aSpread.MaxCols Step 1
            
            aSpread.Row = nRow
            sprData.Row = nRow
            
            aSpread.Col = nCol
            sprData.Col = nCol
            
            Select Case aSpread.Col
                Case 1, 2
                    sTmp = Trim(aSpread.Text)
                    Call basFunction.Set_SprType_Text(sprData, "center", "left", basFunction.LenKor(sTmp), sTmp)
                    
                Case Else
                    nTmp = aSpread.Value
                    Call basFunction.Set_SprType_Numeric(sprData, 0, -99999, 99999, "", nTmp)
                    
            End Select
            
        Next nCol
    Next nRow
    
    
End Sub



Private Sub sprData_Click(ByVal Col As Long, ByVal Row As Long)
    If Row < 1 Then Exit Sub
    If Col < 1 Then Exit Sub
    
    With sprData
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
        
    End With
End Sub
