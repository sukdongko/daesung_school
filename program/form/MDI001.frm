VERSION 5.00
Begin VB.MDIForm MDI001 
   BackColor       =   &H8000000C&
   Caption         =   "메인화면"
   ClientHeight    =   11370
   ClientLeft      =   2055
   ClientTop       =   2415
   ClientWidth     =   14580
   Icon            =   "MDI001.frx":0000
   LinkTopic       =   "MDIForm1"
   Begin VB.Menu mnu100 
      Caption         =   "입학사정"
      Begin VB.Menu mnuSTD010_N 
         Caption         =   "학생등록(노량진)"
      End
      Begin VB.Menu mnuSTD010 
         Caption         =   "학생 등록 (강남,양재,마강,송파,부산)"
      End
      Begin VB.Menu mnuSTD011_N 
         Caption         =   "학생전체 조회(노량진)"
      End
      Begin VB.Menu mnuSTD011 
         Caption         =   "학생전체 조회"
      End
      Begin VB.Menu mnuSTD012_N 
         Caption         =   "학생전체 조회 (차수별)(노량진,송파)"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSTD012 
         Caption         =   "학생전체 조회 (차수별)"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSTD092 
         Caption         =   "학생취소자 조회"
      End
      Begin VB.Menu mnu100_Line11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSTD200 
         Caption         =   "윈터 && 수학클리닉 등록 및 조회"
      End
      Begin VB.Menu mnu100_Line03 
         Caption         =   "-"
      End
      Begin VB.Menu mnuINT 
         Caption         =   "입학원서 출력"
         Begin VB.Menu mnuINT010_N 
            Caption         =   "종합 입학원서 출력 (노량진)"
         End
         Begin VB.Menu mnuINT010 
            Caption         =   "종합 입학원서 출력"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuINT110 
            Caption         =   "종합 입학원서 출력 (부산)"
         End
         Begin VB.Menu mnuINT112 
            Caption         =   "종합 입학원서 출력 (마강)"
         End
         Begin VB.Menu mnuINT113 
            Caption         =   "종합 입학원서 출력 (양재,송파)"
         End
         Begin VB.Menu mnuINT111 
            Caption         =   "종합 입학원서 출력 (강남)"
         End
         Begin VB.Menu mnu100_Line12 
            Caption         =   "-"
         End
         Begin VB.Menu mnuINT011 
            Caption         =   "선행 입학원서 출력 (노량진, 송파)"
         End
         Begin VB.Menu mnuINT012 
            Caption         =   "선행 입학원서 출력"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuINT013 
            Caption         =   "선행 입학원서 출력 (마이맥)"
         End
         Begin VB.Menu mnuINT014 
            Caption         =   "선행 입학원서 출력 (양재)"
         End
         Begin VB.Menu mnu100_Line05 
            Caption         =   "-"
         End
         Begin VB.Menu mnuINT020 
            Caption         =   "강남 윈터스쿨 출력"
         End
         Begin VB.Menu mnuINT021 
            Caption         =   "송파 윈터스쿨 출력"
         End
         Begin VB.Menu mnuINT022 
            Caption         =   "노량진 윈터스쿨 출력"
         End
         Begin VB.Menu mnuINT023 
            Caption         =   "양재 윈터스쿨 출력"
         End
         Begin VB.Menu mnuINT024 
            Caption         =   "부산 윈터스쿨 출력"
         End
         Begin VB.Menu mnu100_Line06 
            Caption         =   "-"
         End
         Begin VB.Menu mnuMAT010 
            Caption         =   "수학 집중 클릭닉 출력 (노량진, 송파)"
         End
         Begin VB.Menu mnuMAT010_J 
            Caption         =   "수학 집중 클리닉 출력 (양재)"
         End
      End
      Begin VB.Menu mnu100_Line01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSTD020 
         Caption         =   "합격생 등록"
      End
      Begin VB.Menu mnuSTD030 
         Caption         =   "등록금 및 가상계좌 부여"
      End
      Begin VB.Menu mnu100_Line02 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSTD040 
         Caption         =   "합격생 ▶ 시간표 작업으로 등록"
      End
      Begin VB.Menu mnu100_Line04 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSTD050 
         Caption         =   "등록취소 학생 엑셀로 삭제하기"
      End
   End
   Begin VB.Menu mnu150 
      Caption         =   "반 배정하기"
      Begin VB.Menu mnuLSN001 
         Caption         =   "반 등록"
      End
      Begin VB.Menu mnu150_Line01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLSN100 
         Caption         =   "반 구성하기"
      End
   End
   Begin VB.Menu mnu200 
      Caption         =   "이동시간표 만들기"
      Begin VB.Menu mnuLSN001_CP 
         Caption         =   "반 정보 등록"
      End
      Begin VB.Menu mnuLSN002 
         Caption         =   "이동 반 정보 등록"
      End
      Begin VB.Menu mnu200_Line01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLSN100_CP 
         Caption         =   "반 구성하기"
      End
      Begin VB.Menu mnu200_Line02 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTMR020 
         Caption         =   "이동수업 시간표 등록_OLD"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuTMR025 
         Caption         =   "이동수업 시간표 등록"
      End
      Begin VB.Menu mnuTMR027 
         Caption         =   "이동수업 과목내역 등록"
      End
   End
   Begin VB.Menu mnu250 
      Caption         =   "시간표 만들기"
      Begin VB.Menu mnuLSN001_CP2 
         Caption         =   "반 정보 등록"
      End
      Begin VB.Menu mnu250_Line01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMTX011 
         Caption         =   "구조별 시간표 등록"
      End
      Begin VB.Menu mnu250_Line02 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTMR011 
         Caption         =   "강사 및 강사별 과목넣기"
      End
      Begin VB.Menu mnuTMR012 
         Caption         =   "강사 시수넣기"
      End
      Begin VB.Menu mnuTMR015 
         Caption         =   "강사 강의불가 시간등록"
      End
      Begin VB.Menu mnu250_Line03 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTMR050 
         Caption         =   "전체시간표 구성"
      End
   End
   Begin VB.Menu mnu300 
      Caption         =   "시간표 출력"
      Begin VB.Menu mnuPRT011 
         Caption         =   "반별 시간표 출력 (노량진)"
      End
      Begin VB.Menu mnuPRT021 
         Caption         =   "강사별 시간표 출력 (노량진)"
      End
      Begin VB.Menu mnu300_Line01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPRT010 
         Caption         =   "반별 시간표 출력"
      End
      Begin VB.Menu mnuPRT020 
         Caption         =   "강사별 시간표 출력"
      End
      Begin VB.Menu mnu300_Line02 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPRT050 
         Caption         =   "빈 시간표 출력 (노량진)"
      End
      Begin VB.Menu mnu300_Line03 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTMR060 
         Caption         =   "강사 출석부"
      End
      Begin VB.Menu mnuPRT030 
         Caption         =   "강사 및 반별 시간표 배정내역 조회"
      End
   End
   Begin VB.Menu mnuEXM 
      Caption         =   "성적등록"
      Begin VB.Menu EXM010 
         Caption         =   "학생성적등록"
      End
      Begin VB.Menu EXM020 
         Caption         =   "학생성적출력"
      End
   End
   Begin VB.Menu TEST 
      Caption         =   "TEST"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "MDI001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################################################
'   시 스 템  명 : 대성학원 입학사정, 반배정 & 시간표 프로그램
'   서브시스템명 :
'   모   듈   명 : MDI001
'   모 듈  목 적 : 메인화면
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


'▤▤▤▤▤▤▤▤▤▤▤▤▤▤▤▤▤▤▤▤▤▤▤▤▤▤▤▤▤▤▤▤▤▤▤▤▤▤▤▤▤▤▤▤▤▤▤▤▤▤▤▤▤▤▤▤

Private Sub MenuYN()
    
    mnu100.Visible = False
    mnu150.Visible = False
    mnu200.Visible = False
    mnu250.Visible = False
    mnu300.Visible = False
    mnuEXM.Visible = False
    
    '노량진만 메뉴를 제어한다. 다른 학원들은 모든 메뉴 보임.
    
    If basModule.SchCD = "N" Then
        Select Case basModule.RegID
            Case "10000", "00002", "10003", "00001" '김영덕과장 (10000), 노량진(00002), 김병철(10003), ADMIN(00001)
                mnu100.Visible = True
                mnu150.Visible = True
                mnu200.Visible = True
                mnu250.Visible = True
                mnu300.Visible = True
                mnuEXM.Visible = True
                
            Case "10001"                            '신현우
                mnu100.Visible = True
                    mnuSTD010.Visible = True
                    mnuSTD011.Visible = True
                    mnuSTD012.Visible = True
                    mnuSTD092.Visible = True
                    
                    mnu100_Line11.Visible = False
                    mnuSTD200.Visible = False
                    mnu100_Line03.Visible = False
                    mnuINT.Visible = False
                    mnuSTD020.Visible = False
                    mnuSTD030.Visible = False
                    mnu100_Line02.Visible = False
                    mnuSTD040.Visible = False
                    mnu100_Line04.Visible = False
                    mnuSTD050.Visible = False
                    
                mnu150.Visible = False
                mnu200.Visible = False
                mnu250.Visible = False
                mnu300.Visible = False
                mnuEXM.Visible = True
                
                
            Case "10002"                            '정순택
                mnu100.Visible = True
                    mnuSTD010.Visible = True
                    mnuSTD011.Visible = True
                    mnuSTD012.Visible = True
                    mnuSTD092.Visible = True
                    
                    mnu100_Line11.Visible = False
                    mnuSTD200.Visible = False
                    mnu100_Line03.Visible = False
                    mnuINT.Visible = False
                    mnuSTD020.Visible = False
                    mnuSTD030.Visible = False
                    mnu100_Line02.Visible = False
                    mnuSTD040.Visible = False
                    mnu100_Line04.Visible = False
                    mnuSTD050.Visible = False
                
                mnu150.Visible = False
                mnu200.Visible = False
                mnu250.Visible = False
                mnu300.Visible = False
                mnuEXM.Visible = False
                
        End Select
    Else
        mnu100.Visible = True
        mnu150.Visible = True
        mnu200.Visible = True
        mnu250.Visible = True
        mnu300.Visible = True
        mnuEXM.Visible = True
        
    End If
    
End Sub

'▤▤▤▤▤▤▤▤▤▤▤▤▤▤▤▤▤▤▤▤▤▤▤▤▤▤▤▤▤▤▤▤▤▤▤▤▤▤▤▤▤▤▤▤▤▤▤▤▤▤▤▤▤▤▤▤



Private Sub MDIForm_Load()
    Dim sAcID               As String
    Dim sConnections        As String
    
    Select Case Trim(basModule.SchCD)
        Case "N"
            sAcID = "노량진"
        Case "K"
            sAcID = "강남"
        Case "S"
            sAcID = "송파"
        Case "P"
            sAcID = "송파mimac"
        Case "M"
            sAcID = "강남mimac"
        
        Case "W"
            sAcID = "주말법의대"
        Case "Q"
            sAcID = "야간법의대"
            
        Case "J"
            sAcID = "양재"
        
        Case "B"
            sAcID = "부산"
        
    End Select
    
    
    Select Case UCase(Trim(basModule.connDB))
        Case "MIMAC"
            sConnections = "실서버"
        Case "DEV"
            sConnections = "개발용"
        Case Else
            sConnections = "실서버"
    End Select
    
    '>> 작성버젼체크
    MDI001.Caption = "입학사정. 반배정 및 시간표 작성 프로그램 (multiple) - 10.11.19 pm 04:02 " & "【 " & sConnections & " 】" & "【 " & sAcID & " 】  " & App.ProductName & " - ver " & App.Major & "." & App.Minor & "." & App.Revision
    
    Call MenuYN
    
End Sub



Private Sub MDIForm_Unload(Cancel As Integer)
    Unload TMR021
    Unload TMR022
    Unload INT900
    Unload TMR028
    
End Sub























Private Sub mnuINT010_N_Click()
    Load INT110_N
    INT110_N.Show
    INT110_N.ZOrder 0
End Sub



Private Sub mnuINT112_Click()
    Load INT112
    INT112.Show
    INT112.ZOrder 0
End Sub

Private Sub mnuINT113_Click()
    Load INT113
    INT113.Show
    INT113.ZOrder 0
End Sub



'>> 입학사정 학생등록
Private Sub mnuSTD010_Click()
    Load STD010
    STD010.Show
    STD010.ZOrder 0
    
End Sub

Private Sub mnuSTD010_N_Click()
    Load STD010_N
    STD010_N.Show
    STD010_N.ZOrder 0
End Sub

'>> 학생전체 조회
Private Sub mnuSTD011_Click()
    Load STD011
    STD011.Show
    STD011.ZOrder 0
    
End Sub

Private Sub mnuSTD011_N_Click()
    Load STD011_N
    STD011_N.Show
    STD011_N.ZOrder 0
End Sub

Private Sub mnuSTD012_Click()
    Load STD012
    STD012.Show
    STD012.ZOrder 0
    
End Sub

Private Sub mnuSTD012_N_Click()
Load STD012_N
    STD012_N.Show
    STD012_N.ZOrder 0
End Sub

'>> 취소학생조회
Private Sub mnuSTD092_Click()
    Load STD092
    STD092.Show
    STD092.ZOrder 0
    
End Sub


'>> 입학원서 출력 // 선행 입학원서 출력
Private Sub mnuINT011_Click()
    Load INT011
    INT011.Show
    INT011.ZOrder 0
    
End Sub

Private Sub mnuINT013_Click()
    Load INT013
    INT013.Show
    INT013.ZOrder 0
    
End Sub
Private Sub mnuINT014_Click()
    Load INT014
    INT014.Show
    INT014.ZOrder 0
End Sub



Private Sub mnuINT110_Click()
    Load INT110
    INT110.Show
    INT110.ZOrder 0
    
End Sub

Private Sub mnuINT111_Click()
    Load INT111
    INT111.Show
    INT111.ZOrder 0
    
End Sub

'>> 입학원서 출력 // 강남 윈터스쿨
Private Sub mnuINT020_Click()
    Load INT020
    INT020.Show
    INT020.ZOrder 0
    
End Sub

'>> 입학원서 출력 // 송파 윈터스쿨
Private Sub mnuINT021_Click()
    Load INT021
    INT021.Show
    INT021.ZOrder 0
    
End Sub

'>> 입학원서 출력 // 노량진 윈터스쿨
Private Sub mnuINT022_Click()
    Load INT022
    INT022.Show
    INT022.ZOrder 0
    
End Sub

'>> 입학원서 출력 // 양재 윈터스쿨
Private Sub mnuINT023_Click()
    Load INT023
    INT023.Show
    INT023.ZOrder 0

End Sub

'>> 입학원서 출력 // 부산 윈터스쿨
Private Sub mnuINT024_Click()
    Load INT024
    INT024.Show
    INT024.ZOrder 0

End Sub



'>> 수학 집중 클리닉 입학원서
Private Sub mnuMAT010_Click()
    Load MAT010
    MAT010.Show
    MAT010.ZOrder 0
    
End Sub

'>> 수학 집중 클리닉 입학원서 양재
Private Sub mnuMAT010_J_Click()
    Load MAT011_J
    MAT011_J.Show
    MAT011_J.ZOrder 0
End Sub

'>> 합격생 등록
Private Sub mnuSTD020_Click()
    Load STD020
    STD020.Show
    STD020.ZOrder 0
    
End Sub

'>> 등록금 및 가상계좌 부여
Private Sub mnuSTD030_Click()
    Load STD031
    STD031.Show
    STD031.ZOrder 0
    
End Sub

'>> 등록취소생 엑셀로 삭제하기
Private Sub mnuSTD090_Click()
    Load STD090
    STD090.Show
    STD090.ZOrder 0
    
End Sub



'>> 등록취소 학생 엑셀로 삭제하기
Private Sub mnuSTD050_Click()
    Load STD090
    STD090.Show
    STD090.ZOrder 0
    
End Sub

'>> 학생 시간표 작업으로...
Private Sub mnuSTD040_Click()
    Load STD040
    STD040.Show
    STD040.ZOrder 0
    
End Sub





'>> 반정보 등록
Private Sub mnuLSN001_Click()
    Load LSN001
    LSN001.Show
    LSN001.ZOrder 0
    
End Sub

'>> 반 구성하기
Private Sub mnuLSN100_Click()
    Load LSN100
    LSN100.Show
    LSN100.ZOrder 0
    
End Sub



'>> 반 등록
Private Sub mnuLSN001_CP_Click()
    Load LSN001
    LSN001.Show
    LSN001.ZOrder 0
    
End Sub


'>> 이동반 정보 등록
Private Sub mnuLSN002_Click()
    Load LSN002
    LSN002.Show
    LSN002.ZOrder 0
End Sub

'>> 반 구성하기
Private Sub mnuLSN100_CP_Click()
    Load LSN100
    LSN100.Show
    LSN100.ZOrder 0
End Sub



'>> 구조별 시간표 코드 등록
Private Sub mnuMTX011_Click()
    Load MTX011
    MTX011.Show
    MTX011.ZOrder 0
    
End Sub





'>> 강사 및 강사별 과목넣기
Private Sub mnuTMR011_Click()
    Load TMR011
    TMR011.Show
    TMR011.ZOrder 0
    
End Sub

'>> 강사 시수넣기
Private Sub mnuTMR012_Click()
    Load TMR012
    TMR012.Show
    TMR012.ZOrder 0
    
End Sub

'>> 강사 강의불가 시간등록
Private Sub mnuTMR015_Click()
    Load TMR015
    TMR015.Show
    TMR015.ZOrder 0
    
End Sub



'>> 이동수업시간표 등록
Private Sub mnuTMR025_Click()
    Load TMR026
    TMR026.Show
    TMR026.ZOrder 0
    
End Sub


Private Sub mnuTMR027_Click()
    Load TMR028
    TMR028.Show
    TMR028.ZOrder 0
    
End Sub

'## 전체 시간표 구성 : 반별
Private Sub mnuTMR050_Click()
    Load TMR051
    TMR051.Show
    TMR051.ZOrder 0
    
End Sub


'>> 반 정보 등록
Private Sub mnuLSN001_CP2_Click()
    Load LSN001
    LSN001.Show
    LSN001.ZOrder 0

End Sub


'>> 반별 시간표 출력
Private Sub mnuPRT010_Click()
    Load PRT010
    PRT010.Show
    PRT010.ZOrder 0
    
End Sub

'>> 강사별 시간표 출력
Private Sub mnuPRT020_Click()
    Load PRT020
    PRT020.Show
    PRT020.ZOrder 0
    
End Sub


'>> 강사 및 반별 시간표 배정내역 조회
Private Sub mnuPRT030_Click()
    Load PRT031
    PRT031.Show
    PRT031.ZOrder 0

End Sub


'## 노량진 출력양식 변경요청 : 2008.02.19
Private Sub mnuPRT011_Click()
'    Load PRT011
'    PRT011.Show
'    PRT011.ZOrder 0
    
    '>> 2008.02.25 변경
    Load PRT012
    PRT012.Show
    PRT012.ZOrder 0
    
End Sub

Private Sub mnuPRT021_Click()
'    Load PRT021
'    PRT021.Show
'    PRT021.ZOrder 0
    
    '>> 2008.02.25 변경
    Load PRT022
    PRT022.Show
    PRT022.ZOrder 0
    
End Sub

Private Sub mnuPRT050_Click()
'    Load PRT050
'    PRT050.Show
'    PRT050.ZOrder 0
    
    Load PRT051
    PRT051.Show
    PRT051.ZOrder 0
    
End Sub

'## 강사 출석부
Private Sub mnuTMR060_Click()
    Load TMR060
    TMR060.Show
    TMR060.ZOrder 0
    
End Sub

'>> 등록금 부분 테스트
Private Sub TEST_Click()
    Load TMR028
    TMR028.Show
    TMR028.ZOrder 0
    
End Sub




'>> 윈터 & 수학클리닉 등록 및 조회
Private Sub mnuSTD200_Click()
    Load STD200
    STD200.Show
    STD200.ZOrder 0
End Sub

'>> 학생성적등록
Private Sub EXM010_Click()
    Load EXM100
    EXM100.Show
    EXM100.ZOrder 0
    
End Sub

'>> 학생관리
Private Sub EXM020_Click()
    Load EXM110
    EXM110.Show
    EXM110.ZOrder 0
    
End Sub
