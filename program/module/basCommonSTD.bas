Attribute VB_Name = "basCommonSTD"
    Option Explicit
    
    
    Dim g_


Function Set_Spread_Design1(ByRef sprControl As Object)

    With sprControl
        .ShadowColor = basModule.ShadowColor1
        .ShadowDark = basModule.ShadowDark1
        .ShadowText = basModule.ShadowText1
        .GridColor = basModule.GridColor1
        .GrayAreaBackColor = basModule.GrayAreaBackColor1
    End With
    
End Function

''노량진 계열 콤보박스 세팅.
'Function Init_Kaeyol_N(ByRef cboControl As Object)
'    With cboControl
'        .Clear
'        .AddItem "인문" & Space(30) & "01"
'        .AddItem "자연" & Space(30) & "02"
'
'
'    '<< 계열 >> : 2008.01.09
'        If Trim(basModule.SchCD) = "N" Then             '< 노량진
'
'            .AddItem "서울대인문" & Space(30) & "21"
'            .AddItem "서울대자연" & Space(30) & "22"
'            .AddItem "예체" & Space(30) & "03"
'            .AddItem "수리(나)" & Space(30) & "04"
'            .AddItem "인문수능" & Space(30) & "05"
'            .AddItem "자연수능" & Space(30) & "06"
'
'            .AddItem "인문-신" & Space(30) & "07"
'            .AddItem "자연-신" & Space(30) & "08"
'            '.AddItem "수능인문-신" & Space(30) & "09"
'            '.AddItem "수능자연-신" & Space(30) & "10"
'
'            .AddItem "편)인문" & Space(30) & "11"
'            .AddItem "편)자연" & Space(30) & "12"
'            .AddItem "편)예체" & Space(30) & "13"
'            .AddItem "편)수리(나)" & Space(30) & "14"
'            .AddItem "편)인문수능" & Space(30) & "15"
'            .AddItem "편)자연수능" & Space(30) & "16"
'
'        End If
'    '<< 계열 >> : 2008.01.10
'        'If Trim(basModule.SchCD) = "K" Then             '< 강남
'        Select Case Trim(basModule.SchCD)
'            Case "K", "W", "Q"
'                .AddItem "주말법대" & Space(30) & "04"
'                .AddItem "주말의대" & Space(30) & "05"
'
'                .AddItem "야간법대" & Space(30) & "06"
'                .AddItem "야간의대" & Space(30) & "07"
'
'                .AddItem "선착순인문" & Space(30) & "11"
'                .AddItem "선착순자연" & Space(30) & "12"
'
'                .AddItem "선착순인문16" & Space(30) & "16"
'                .AddItem "선착순자연17" & Space(30) & "17"
'
'        End Select
'
'    '<< 계열 >> : 2008.02.15
'        Select Case Trim(basModule.SchCD)               '< 송파
'            Case "S"
''                .AddItem "예체능" & Space(30) & "03"
''
''                .AddItem "인문수능" & Space(30) & "05"
''                .AddItem "자연수능" & Space(30) & "06"
''
'                .AddItem "신설인문" & Space(30) & "11"
'                .AddItem "신설자연" & Space(30) & "12"
'
''                .AddItem "인문프리미엄" & Space(30) & "18"
''                .AddItem "자연프리미엄" & Space(30) & "19"
'
'                .AddItem "서울대특별인문" & Space(30) & "21"
'                .AddItem "서울대특별자연" & Space(30) & "22"
'
'                .AddItem "야간서울대인문" & Space(30) & "21"
'                .AddItem "야간서울대자연" & Space(30) & "22"
'
'        End Select
'
'        Select Case Trim(basModule.SchCD)               '< 양재
'            Case "J"
'                .AddItem "신설인문" & Space(30) & "11"
'                .AddItem "신설자연" & Space(30) & "12"
'                .AddItem "인문프리미엄" & Space(30) & "18"
'                .AddItem "자연프리미엄" & Space(30) & "19"
'                .AddItem "서울대특별인문" & Space(30) & "21"
'                .AddItem "서울대특별자연" & Space(30) & "22"
'        End Select
'
'    '<< 계열 >> : 2009.01.09
'        If Trim(basModule.SchCD) = "B" Then             '< 부산
'
'            .AddItem "인문PS반" & Space(30) & "23"
'            .AddItem "자연PM반" & Space(30) & "24"
'
'            .AddItem "수학선행인문" & Space(30) & "05"
'            .AddItem "수학선행자연" & Space(30) & "06"
'
'            .AddItem "연.고대인문" & Space(30) & "07"
'            .AddItem "연.고대자연" & Space(30) & "08"
'
'            .AddItem "심화인문" & Space(30) & "09"
'            .AddItem "심화자연" & Space(30) & "10"
'        End If
'
'        Select Case Trim(basModule.SchCD)               '< 마강
'            Case "M"
'                .AddItem "서울대특별인문" & Space(30) & "21"
'                .AddItem "서울대특별자연" & Space(30) & "22"
'        End Select
'
'        .ListIndex = 0
'    End With
'End Function


Function Init_CboKaeyolDefault(ByRef cboControl As Object)
    With cboControl
        .Clear
        .AddItem "인문" & Space(30) & "01"
        .AddItem "자연" & Space(30) & "02"
        
        
    '<< 계열 >> : 2008.01.09
        If Trim(basModule.SchCD) = "N" Then             '< 노량진
        
            .AddItem "서울대인문" & Space(30) & "21"
            .AddItem "서울대자연" & Space(30) & "22"
            .AddItem "예체" & Space(30) & "03"
            .AddItem "수리(나)" & Space(30) & "04"
            .AddItem "인문수능" & Space(30) & "05"
            .AddItem "자연수능" & Space(30) & "06"
            
            .AddItem "인문-신" & Space(30) & "07"
            .AddItem "자연-신" & Space(30) & "08"
            '.AddItem "수능인문-신" & Space(30) & "09"
            '.AddItem "수능자연-신" & Space(30) & "10"
            
            .AddItem "편)인문" & Space(30) & "11"
            .AddItem "편)자연" & Space(30) & "12"
            .AddItem "편)예체" & Space(30) & "13"
            .AddItem "편)수리(나)" & Space(30) & "14"
            .AddItem "편)인문수능" & Space(30) & "15"
            .AddItem "편)자연수능" & Space(30) & "16"
            
            
        End If
    '<< 계열 >> : 2008.01.10
        'If Trim(basModule.SchCD) = "K" Then             '< 강남
        Select Case Trim(basModule.SchCD)
            Case "K", "W", "Q"
                .AddItem "주말법대" & Space(30) & "04"
                .AddItem "주말의대" & Space(30) & "05"
                
                .AddItem "야간법대" & Space(30) & "06"
                .AddItem "야간의대" & Space(30) & "07"
                
                .AddItem "선착순인문" & Space(30) & "11"
                .AddItem "선착순자연" & Space(30) & "12"
                
                .AddItem "선착순인문16" & Space(30) & "16"
                .AddItem "선착순자연17" & Space(30) & "17"
                
                .AddItem "내신우수자인문" & Space(30) & "19"
                .AddItem "내신우수자자연" & Space(30) & "20"
        End Select
    
        '<< 계열 >> : 2008.02.15
        Select Case Trim(basModule.SchCD)               '< 송파
            Case "S"
'                .AddItem "예체능" & Space(30) & "03"
'
'                .AddItem "인문수능" & Space(30) & "05"
'                .AddItem "자연수능" & Space(30) & "06"
'
                .AddItem "신설인문" & Space(30) & "11"
                .AddItem "신설자연" & Space(30) & "12"
                
'                .AddItem "인문프리미엄" & Space(30) & "18"
'                .AddItem "자연프리미엄" & Space(30) & "19"

                .AddItem "서울대특별인문" & Space(30) & "21"
                .AddItem "서울대특별자연" & Space(30) & "22"
                
                .AddItem "야간서울대인문" & Space(30) & "21"
                .AddItem "야간서울대자연" & Space(30) & "22"
                
        End Select
        
        
        Select Case Trim(basModule.SchCD)               '< 양재
            Case "J"
                .AddItem "신설인문" & Space(30) & "11"
                .AddItem "신설자연" & Space(30) & "12"
                .AddItem "인문프리미엄" & Space(30) & "18"
                .AddItem "자연프리미엄" & Space(30) & "19"
                .AddItem "서울대특별인문" & Space(30) & "21"
                .AddItem "서울대특별자연" & Space(30) & "22"
        End Select
        
    '<< 계열 >> : 2009.01.09
        If Trim(basModule.SchCD) = "B" Then             '< 부산
            
            .AddItem "인문PS반" & Space(30) & "23"
            .AddItem "자연PM반" & Space(30) & "24"
            
            .AddItem "선행인문" & Space(30) & "05"
            .AddItem "선행자연" & Space(30) & "06"
            
            .AddItem "연.고대인문" & Space(30) & "07"
            .AddItem "연.고대자연" & Space(30) & "08"
            
            .AddItem "심화인문" & Space(30) & "09"
            .AddItem "심화자연" & Space(30) & "10"
        End If
        
        Select Case Trim(basModule.SchCD)               '< 마강
            Case "M"
                .AddItem "서울대특별인문" & Space(30) & "21"
                .AddItem "서울대특별자연" & Space(30) & "22"
        End Select
    
        .ListIndex = 0
    End With
End Function


Function Set_CboKaeyol(ByRef cboControl As Object, ByVal SchCD As String, ByVal kaeyol As String) As Boolean
    
    If IsNull(SchCD) = True Or SchCD = "" Then
        Set_CboKaeyol = False
        Exit Function
    End If
    
    If IsNull(kaeyol) = True Or kaeyol = "" Then
        Set_CboKaeyol = False
        Exit Function
    End If
    
    
    If Trim(SchCD) = "N" Then
        If IsNull(kaeyol) = False Then
            Select Case Trim(kaeyol)
                Case "01"
                    cboControl.ListIndex = 0
                Case "02"
                    cboControl.ListIndex = 1
                Case "03"
                    cboControl.ListIndex = 4
                Case "04"
                    cboControl.ListIndex = 5
                Case "05"
                    cboControl.ListIndex = 6
                Case "06"
                    cboControl.ListIndex = 7
                    
                    
                Case "07"
                    cboControl.ListIndex = 8
                Case "08"
                    cboControl.ListIndex = 9
                'Case "09"
                '    cboControl.ListIndex = 8
                'Case "10"
                '    cboControl.ListIndex = 9
                    
                Case "11"
                    cboControl.ListIndex = 10
                Case "12"
                    cboControl.ListIndex = 11
                Case "13"
                    cboControl.ListIndex = 12
                Case "14"
                    cboControl.ListIndex = 13
                Case "15"
                    cboControl.ListIndex = 14
                Case "16"
                    cboControl.ListIndex = 15
                Case "21"
                    cboControl.ListIndex = 2
                Case "22"
                    cboControl.ListIndex = 3
            End Select
        End If
        
    ElseIf Trim(SchCD) = "B" Then
        If IsNull(kaeyol) = False Then
            Select Case Trim(kaeyol)
                Case "01"
                    cboControl.ListIndex = 0
                Case "02"
                    cboControl.ListIndex = 1
                Case "05"
                    cboControl.ListIndex = 4
                Case "06"
                    cboControl.ListIndex = 5
                    
                Case "07"
                    cboControl.ListIndex = 6
                Case "08"
                    cboControl.ListIndex = 7
                Case "09"
                    cboControl.ListIndex = 8
                Case "10"
                    cboControl.ListIndex = 9
                Case "23"
                    cboControl.ListIndex = 2
                Case "24"
                    cboControl.ListIndex = 3
            End Select
        End If
        
    ElseIf (Trim(SchCD) = "K") Or (Trim(SchCD) = "W") Or (Trim(SchCD) = "Q") Then
        If IsNull(kaeyol) = False Then
            Select Case Trim(kaeyol)
                Case "01"
                    cboControl.ListIndex = 0
                Case "02"
                    cboControl.ListIndex = 1
                    
                Case "04"
                    cboControl.ListIndex = 2
                Case "05"
                    cboControl.ListIndex = 3
                Case "06"
                    cboControl.ListIndex = 4
                Case "07"
                    cboControl.ListIndex = 5
                    
                Case "11"
                    cboControl.ListIndex = 6
                Case "12"
                    cboControl.ListIndex = 7
                    
                Case "16"
                    cboControl.ListIndex = 8
                Case "17"
                    cboControl.ListIndex = 9
                    
                Case "19"
                    cboControl.ListIndex = 10
                Case "20"
                    cboControl.ListIndex = 11
                    
            End Select
        End If
        
    ElseIf Trim(SchCD) = "S" Then
        If IsNull(kaeyol) = False Then
            Select Case Trim(kaeyol)
                Case "01"
                    cboControl.ListIndex = 0
                Case "02"
                    cboControl.ListIndex = 1
                    
        ''                        Case "03"
        ''                            cboControl.ListIndex = 2
        ''
        ''                        Case "05"
        ''                            cboControl.ListIndex = 3
        ''                        Case "06"
        ''                            cboControl.ListIndex = 4
        ''
        ''                        Case "11"
        ''                            cboControl.ListIndex = 5
        ''                        Case "12"
        ''                            cboControl.ListIndex = 6
        '
        '                        Case "18"
        '                            cboControl.ListIndex = 7
        '                        Case "19"
        '                            cboControl.ListIndex = 8
                    
        '                        Case "18"
        '                            cboControl.ListIndex = 2
        '                        Case "19"
        '                            cboControl.ListIndex = 3
                Case "21"
                    cboControl.ListIndex = 2
                Case "22"
                    cboControl.ListIndex = 3
                    
                    
            End Select
        End If
    ElseIf Trim(SchCD) = "P" Then                 '< 마송
        If IsNull(kaeyol) = False Then
            Select Case Trim(kaeyol)
                Case "01"
                    cboControl.ListIndex = 0
                Case "02"
                    cboControl.ListIndex = 1
                    
                Case "03"
                    cboControl.ListIndex = 2
                Case "04"
                    cboControl.ListIndex = 3
            End Select
        End If
        
    ElseIf Trim(SchCD) = "J" Then                 '< 양재
        If IsNull(kaeyol) = False Then
            Select Case Trim(kaeyol)
                Case "01"
                    cboControl.ListIndex = 0
                Case "02"
                    cboControl.ListIndex = 1
                    
                Case "11"
                    cboControl.ListIndex = 2
                Case "12"
                    cboControl.ListIndex = 3
                    
                Case "18"
                    cboControl.ListIndex = 4
                Case "19"
                    cboControl.ListIndex = 5
                    
            End Select
        End If
    ElseIf Trim(SchCD) = "M" Then                 '< 마강
        If IsNull(kaeyol) = False Then
            Select Case Trim(kaeyol)
                Case "01"
                    cboControl.ListIndex = 0
                Case "02"
                    cboControl.ListIndex = 1
                    
                Case "21"
                    cboControl.ListIndex = 2
                Case "22"
                    cboControl.ListIndex = 3
            End Select
        End If
        
    Else
        If IsNull(kaeyol) = False Then
            Select Case Trim(kaeyol)
                Case "01"
                    cboControl.ListIndex = 0
                Case "02"
                    cboControl.ListIndex = 1
                Case "03"
                    cboControl.ListIndex = 2
            End Select
        End If
    End If
End Function
'학원
Function Init_CboSch(ByRef cboControl As Object)
    With cboControl
        .Clear
        .AddItem "없음" & Space(30) & "X"
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
End Function

'학원
Function Set_CboSch(ByRef cboControl As Object, ByVal sSch As String)

    Select Case Trim(sSch)
        Case "N"
            cboControl.ListIndex = 1
        Case "K"
            cboControl.ListIndex = 2
        Case "S"
            cboControl.ListIndex = 3
        Case "P"
            cboControl.ListIndex = 4
        Case "M"
            cboControl.ListIndex = 5
        Case "W"
            cboControl.ListIndex = 6
        Case "Q"
            cboControl.ListIndex = 7
        Case "J"
            cboControl.ListIndex = 8
        Case "B"
            cboControl.ListIndex = 9
        Case Else
            cboControl.ListIndex = 0
    End Select
End Function


'합격
Function Init_PassCN(ByRef cboControl As Object)
    With cboControl
        .Clear
        .AddItem "전체" & Space(30) & "ALL"
        .AddItem "1차" & Space(30) & "1"
        .AddItem "2차" & Space(30) & "2"
        .AddItem "3차" & Space(30) & "3"
        .AddItem "4차" & Space(30) & "4"
        
        .ListIndex = 0
    End With
End Function

'결제
Function Init_Pay(ByRef cboControl As Object)
    With cboControl
        .Clear
        .AddItem "전체" & Space(30) & "ALL"
        .AddItem "결재" & Space(30) & "OK"
        .AddItem "미결재" & Space(30) & "NOT"
        
        .ListIndex = 0
    End With
End Function

'시험유형
Function Init_ExmType(ByRef cboControl As Object)
    With cboControl
        .Clear
        .AddItem "전체" & Space(30) & "ALL"
        .AddItem "유시험" & Space(30) & "1"
        .AddItem "무시험" & Space(30) & "0"
        
        .ListIndex = 0
    End With
End Function

'인터넷/학원
Function Init_InGbn(ByRef cboControl As Object)
    With cboControl
        .Clear
        .AddItem "전체" & Space(30) & "ALL"
        .AddItem "인터넷" & Space(30) & "INT"
        .AddItem "학원" & Space(30) & "HAK"
        
        .ListIndex = 0
    End With
End Function

'등급
Function Init_Mu_type(ByRef cboControl As Object)
    With cboControl
        .Clear
        
        .AddItem "수능등급" & Space(30) & "1"   '점수
        .AddItem "2013 6월 평가원" & Space(30) & "2"
        .AddItem "2013 9월 평가원" & Space(30) & "3"
        
        If basModule.SchCD = "N" Or basModule.SchCD = "S" _
            Or basModule.SchCD = "J" Or basModule.SchCD = "K" Or basModule.SchCD = "M" Then
            .AddItem "내신등급" & Space(30) & "9"
        End If
        
        .AddItem "없음" & Space(30) & "X"
        
        .Enabled = True
        .ListIndex = .ListCount - 1
        
    End With
End Function

'수리점수 구분
Function Init_PTS_Sel(ByRef cboControl As Object)
    With cboControl
        .Clear
        .AddItem "1"
        .AddItem "2"
        .AddItem ""
        
        .Enabled = False
        .ListIndex = 2
    End With
End Function

'카드
Function Init_Card(ByRef cboControl As Object)
    With cboControl
        Select Case Trim(basModule.SchCD)
            Case "N", "K", "W", "Q", "S"
                .AddItem "아멕스카드               AMX"
                .AddItem "전북은행카드             CBB"
                .AddItem "다이너스카드             DIN"
                .AddItem "한미은행카드             KAB"
                .AddItem "강원카드                 KWB"
                .AddItem "축협카드                 NLC"
                .AddItem "신세계카드               SIN"
                .AddItem "BC카드                   BCC"
                .AddItem "제주은행카드             CJB"
                .AddItem "하나은행카드             HNB"
                .AddItem "외환은행카드             KEB"
                .AddItem "LG카드                   LGC"
                .AddItem "평화은행카드             PHB"
                .AddItem "삼성카드                 WIN"
                .AddItem "외국은행카드             BRD"
                .AddItem "국민은행카드             CNB"
                .AddItem "JCB카드                  JCB"
                .AddItem "광주은행카드             KJB"
                .AddItem "수협카드                 NFF"
                .AddItem "신한은행카드             SHB"
                
            Case "M", "P", "J", "B"
    
                '20121221
                .AddItem "KB국민카드        CCKM"
                .AddItem "NH채움카드        CCNH"
                .AddItem "신세계한미        CCSG"
                .AddItem "씨티카드          CCCT"
                .AddItem "한미카드          CCHM"
                .AddItem "해외비자          CVSF"
                .AddItem "국내아멕스        CCAM"
                .AddItem "롯데카드          CCLO"
                .AddItem "해외아멕스        CAMF"
                .AddItem "BC카드            CCBC"
                .AddItem "우리카드          CCPH"
                .AddItem "하나SK카드        CCHN"
                .AddItem "삼성카드          CCSS"
                .AddItem "광주카드          CCKJ"
                .AddItem "수협카드          CCSU"
                .AddItem "신협카드          CCCU"
                .AddItem "신한카드          CCSH"
                .AddItem "전북카드          CCJB"
                .AddItem "제주카드          CCCJ"
                .AddItem "신한카드          CCLG"
                .AddItem "해외마스터        CMCF"
                .AddItem "해외JCB           CJCF"
                .AddItem "외환카드          CCKE"
                .AddItem "현대카드          CCDI"
                .AddItem "저축카드          CCSB"
                .AddItem "산은카드          CCKD"
                .AddItem "은련카드          CCUF"
'                .AddItem "BC카드                      CCBC"
'                .AddItem "국민카드                    CCKM"
'                .AddItem "LG카드                      CCLG"
'                .AddItem "삼성카드                    CCSS"
'                .AddItem "외환카드                    CCKE"
'                .AddItem "신한카드                    CCSH"
'                .AddItem "수협카드                    CCSU"
'                .AddItem "광주은행                    CCKJ"
'                .AddItem "강원은행                    CCKW"
'                .AddItem "하나은행                    CCHN"
'                .AddItem "국내아멕스                  CCAM"
'                .AddItem "해외아멕스                  CAMF"
'                .AddItem "한미은행                    CCYJ"
'                .AddItem "축협카드                    CCCH"
'                .AddItem "평화은행                    CCPH"
'                .AddItem "제주은행                    CCCJ"
'                .AddItem "전북은행                    CCJB"
'                .AddItem "현대카드                    CCDI"
'                .AddItem "시티은행                    CCCT"
'                .AddItem "동남은행                    CCDN"
'                .AddItem "해외비자                    CVSF"
'                .AddItem "해외마스타카드              CMCF"
'                .AddItem "해외JCB카드                 CJCF"
'                .AddItem "롯데카드                    CCLO"
                
        End Select
        .ListIndex = 0
    End With

End Function

'클리닉 콤보 초기화
Sub Init_Clinic(ByRef chkClinic_L As Object, ByRef chkClinic_M As Object, ByRef chkClinic_E As Object)

    If SchCD = "N" Or SchCD = "S" Then
        Dim i As Integer
    
        For i = 0 To chkClinic_L.count - 1
            chkClinic_L(i).value = False
        Next i
        
        For i = 0 To chkClinic_M.count - 1
            chkClinic_M(i).value = False
        Next i
        
        For i = 0 To chkClinic_E.count - 1
            chkClinic_E(i).value = False
        Next i
        
    Else
        
        
    End If

End Sub


'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'초기화 끝
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'데이터 가져오기/불러오기
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Function Get_SchName(sSch As String)

    If IsNull(sSch) = True Then
        Get_SchName = ""
        Exit Function
    End If
    
    Dim sTmp As String
    Select Case Trim(sSch)
        Case "N"
            sTmp = "노량진"
        Case "K"
            sTmp = "강남"
        Case "S"
            sTmp = "송파"
        Case "P"
            sTmp = "송파 M"
        Case "M"
            sTmp = "강남 M"
        Case "W"
            sTmp = "주말법의대"
        Case "Q"
            sTmp = "야간법의대"
        Case "J"
            sTmp = "양재"
        Case "B"
            sTmp = "부산"
        Case "E"
            sTmp = "강남기숙(이천)"
        Case Else
            sTmp = ""
    End Select
    
    Get_SchName = sTmp
End Function


'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'데이터 가져오기/불러오기
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

    
'클리닉 콤보 설정
Sub Set_Clinic(ByRef chkClinic_L As Object, ByRef chkClinic_M As Object, ByRef chkClinic_E As Object, ByVal SEL7 As String)

    Dim sDiv()      As String
    Dim ni          As Integer
    
    
    sDiv = Split(SEL7, "|", -1, vbTextCompare)
    
    For ni = 0 To UBound(sDiv) - 1 Step 1
        If sDiv(ni) <> "" Then
            If sDiv(ni) >= CLINIC_L_CLASS And sDiv(ni) < CLINIC_L_CLASS + 10 Then
            
                chkClinic_L(CInt(sDiv(ni)) - CLINIC_L_CLASS).value = True
                
            ElseIf sDiv(ni) >= CLINIC_M_CLASS And sDiv(ni) < CLINIC_M_CLASS + 10 Then
            
                chkClinic_M(CInt(sDiv(ni)) - CLINIC_M_CLASS).value = True
                
            ElseIf sDiv(ni) >= CLINIC_E_CLASS And sDiv(ni) < CLINIC_E_CLASS + 10 Then
            
                chkClinic_E(CInt(sDiv(ni)) - CLINIC_E_CLASS).value = True
                
            End If
        End If
    Next ni
    
End Sub

'등급
Public Sub Set_Mu_type(ByRef cboControl As Object, ByVal val As Integer)
    Select Case val
        Case "1"
            cboControl.ListIndex = 0 '수능등급
        Case "2"
            cboControl.ListIndex = 1 '6월 모평
        Case "3"
            cboControl.ListIndex = 2 '9월 모평
        Case "9"
            cboControl.ListIndex = 3 '내신등급
    End Select
    
End Sub


'등급의 값 가져오기
Public Function Get_StrMuType(ByVal value)
     Select Case value
        Case "1"
            Get_StrMuType = "수능등급"
        Case "2"
            Get_StrMuType = "6월 모평"
        Case "3"
            Get_StrMuType = "9월 모평"
        Case "9"
            Get_StrMuType = "내신등급"
    End Select
End Function

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'학원별 공지사항
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Public Function Get_StrGongji() As String()
    Dim strReturn() As String
    
    '>> 학년별 내역
    
    Select Case Trim(basModule.SchCD)
    
        Case "N"
            ReDim strReturn(3)
            strReturn(0) = "● 국어, 수학, 영어, 과목 중 (심화)또는(기본) 수업 1과목을 선택해야 하며, 탐구과목 4과목 중 1과목을 선택할 수 있습니다."
            strReturn(1) = "● 인문계는 생활과 윤리, 윤리와 사상, 세계지리, 동아시아사, 세계사, 경제, 제2외국어, 자연계는 과학Ⅱ(4과목)는 재수 정규반부터 수업합니다."
            strReturn(2) = "● 반당 수강생 수 증감에 따라 분반 또는 합반할 수 있습니다."
            strReturn(3) = "● 인문계는(국어B, 수학A, 영어B) / 자연계(국어A, 수학B, 영어B형)으로 수업합니다."
           
        Case "K", "W", "Q"
            ReDim strReturn(2)
            strReturn(0) = "▶인문계 사회탐구 중 법과 정치, 세계사, 세계지리, 동아시아사, 윤리와 사상, 생활과 윤리 및 제2외국어는 정규반에서 선택하여 수강할 수 있습니다."
            strReturn(1) = "▶자연계 과학탐구 중 물리Ⅱ, 화학Ⅱ, 생명과학Ⅱ, 지구과학Ⅱ는 정규반에서 선택하여 수강할 수 있습니다."
            strReturn(2) = "▶인문계(국어B, 수학A, 영어B형)/자연계(국어A, 수학B, 영어B형)으로 수업합니다."
            
        Case "S"
            ReDim strReturn(3)
            strReturn(0) = "● 국어, 수학, 영어, 과목 중 (심화)또는(기본) 수업 2과목을 선택해야 하며, 탐구과목중 1과목을 선택할 수 있습니다."
            strReturn(1) = "● 인문계는 생활과 윤리, 윤리와 사상, 세계지리, 동아시아사, 세계사, 경제, 제2외국어, 자연계는 과학Ⅱ(4과목)는 재수 정규반부터 수업합니다."
            strReturn(2) = "● 반당 수강생 수 증감에 따라 분반 또는 합반할 수 있습니다."
            strReturn(3) = "● 인문계는(국어B, 수학A, 영어B) / 자연계(국어A, 수학B, 영어B형)으로 수업합니다."
                  
        Case "P"
            ReDim strReturn(1)
            strReturn(0) = ""
            strReturn(1) = ""
                             
         Case "M"
            ReDim strReturn(2)
            strReturn(0) = "▶인문계 사회탐구 중 세계사, 세계지리, 동아시아사, 생활과 윤리 및 제2외국어는 정규반에서 선택하여 수강할 수 있습니다."
            strReturn(1) = "▶인문계(국어B, 수학A, 영어B형)/자연계(국어A, 수학B, 영어B형)으로 수업합니다."
            strReturn(2) = ""
            
'
        Case "J"        '> 양재
            ReDim strReturn(3)
            strReturn(0) = "▶인문계(국어B, 수학A, 영어B형)/자연계(국어A, 수학B, 영어B형)으로 수업합니다."
            strReturn(1) = "▶인문계 사회탐구 중 세계사, 세계지리, 동아시아사, 생활과 윤리 및 제2외국어는 정규반에서 선택하여 수강할 수 있습니다."
            strReturn(2) = "▶자연계 과학탐구 중 물리Ⅱ, 지구과학Ⅱ는 정규반에서 선택하여 수강할 수 있습니다."
            strReturn(3) = "▶선택과목의 신청자 수가 극소수일 경우 개설되지 않을 수도 있습니다. "
            
        Case "B"        '> 부산
            ReDim strReturn(2)
            strReturn(0) = "▶ 선행학습반 인문계는 6과목 중 3과목을 선택할 수 있으며, 경제, 세계지리, 세계사, 법과사회, 경제지리, 제2외국어 과목은 정규 종합반부터 수업합니다."
            strReturn(1) = "▶ 선행학습반 자연계는 4과목 중 3과목을 선택할 수 있으며, 과학II(4과목), 수리영역 선택과목 적분, 확률통계는 정규 종합반부터 수업합니다."
            strReturn(2) = ""
            
    End Select
    
    Get_StrGongji = strReturn
    
End Function

Public Function Get_StrGongjiJonghab() As String()
    Dim strReturn() As String
    
    '>> 학년별 내역
    
    Select Case Trim(basModule.SchCD)
    
        Case "N"
            ReDim strReturn(3)
            strReturn(0) = "● 국어, 수학, 영어, 과목 중 (심화)또는(기본) 수업 1과목을 선택해야 하며, 탐구과목 4과목 중 1과목을 선택할 수 있습니다."
            strReturn(1) = "● 인문계는 생활과 윤리, 윤리와 사상, 세계지리, 동아시아사, 세계사, 경제, 제2외국어, 자연계는 과학Ⅱ(4과목)는 재수 정규반부터 수업합니다."
            strReturn(2) = "● 반당 수강생 수 증감에 따라 분반 또는 합반할 수 있습니다."
            strReturn(3) = "● 인문계는(국어B, 수학A, 영어B) / 자연계(국어A, 수학B, 영어B형)으로 수업합니다."
           
        Case "K", "W", "Q"
            ReDim strReturn(2)
            strReturn(0) = "▶인문계 사회탐구 중 법과 정치, 세계사, 세계지리, 동아시아사, 윤리와 사상, 생활과 윤리 및 제2외국어는 정규반에서 선택하여 수강할 수 있습니다."
            strReturn(1) = "▶자연계 과학탐구 중 물리Ⅱ, 화학Ⅱ, 생명과학Ⅱ, 지구과학Ⅱ는 정규반에서 선택하여 수강할 수 있습니다."
            strReturn(2) = "▶인문계(국어B, 수학A, 영어B형)/자연계(국어A, 수학B, 영어B형)으로 수업합니다."
            
            
        Case "S"
            ReDim strReturn(3)
            strReturn(0) = "● 국어, 수학, 영어, 과목 중 (심화)또는(기본) 수업 2과목을 선택해야 하며, 탐구과목중 1과목을 선택할 수 있습니다."
            strReturn(1) = "● 인문계는 생활과 윤리, 윤리와 사상, 세계지리, 동아시아사, 세계사, 경제, 제2외국어, 자연계는 과학Ⅱ(4과목)는 재수 정규반부터 수업합니다."
            strReturn(2) = "● 반당 수강생 수 증감에 따라 분반 또는 합반할 수 있습니다."
            strReturn(3) = "● 인문계는(국어B, 수학A, 영어B) / 자연계(국어A, 수학B, 영어B형)으로 수업합니다."
                  
        Case "P"
            ReDim strReturn(1)
            strReturn(0) = ""
            strReturn(1) = ""
                             
         Case "M"
            ReDim strReturn(2)
            strReturn(0) = "▶인문계 사회탐구 중 세계사, 세계지리, 동아시아사, 생활과 윤리 및 제2외국어는 정규반에서 선택하여 수강할 수 있습니다."
            strReturn(1) = "▶인문계(국어B, 수학A, 영어B형)/자연계(국어A, 수학B, 영어B형)으로 수업합니다."
            strReturn(2) = ""
            
'
        Case "J"        '> 양재
            ReDim strReturn(3)
            strReturn(0) = "▶인문계(국어B, 수학A, 영어B형)/자연계(국어A, 수학B, 영어B형)으로 수업합니다."
            strReturn(1) = "▶인문계 사회탐구 중 세계사, 세계지리, 동아시아사, 생활과 윤리 및 제2외국어는 정규반에서 선택하여 수강할 수 있습니다."
            strReturn(2) = "▶자연계 과학탐구 중 물리Ⅱ, 지구과학Ⅱ는 정규반에서 선택하여 수강할 수 있습니다."
            strReturn(3) = "▶선택과목의 신청자 수가 극소수일 경우 개설되지 않을 수도 있습니다. "
            
        Case "B"        '> 부산
            ReDim strReturn(2)
            strReturn(0) = "▶ 선행학습반 인문계는 6과목 중 3과목을 선택할 수 있으며, 경제, 세계지리, 세계사, 법과사회, 경제지리, 제2외국어 과목은 정규 종합반부터 수업합니다."
            strReturn(1) = "▶ 선행학습반 자연계는 4과목 중 3과목을 선택할 수 있으며, 과학II(4과목), 수리영역 선택과목 적분, 확률통계는 정규 종합반부터 수업합니다."
            strReturn(2) = ""
            
    End Select
    
    Get_StrGongjiJonghab = strReturn
    
End Function

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'엑셀 저장 SQL문
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Function Get_SqlKaeyolDecode()
    Dim sStr    As String
    
    sStr = ""
    If Trim(basModule.SchCD) = "N" Then
        sStr = sStr & "     DECODE(KAEYOL,'01','인문',"
        sStr = sStr & "                   '02','자연',"
        sStr = sStr & "                   '03','예체',"
        sStr = sStr & "                   '04','수리(나)',"
        sStr = sStr & "                   '05','인문수능',"
        sStr = sStr & "                   '06','자연수능',"
        
        sStr = sStr & "                   '06','자연수능',"
        sStr = sStr & "                   '07','신설인문',"
        sStr = sStr & "                   '08','신설자연',"
        sStr = sStr & "                   '09','신설수능인문',"
        sStr = sStr & "                   '10','신설수능자연',"
        
        sStr = sStr & "                   '11','편)인문',"
        sStr = sStr & "                   '12','편)자연',"
        sStr = sStr & "                   '13','편)예체',"
        sStr = sStr & "                   '14','편)수리(나)',"
        sStr = sStr & "                   '15','편)인문수능',"
        sStr = sStr & "                   '16','편)자연수능',"
        sStr = sStr & "                   '21','서울대인문',"
        sStr = sStr & "                   '22','서울대인문'"
        sStr = sStr & "            ) AS GAEYUL,"
        
    '<< 계열 >> : 2008.01.10/ 2008.03.24
    ElseIf Trim(basModule.SchCD) = "K" Or Trim(basModule.SchCD) = "W" Or Trim(basModule.SchCD) = "Q" Then
        sStr = sStr & "     DECODE(KAEYOL,'01','인문',"
        sStr = sStr & "                   '02','자연',"
        
        sStr = sStr & "                   '04','주말법대',"
        sStr = sStr & "                   '05','주말의대',"
        sStr = sStr & "                   '06','야간법대',"
        sStr = sStr & "                   '07','야간의대',"
        
        sStr = sStr & "                   '11','선착순인문',"
        sStr = sStr & "                   '12','선착순자연',"
        
        sStr = sStr & "                   '16','선착순인문16',"
        sStr = sStr & "                   '17','선착순자연17',"
        
        sStr = sStr & "                   '19','내신우수자인문',"
        sStr = sStr & "                   '20','내신우수자자연'"
        
        sStr = sStr & "            ) AS GAEYUL,"
        
    '<< 계열 >> : 2008.02.15
    ElseIf Trim(basModule.SchCD) = "S" Then
       sStr = sStr & "     DECODE(KAEYOL,'01','인문',"
        sStr = sStr & "                   '02','자연',"
        
        sStr = sStr & "                   '03','예체능',"
        
        sStr = sStr & "                   '05','수능인문',"
        sStr = sStr & "                   '06','수능자연',"
        
        sStr = sStr & "                   '11','신설인문',"
        sStr = sStr & "                   '12','신설자연',"
        
        sStr = sStr & "                   '18','인문프리미엄',"
        sStr = sStr & "                   '19','자연프리미엄',"
        sStr = sStr & "                   '21','서울대특별인문',"
        sStr = sStr & "                   '22','서울대특별자연',"
        sStr = sStr & "                   '23','야간서울대인문',"
        sStr = sStr & "                   '24','야간서울대자연'"
        
        sStr = sStr & "            ) AS GAEYUL,"
    ElseIf Trim(basModule.SchCD) = "J" Then                 '< 양재
        sStr = sStr & "     DECODE(KAEYOL,'01','인문',"
        sStr = sStr & "                   '02','자연',"
        sStr = sStr & "                   '11','신설인문',"
        sStr = sStr & "                   '12','신설자연',"
        
        sStr = sStr & "                   '18','인문프리미엄',"
        sStr = sStr & "                   '19','자연프리미엄'"
        
        sStr = sStr & "            ) AS GAEYUL,"
        
    ElseIf Trim(basModule.SchCD) = "P" Then                 '< 마송
        sStr = sStr & "     DECODE(KAEYOL,'01','인문',"
        sStr = sStr & "                   '02','자연',"
        sStr = sStr & "                   '03','특별인문',"
        sStr = sStr & "                   '04','특별자연'"
        sStr = sStr & "            ) AS GAEYUL,"
        
    ElseIf Trim(basModule.SchCD) = "B" Then                 '< 부산 : 2009.01.09
        sStr = sStr & "     DECODE(KAEYOL,'01','인문',"
        sStr = sStr & "                   '02','자연',"
        sStr = sStr & "                   '23','인문PS',"
        sStr = sStr & "                   '24','자연PM',"
        sStr = sStr & "                   '05','특별인문',"
        sStr = sStr & "                   '06','특별자연',"
        sStr = sStr & "                   '07','연고대인문',"
        sStr = sStr & "                   '08','연고대자연',"
        sStr = sStr & "                   '09','심화인문',"
        sStr = sStr & "                   '10','심화자연'"
        sStr = sStr & "            ) AS GAEYUL,"
        
    Else
        sStr = sStr & "     DECODE(KAEYOL,'01','인문',"
        sStr = sStr & "                   '02','자연'"
        sStr = sStr & "            ) AS GAEYUL,"
    End If
    
    Get_SqlKaeyolDecode = sStr
End Function

Public Function AddSQL_ClinicToExcel()
    Dim sStr    As String
    
    sStr = ""
    sStr = sStr & " CASE WHEN SEL7 > ' ' AND    INSTR (SEL7, '101') > 0 THEN '" & g_sClinic_Ls(0) & "'"
    sStr = sStr & " ELSE CASE WHEN SEL7 > ' ' AND    INSTR (SEL7, '102') > 0 THEN '" & g_sClinic_Ls(1) & "'"
    sStr = sStr & " ELSE CASE WHEN SEL7 > ' ' AND    INSTR (SEL7, '103') > 0 THEN '" & g_sClinic_Ls(2) & "'"
    sStr = sStr & " ELSE CASE WHEN SEL7 > ' ' AND    INSTR (SEL7, '104') > 0 THEN '" & g_sClinic_Ls(3) & "'"
    sStr = sStr & " END END END END 국어클리닉,"
    sStr = sStr & " CASE WHEN SEL7 > ' ' AND    INSTR (SEL7, '111') > 0 THEN '" & g_sClinic_Ms(0) & "'"
    sStr = sStr & " ELSE CASE WHEN SEL7 > ' ' AND    INSTR (SEL7, '112') > 0 THEN '" & g_sClinic_Ms(1) & "'"
    sStr = sStr & " ELSE CASE WHEN SEL7 > ' ' AND    INSTR (SEL7, '113') > 0 THEN '" & g_sClinic_Ms(2) & "'"
    sStr = sStr & " ELSE CASE WHEN SEL7 > ' ' AND    INSTR (SEL7, '114') > 0 THEN '" & g_sClinic_Ms(3) & "'"
    sStr = sStr & " END END END END 수학클리닉,"
    sStr = sStr & " CASE WHEN SEL7 > ' ' AND    INSTR (SEL7, '121') > 0 THEN '" & g_sClinic_Es(0) & "'"
    sStr = sStr & " ELSE CASE WHEN SEL7 > ' ' AND    INSTR (SEL7, '122') > 0 THEN '" & g_sClinic_Es(1) & "'"
    sStr = sStr & " ELSE CASE WHEN SEL7 > ' ' AND    INSTR (SEL7, '123') > 0 THEN '" & g_sClinic_Es(2) & "'"
    sStr = sStr & " ELSE CASE WHEN SEL7 > ' ' AND    INSTR (SEL7, '124') > 0 THEN '" & g_sClinic_Es(3) & "'"
    sStr = sStr & " END END END END 영어클리닉"
    
    AddSQL_ClinicToExcel = sStr
End Function

'학생 엑셀 저장 쿼리문 (노량진,송파)
Public Function Get_StdExcuteSqlToExcel_N(kaeyol As String, Optional day1 As String, Optional day2 As String) As String
    Dim sStr        As String
    
    
    sStr = ""
    sStr = sStr & "  SELECT SCHNO AS 시스템코드   , "
    sStr = sStr & "         ACID  AS 학원   , "
    sStr = sStr & "         EXMID AS 수험번호, STDNM AS 학생, "
    
    sStr = sStr & " Birth_ymd as 생년월일, "
    
    sStr = sStr & "         DECODE(EXMTYPE,'0','무시험','1','유시험') AS 시험형태, "
    sStr = sStr & "         DECODE(KAEYOL,'01','인문',"
    sStr = sStr & "                       '02','자연',"
'<< 계열 >> : 2008.01.09
    If Trim(basModule.SchCD) = "N" Then
        sStr = sStr & "                   '03','예체',"
        sStr = sStr & "                   '04','수리(나)',"
        sStr = sStr & "                   '05','인문수능',"
        sStr = sStr & "                   '06','자연수능',"
        
        sStr = sStr & "                   '07','신설인문',"
        sStr = sStr & "                   '08','신설자연',"
        sStr = sStr & "                   '09','신설수능인문',"
        sStr = sStr & "                   '10','신설수능자연',"
        
        sStr = sStr & "                   '11','편)인문',"
        sStr = sStr & "                   '12','편)자연',"
        sStr = sStr & "                   '13','편)예체',"
        sStr = sStr & "                   '14','편)수리(나)',"
        sStr = sStr & "                   '15','편)인문수능',"
        sStr = sStr & "                   '16','편)자연수능',"
        sStr = sStr & "                   '21','서울대인문',"
        sStr = sStr & "                   '22','서울대자연',"
    End If
'<< 계열 >> : 2008.01.10
    If Trim(basModule.SchCD) = "K" Or Trim(basModule.SchCD) = "W" Or Trim(basModule.SchCD) = "Q" Then
        sStr = sStr & "                   '04','주말법대',"
        sStr = sStr & "                   '05','주말의대',"
        sStr = sStr & "                   '06','야간법대',"
        sStr = sStr & "                   '07','야간의대',"
    
        sStr = sStr & "                   '11','선착순인문',"
        sStr = sStr & "                   '12','선착순자연',"
        
        sStr = sStr & "                   '16','선착순인문16',"
        sStr = sStr & "                   '17','선착순자연17',"
        
        sStr = sStr & "                   '19','내신우수자인문',"
        sStr = sStr & "                   '20','내신우수자자연',"
        

    End If
'<< 계열 >> : 2008.02.15
    If Trim(basModule.SchCD) = "S" Then
        sStr = sStr & "                   '03','예체능',"
        'sStr = sStr & "                   '04','특별자연',"
        
        sStr = sStr & "                   '05','수능인문',"
        sStr = sStr & "                   '06','수능자연',"
        
        sStr = sStr & "                   '11','신설인문',"
        sStr = sStr & "                   '12','신설자연',"
        
        sStr = sStr & "                   '18','인문프리미엄',"
        sStr = sStr & "                   '19','자연프리미엄',"
        sStr = sStr & "                   '21','서울대특별인문',"
        sStr = sStr & "                   '22','서울대특별자연',"
        sStr = sStr & "                   '23','야간서울대인문',"
        sStr = sStr & "                   '24','야간서울대자연',"
        
    End If
'<< 계열 >> : 2008.02.15
    If Trim(basModule.SchCD) = "P" Then         '< 마송
        sStr = sStr & "                   '03','특별인문',"
        sStr = sStr & "                   '04','특별자연',"
    End If
    
    If Trim(basModule.SchCD) = "J" Then         '< 양재
        sStr = sStr & "                   '11','신설인문',"
        sStr = sStr & "                   '12','신설자연',"
        
        sStr = sStr & "                   '18','인문프리미엄',"
        sStr = sStr & "                   '19','자연프리미엄',"
    End If
    
'<< 계열 >> : 2009.01.09
    If Trim(basModule.SchCD) = "B" Then         '< 부산
        sStr = sStr & "                   '05','선행인문',"
        sStr = sStr & "                   '06','선행자연',"
        sStr = sStr & "                   '07','연고대인문',"
        sStr = sStr & "                   '08','연고대자연',"
        sStr = sStr & "                   '09','심화인문',"
        sStr = sStr & "                   '10','심화자연',"
    End If
    
    sStr = sStr & "                       '','기타') AS 계열,"
    
    sStr = sStr & "     /* 사탐, 과탐 분리 */"
    sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'" & constSatamCodes(0) & "|') > 0 THEN          /* 사탐-국사 */"
    sStr = sStr & "             '" & constSatams(0) & "'"
    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'51|') > 0 THEN     /* 과탐-물리1 */"
    sStr = sStr & "             '물1'"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             ' '"
    sStr = sStr & "         END END AS 탐구1,"
    sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'" & constSatamCodes(1) & "|') > 0 THEN          /* 사탐-윤리 */"
    sStr = sStr & "             '" & constSatams(1) & "'"
    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'52|') > 0 THEN     /* 과탐-화학1 */"
    sStr = sStr & "             '화1'"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             ' '"
    sStr = sStr & "         END END AS 탐구2,"
    sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'" & constSatamCodes(2) & "|') > 0 THEN          /* 사탐-경제 */"
    sStr = sStr & "             '" & constSatams(2) & "'"
    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'53|') > 0 THEN     /* 과탐-생명과학1 */"
    sStr = sStr & "             '생1'"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             ' '"
    sStr = sStr & "         END END AS 탐구3,"
    sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'" & constSatamCodes(3) & "|') > 0 THEN          /* 사탐-한국근현대 */"
    sStr = sStr & "             '" & constSatams(3) & "'"
    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'54|') > 0 THEN     /* 과탐-지구과학1 */"
    sStr = sStr & "             '지1'"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             ' '"
    sStr = sStr & "         END END AS 탐구4,"
    sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'" & constSatamCodes(4) & "|') > 0 THEN          /* 사탐-세계사 */"
    sStr = sStr & "             '" & constSatams(4) & "'"
    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'55|') > 0 THEN     /* 과탐-물리2 */"
    sStr = sStr & "             '물2'"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             ' '"
    sStr = sStr & "         END END AS 탐구5,"
    sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'" & constSatamCodes(5) & "|') > 0 THEN          /* 사탐-경제지리 */"
    sStr = sStr & "             '" & constSatams(5) & "'"
    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'56|') > 0 THEN     /* 과탐-화학2 */"
    sStr = sStr & "             '화2'"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             ' '"
    sStr = sStr & "         END END AS 탐구6,"
    sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'" & constSatamCodes(6) & "|') > 0 THEN          /* 사탐-한국지리 */"
    sStr = sStr & "             '" & constSatams(6) & "'"
    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'57|') > 0 THEN     /* 과탐-생명과학2 */"
    sStr = sStr & "             '생2'"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             ' '"
    sStr = sStr & "         END END AS 탐구7,"
    sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'" & constSatamCodes(7) & "|') > 0 THEN          /* 사탐-정치 */"
    sStr = sStr & "             '" & constSatams(7) & "'"
    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'58|') > 0 THEN     /* 과탐-지구과학2 */"
    sStr = sStr & "             '지2'"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             ' '"
    sStr = sStr & "         END END AS 탐구8,"
    sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'" & constSatamCodes(8) & "|') > 0 THEN          /* 사탐-사회문화 */"
    sStr = sStr & "             '" & constSatams(8) & "'"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             ' '"
    sStr = sStr & "         END AS 탐구9,"
    sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'" & constSatamCodes(9) & "|') > 0 THEN          /* 사탐-법과사회 */"
    sStr = sStr & "             '" & constSatams(9) & "'"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             ' '"
    sStr = sStr & "         END AS 탐구10,"
    sStr = sStr & " '' AS 탐구11,"
    
    sStr = sStr & "  "
    sStr = sStr & "      /* 제2외국어 & 수리 */"
    sStr = sStr & "              CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'31|') > 0 THEN '독어'"
    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'32|') > 0 THEN '일어'"
    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'33|') > 0 THEN '에파'"
    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'34|') > 0 THEN '불어'"
    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'35|') > 0 THEN '중어'"
    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'36|') > 0 THEN '한문'"
    
    '<< 송파 >> : 2008.01.09
    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'37|') > 0 THEN '언어'"
    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'38|') > 0 THEN '수리'"
    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'39|') > 0 THEN '영어'"
    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'40|') > 0 THEN '세계사'"
    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'41|') > 0 THEN '세지'"
    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'42|') > 0 THEN '아랍어'"
    
    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL4,'81|') > 0 THEN '미적'"
    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL4,'82|') > 0 THEN '이산'"
    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL4,'83|') > 0 THEN '확률'"
    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL4,'84|') > 0 THEN '나형'"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             ' '"
    sStr = sStr & "         END END END END END END END END END END END END END END END END 제2선택,"
    sStr = sStr & "  "
    sStr = sStr & "      /* 논술 */"
    sStr = sStr & "         CASE WHEN INSTR(SEL5,'91|') > 0 THEN         /* 언어 */"
    sStr = sStr & "             '언어'"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             ' '"
    sStr = sStr & "         END 언어논술,"
    sStr = sStr & "         CASE WHEN INSTR(SEL5,'92|') > 0 THEN         /* 수리 */"
    sStr = sStr & "             '수리'"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             ' '"
    sStr = sStr & "         END 수리논술,"
    sStr = sStr & "         CASE WHEN INSTR(SEL5,'93|') > 0 THEN         /* 외국어 */"      '< 변경
    sStr = sStr & "             '외국어'"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             ' '"
    sStr = sStr & "         END 사탐논술,"
    sStr = sStr & "         CASE WHEN INSTR(SEL5,'94|') > 0 THEN         /*  */"            '< 변경
    sStr = sStr & "             ' '"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             ' '"
    sStr = sStr & "         END 과탐논술,"
    sStr = sStr & "  "
    sStr = sStr & "         CY_ACNT AS 가상계좌, TOT_AMT AS 전체금액    ,"
    sStr = sStr & "         NVL(BASE_AMT1    ,0) AS 기본금액1  ,"
    sStr = sStr & "         NVL(BASE_AMT2    ,0) AS 기본금액2  ,"
    sStr = sStr & "         NVL(BASE_AMT3    ,0) AS 기본금액3  ,"
    sStr = sStr & "         NVL(BASE_AMT4    ,0) AS 기본금액4  ,"
    sStr = sStr & "         NVL(BASE_AMT5    ,0) AS 기본금액5  ,"
    sStr = sStr & "         NVL(BASE_AMT6    ,0) AS 기본금액6  ,"
    sStr = sStr & "         NVL(BASE_AMT7    ,0) AS 기본금액7  ,"
    sStr = sStr & "         NVL(BASE_AMT8    ,0) AS 기본금액8  ,"
    sStr = sStr & "         NVL(TAMGU_AMT1   ,0) AS 탐구영역금액1 ,"
    sStr = sStr & "         NVL(TAMGU_AMT2   ,0) AS 탐구영역금액2 ,"
    sStr = sStr & "         NVL(TAMGU_AMT3   ,0) AS 탐구영역금액3 ,"
    sStr = sStr & "         NVL(TAMGU_AMT4   ,0) AS 탐구영역금액4 ,"
    sStr = sStr & "         NVL(TAMGU_AMT5   ,0) AS 탐구영역금액5 ,"
    sStr = sStr & "         NVL(TAMGU_AMT6   ,0) AS 탐구영역금액6 ,"
    sStr = sStr & "         NVL(TAMGU_AMT7   ,0) AS 탐구영역금액7 ,"
    sStr = sStr & "         NVL(TAMGU_AMT8   ,0) AS 탐구영역금액8 ,"
    sStr = sStr & "         NVL(TAMGU_AMT9   ,0) AS 탐구영역금액9 ,"
    sStr = sStr & "         NVL(TAMGU_AMT10  ,0) AS 탐구영역금액10,"
    sStr = sStr & "         NVL(TAMGU_AMT11  ,0) AS 탐구영역금액11,"
    
    sStr = sStr & "      /* 탐구 성적 떄문에 처리.. */"
    sStr = sStr & "              CASE WHEN SEL6 > ' ' AND INSTR(SEL6,'" & constSatamCodes(0) & "') > 0 THEN '" & constSatams(0) & "'"
    sStr = sStr & "         ELSE CASE WHEN SEL6 > ' ' AND INSTR(SEL6,'" & constSatamCodes(1) & "') > 0 THEN '" & constSatams(1) & "'"
    sStr = sStr & "         ELSE CASE WHEN SEL6 > ' ' AND INSTR(SEL6,'" & constSatamCodes(2) & "') > 0 THEN '" & constSatams(2) & "'"
    sStr = sStr & "         ELSE CASE WHEN SEL6 > ' ' AND INSTR(SEL6,'" & constSatamCodes(3) & "') > 0 THEN '" & constSatams(3) & "'"
    sStr = sStr & "         ELSE CASE WHEN SEL6 > ' ' AND INSTR(SEL6,'" & constSatamCodes(4) & "') > 0 THEN '" & constSatams(4) & "'"
    sStr = sStr & "         ELSE CASE WHEN SEL6 > ' ' AND INSTR(SEL6,'" & constSatamCodes(5) & "') > 0 THEN '" & constSatams(5) & "'"
    sStr = sStr & "         ELSE CASE WHEN SEL6 > ' ' AND INSTR(SEL6,'" & constSatamCodes(6) & "') > 0 THEN '" & constSatams(6) & "'"
    sStr = sStr & "         ELSE CASE WHEN SEL6 > ' ' AND INSTR(SEL6,'" & constSatamCodes(7) & "') > 0 THEN '" & constSatams(7) & "'"
    sStr = sStr & "         ELSE CASE WHEN SEL6 > ' ' AND INSTR(SEL6,'" & constSatamCodes(8) & "') > 0 THEN '" & constSatams(8) & "'"
    sStr = sStr & "         ELSE CASE WHEN SEL6 > ' ' AND INSTR(SEL6,'" & constSatamCodes(9) & "') > 0 THEN '" & constSatams(9) & "'"
    sStr = sStr & "         ELSE CASE WHEN SEL6 > ' ' AND INSTR(SEL6,'51') > 0 THEN '물I'"
    sStr = sStr & "         ELSE CASE WHEN SEL6 > ' ' AND INSTR(SEL6,'52') > 0 THEN '화I'"
    sStr = sStr & "         ELSE CASE WHEN SEL6 > ' ' AND INSTR(SEL6,'53') > 0 THEN '생I'"
    sStr = sStr & "         ELSE CASE WHEN SEL6 > ' ' AND INSTR(SEL6,'54') > 0 THEN '지I'"
    sStr = sStr & "         ELSE CASE WHEN SEL6 > ' ' AND INSTR(SEL6,'55') > 0 THEN '물II'"
    sStr = sStr & "         ELSE CASE WHEN SEL6 > ' ' AND INSTR(SEL6,'56') > 0 THEN '화II'"
    sStr = sStr & "         ELSE CASE WHEN SEL6 > ' ' AND INSTR(SEL6,'57') > 0 THEN '생II'"
    sStr = sStr & "         ELSE CASE WHEN SEL6 > ' ' AND INSTR(SEL6,'58') > 0 THEN '지II'"
    sStr = sStr & "         END END END END END END END END END END END END END END END END END END SEL_X6,"
    
    sStr = sStr & "         K_NUM AS 언어점수, M_NUM AS 수학점수, E_NUM AS 영어점수, "
    sStr = sStr & "         (NVL(K_NUM,0)+NVL(M_NUM,0)+NVL(E_NUM,0)) AS 전체점수,"
    sStr = sStr & "         N_NUM AS 내신등급,"
    
    
    sStr = sStr & "         DECODE(SEL1_SCH,'N','노량진','K','강남','S','송파','P','송파마이맥','M','강남마이맥', 'W', '주말법의대','Q','야간법의대','Y','양재', 'B','부산') AS 제1지망,"
    sStr = sStr & "         DECODE(SEL2_SCH,'N','노량진','K','강남','S','송파','P','송파마이맥','M','강남마이맥', 'W', '주말법의대','Q','야간법의대','Y','양재', 'B','부산') AS 제2지망,"
    
    sStr = sStr & "         DECODE(PASS1,'N','노량진','K','강남','S','송파','P','송파마이맥','M','강남마이맥', 'W', '주말법의대','Q','야간법의대','Y','양재', 'B','부산') AS 합격1   ,"
    sStr = sStr & "         DECODE(PASS2,'N','노량진','K','강남','S','송파','P','송파마이맥','M','강남마이맥', 'W', '주말법의대','Q','야간법의대','Y','양재', 'B','부산') AS 합격2   ,"
    sStr = sStr & "         DECODE(PASS3,'N','노량진','K','강남','S','송파','P','송파마이맥','M','강남마이맥', 'W', '주말법의대','Q','야간법의대','Y','양재', 'B','부산') AS 합격3   ,"
    sStr = sStr & "         DECODE(PASS4,'N','노량진','K','강남','S','송파','P','송파마이맥','M','강남마이맥', 'W', '주말법의대','Q','야간법의대','Y','양재', 'B','부산') AS 합격4   ,"
    
    
    sStr = sStr & "         DECODE(SEX,'M','남','F','여') AS 성별        , "
    sStr = sStr & "         SUBSTR(ZIP,1,3)||'-'||SUBSTR(ZIP,4,3) AS 우편번호, ADDR1 AS 우편주소      , ADDR2 AS 상세주소     ,"
    sStr = sStr & "         TEL AS 전화번호, CEL AS 핸드폰        , EMAIL AS 이메일     ,"
    sStr = sStr & "         HIGH_SCH AS 고등학교 , GRADE_YEAR AS 졸업년도 ,"
    sStr = sStr & "         PRNT_NM AS 학부모명 , DECODE(PRNT_RLTN,'1','부','2','모','3','기타') AS 학부모관계, "
    sStr = sStr & "         SUBSTR(PRNT_ZIP,1,3)||'-'||SUBSTR(PRNT_ZIP,4,3) AS 학부모_우편번호, PRNT_ADDR1 AS 학부모_우편주소 , PRNT_ADDR2 AS 학부모_상세주소,"
    sStr = sStr & "         PRNT_TEL AS 학부모_전화번호  , PRNT_CEL AS 학부모_핸드폰   , PRNT_JOB AS 학부모_직업   , PRNT_W_TEL AS 학부모_직장전화 ,"
    sStr = sStr & "         PHOTO_PATH AS 사진저장장소, "
    sStr = sStr & "         DECODE(R_WAY,'1','학원등록','2','인터넷등록','3','학원등록') AS 등록번호, "
    sStr = sStr & "         ORD_NO AS 주문번호, "
    sStr = sStr & "         ACID||EXMID AS 이미지파일명, "
    'sStr = sStr & "         DECODE(ACID,'" & Trim(basModule.SchCD) & "','',ACID) AS WANT_ACID "
    'sStr = sStr & "         DECODE(ACID,'" & Trim(basModule.SchCD) & "','" & Trim(basModule.SchCD) & "',ACID) AS WANT_ACID, "       '< TEST
    sStr = sStr & "         REGDATE AS 등록일자, GET_PAYGUBN(ORD_NO) AS 결재방법, CASH_BILL_NUM AS 현금영수증,"
    sStr = sStr & "         DECODE(MU_TYPE,'1','수능평가','2','6월 평가원','3','9월 평가원','4','6월 평가원','5','9','내신등급','9월 평가원','') AS 등급, "
    sStr = sStr & "         CL_CLOSE AS 완료년월 "
    
    sStr = sStr & " , "
        sStr = sStr & "        J01 AS 언어          ,"
        sStr = sStr & "        K01 AS 언어_백       ,"
        sStr = sStr & "        J02 AS 수리가        ,"
        sStr = sStr & "        K02 AS 수리가형_백   ,"
        sStr = sStr & "        J03 AS 외국어        ,"
        sStr = sStr & "        K03 AS 외국어_백     ,"
                                   
        sStr = sStr & "        J04 AS " & constSatams(0) & "_물1      ,"
        sStr = sStr & "        K04 AS " & constSatams(0) & "_물1_백   ,"
        sStr = sStr & "        J05 AS " & constSatams(1) & "_화1      ,"
        sStr = sStr & "        K05 AS " & constSatams(1) & "_화1_백   ,"
        sStr = sStr & "        J06 AS " & constSatams(2) & "_생1      ,"
        sStr = sStr & "        K06 AS " & constSatams(2) & "_생1_백   ,"
        sStr = sStr & "        J07 AS " & constSatams(3) & "_지학1    ,"
        sStr = sStr & "        K07 AS " & constSatams(3) & "_지학1_백 ,"
        sStr = sStr & "        J08 AS " & constSatams(4) & "_물2      ,"
        sStr = sStr & "        K08 AS " & constSatams(4) & "_물2_백   ,"
        sStr = sStr & "        J09 AS " & constSatams(5) & "_화2      ,"
        sStr = sStr & "        K09 AS " & constSatams(5) & "_화2_백   ,"
        sStr = sStr & "        J10 AS " & constSatams(6) & "_생2      ,"
        sStr = sStr & "        K10 AS " & constSatams(6) & "_생2_백   ,"
        sStr = sStr & "        J11 AS " & constSatams(7) & "_지학2    ,"
        sStr = sStr & "        K11 AS " & constSatams(7) & "_지학2_백 ,"
                                   
        sStr = sStr & "        J12 AS " & constSatams(8) & "          ,"
        sStr = sStr & "        K12 AS " & constSatams(8) & "_백       ,"
        sStr = sStr & "        J13 AS " & constSatams(9) & "          ,"
        sStr = sStr & "        K13 AS " & constSatams(9) & "_백       ,"
        sStr = sStr & " ' ' AS K14, "
        sStr = sStr & " ' ' AS J14, "
        sStr = sStr & "        J15 AS 독어_미적     ,"
        sStr = sStr & "        K15 AS 독어_미적_백  ,"
        sStr = sStr & "        J16 AS 일어_이산     ,"
        sStr = sStr & "        K16 AS 일어_이산_백  ,"
        sStr = sStr & "        J17 AS 에파_확통     ,"
        sStr = sStr & "        K17 AS 에파_확통_백  ,"
        sStr = sStr & "        J18 AS 불어_수리나   ,"
        sStr = sStr & "        K18 AS 불어_수리나_백,"
                                   
        sStr = sStr & "        J19 AS 중어          ,"
        sStr = sStr & "        K19 AS 중어_백       ,"
        sStr = sStr & "        J20 AS 한문          ,"
        sStr = sStr & "        K20 AS 한문_백       ,"
        sStr = sStr & "        J21 AS 아랍어        ,"
        sStr = sStr & "        K21 AS 아랍어_백     ,"
        
        ' 노량진 요청에 의한 지원단대... 그러나 한욱씨께서 위에다 추가해놓으셨다.. 그래서 엑셀이 안되었었다.
        ' 밑에다가로.. 변경..
        sStr = sStr & "        D_UNIVCD AS 지원대학, D_MAJORCD AS 지원단대, "
        
        ' 클리닉 추가
        sStr = sStr & " CASE WHEN SEL7 > ' ' AND    INSTR (SEL7, '101') > 0 THEN '" & g_sClinic_Ls(0) & "'"
        sStr = sStr & " ELSE CASE WHEN SEL7 > ' ' AND    INSTR (SEL7, '102') > 0 THEN '" & g_sClinic_Ls(1) & "'"
        sStr = sStr & " ELSE CASE WHEN SEL7 > ' ' AND    INSTR (SEL7, '103') > 0 THEN '" & g_sClinic_Ls(2) & "'"
        sStr = sStr & " ELSE CASE WHEN SEL7 > ' ' AND    INSTR (SEL7, '104') > 0 THEN '" & g_sClinic_Ls(3) & "'"
        sStr = sStr & " END END END END 국어클리닉,"
        sStr = sStr & " CASE WHEN SEL7 > ' ' AND    INSTR (SEL7, '111') > 0 THEN '" & g_sClinic_Ms(0) & "'"
        sStr = sStr & " ELSE CASE WHEN SEL7 > ' ' AND    INSTR (SEL7, '112') > 0 THEN '" & g_sClinic_Ms(1) & "'"
        sStr = sStr & " ELSE CASE WHEN SEL7 > ' ' AND    INSTR (SEL7, '113') > 0 THEN '" & g_sClinic_Ms(2) & "'"
        sStr = sStr & " ELSE CASE WHEN SEL7 > ' ' AND    INSTR (SEL7, '114') > 0 THEN '" & g_sClinic_Ms(3) & "'"
        sStr = sStr & " END END END END 수학클리닉,"
        sStr = sStr & " CASE WHEN SEL7 > ' ' AND    INSTR (SEL7, '121') > 0 THEN '" & g_sClinic_Es(0) & "'"
        sStr = sStr & " ELSE CASE WHEN SEL7 > ' ' AND    INSTR (SEL7, '122') > 0 THEN '" & g_sClinic_Es(1) & "'"
        sStr = sStr & " ELSE CASE WHEN SEL7 > ' ' AND    INSTR (SEL7, '123') > 0 THEN '" & g_sClinic_Es(2) & "'"
        sStr = sStr & " ELSE CASE WHEN SEL7 > ' ' AND    INSTR (SEL7, '124') > 0 THEN '" & g_sClinic_Es(3) & "'"
        sStr = sStr & " END END END END 영어클리닉"
        
        sStr = sStr & "    FROM ( "
    
            sStr = sStr & "  SELECT A.SCHNO           ,"
            sStr = sStr & "         MAX(ACID      ) AS ACID       ,"
            sStr = sStr & "         MAX(EXMID     ) AS EXMID      ,"
            sStr = sStr & "         MAX(STDNM     ) AS STDNM      ,"
            sStr = sStr & "         MAX(D_UNIVCD     ) AS D_UNIVCD      ,"
            sStr = sStr & "         MAX(D_MAJORCD     ) AS D_MAJORCD      ,"
            sStr = sStr & "         MAX(Birth_ymd     ) AS Birth_ymd      ,"
            sStr = sStr & "         MAX(EXMTYPE   ) AS EXMTYPE    , MAX(KAEYOL    ) AS KAEYOL     ,"
            sStr = sStr & "         MAX(SEL1      ) AS SEL1       , MAX(SEL2      ) AS SEL2       , MAX(SEL3      ) AS SEL3      , MAX(SEL4      ) AS SEL4      , MAX(SEL5      ) AS  SEL5 , MAX(SEL6      ) AS  SEL6, MAX(SEL7      ) AS  SEL7           ,"
            sStr = sStr & "         MAX(K_NUM     ) AS K_NUM      , MAX(M_NUM     ) AS M_NUM      , MAX(E_NUM     ) AS E_NUM     , MAX(N_NUM     ) AS N_NUM     , MAX(TOT_NUM   ) AS TOT_NUM   ,"
            sStr = sStr & "         MAX(SEL1_SCH  ) AS SEL1_SCH   , MAX(SEL2_SCH  ) AS SEL2_SCH   ,"
            sStr = sStr & "         MAX(PASS1     ) AS PASS1      , MAX(PASS2     ) AS PASS2      , MAX(PASS3     ) AS PASS3     , MAX(PASS4     ) AS PASS4     , MAX(CL_CLOSE  ) AS  CL_CLOSE  ,"
            sStr = sStr & "         MAX(CY_ACNT   ) AS CY_ACNT    , MAX(TOT_AMT   ) AS TOT_AMT    ,"
            sStr = sStr & "         MAX(BASE_AMT1 ) AS BASE_AMT1  , MAX(BASE_AMT2 ) AS BASE_AMT2  , MAX(BASE_AMT3 ) AS BASE_AMT3 , MAX(BASE_AMT4 ) AS BASE_AMT4 ,"
            sStr = sStr & "         MAX(BASE_AMT5 ) AS BASE_AMT5  , MAX(BASE_AMT6 ) AS BASE_AMT6  , MAX(BASE_AMT7 ) AS BASE_AMT7 , MAX(BASE_AMT8 ) AS BASE_AMT8 ,"
            sStr = sStr & "         MAX(TAMGU_AMT1) AS TAMGU_AMT1 , MAX(TAMGU_AMT2) AS TAMGU_AMT2 , MAX(TAMGU_AMT3) AS TAMGU_AMT3, MAX(TAMGU_AMT4) AS TAMGU_AMT4, MAX(TAMGU_AMT5) AS  TAMGU_AMT5,"
            sStr = sStr & "         MAX(TAMGU_AMT6) AS TAMGU_AMT6 , MAX(TAMGU_AMT7) AS TAMGU_AMT7 , MAX(TAMGU_AMT8) AS TAMGU_AMT8, MAX(TAMGU_AMT9) AS TAMGU_AMT9, MAX(TAMGU_AMT10) AS TAMGU_AMT10, MAX(TAMGU_AMT11) AS TAMGU_AMT11,"
            sStr = sStr & "         MAX(SEX       ) AS SEX        ,"
            sStr = sStr & "         MAX(ZIP       ) AS ZIP        , MAX(ADDR1     ) AS ADDR1      , MAX(ADDR2     ) AS ADDR2     ,"
            sStr = sStr & "         MAX(TEL       ) AS TEL        , MAX(CEL       ) AS CEL        , MAX(EMAIL     ) AS EMAIL     ,"
            sStr = sStr & "         MAX(HIGH_SCH  ) AS HIGH_SCH   , MAX(GRADE_YEAR) AS GRADE_YEAR ,"
            sStr = sStr & "         MAX(PRNT_NM   ) AS PRNT_NM    , MAX(PRNT_RLTN ) AS PRNT_RLTN  ,"
            sStr = sStr & "         MAX(PRNT_ZIP  ) AS PRNT_ZIP   , MAX(PRNT_ADDR1) AS PRNT_ADDR1 , MAX(PRNT_ADDR2) AS PRNT_ADDR2,"
            sStr = sStr & "         MAX(PRNT_TEL  ) AS PRNT_TEL   , MAX(PRNT_CEL  ) AS PRNT_CEL   , MAX(PRNT_JOB  ) AS PRNT_JOB  , MAX(PRNT_W_TEL) AS PRNT_W_TEL,"
            sStr = sStr & "         MAX(PHOTO_PATH) AS PHOTO_PATH , MAX(R_WAY     ) AS R_WAY      , MAX(ORD_NO    ) AS ORD_NO    , "
            sStr = sStr & "         MAX(TO_CHAR(REGDATE,'YYYY-MM-DD HH24:MI:SS')) AS REGDATE      , MAX(MU_TYPE   ) AS MU_TYPE   , MAX(CASH_BILL_NUM) AS CASH_BILL_NUM"
            
                    sStr = sStr & " , "
                    sStr = sStr & "        SUM(J01) AS J01,"
                    sStr = sStr & "        SUM(K01) AS K01,"
                    sStr = sStr & "        SUM(J02) AS J02,"
                    sStr = sStr & "        SUM(K02) AS K02,"
                    sStr = sStr & "        SUM(J03) AS J03,"
                    sStr = sStr & "        SUM(K03) AS K03,"
                    
                    sStr = sStr & "        SUM(J04) AS J04,"
                    sStr = sStr & "        SUM(K04) AS K04,"
                    sStr = sStr & "        SUM(J05) AS J05,"
                    sStr = sStr & "        SUM(K05) AS K05,"
                    sStr = sStr & "        SUM(J06) AS J06,"
                    sStr = sStr & "        SUM(K06) AS K06,"
                    sStr = sStr & "        SUM(J07) AS J07,"
                    sStr = sStr & "        SUM(K07) AS K07,"
                    sStr = sStr & "        SUM(J08) AS J08,"
                    sStr = sStr & "        SUM(K08) AS K08,"
                    sStr = sStr & "        SUM(J09) AS J09,"
                    sStr = sStr & "        SUM(K09) AS K09,"
                    sStr = sStr & "        SUM(J10) AS J10,"
                    sStr = sStr & "        SUM(K10) AS K10,"
                    sStr = sStr & "        SUM(J11) AS J11,"
                    sStr = sStr & "        SUM(K11) AS K11,"
                    
                    sStr = sStr & "        SUM(J12) AS J12,"
                    sStr = sStr & "        SUM(K12) AS K12,"
                    sStr = sStr & "        SUM(J13) AS J13,"
                    sStr = sStr & "        SUM(K13) AS K13,"
                    sStr = sStr & "   ' ' AS J14, "
                    sStr = sStr & "   ' ' AS K14, "
                    
                    sStr = sStr & "        SUM(J15) AS J15,"
                    sStr = sStr & "        SUM(K15) AS K15,"
                    sStr = sStr & "        SUM(J16) AS J16,"
                    sStr = sStr & "        SUM(K16) AS K16,"
                    sStr = sStr & "        SUM(J17) AS J17,"
                    sStr = sStr & "        SUM(K17) AS K17,"
                    sStr = sStr & "        SUM(J18) AS J18,"
                    sStr = sStr & "        SUM(K18) AS K18,"
                    
                    sStr = sStr & "        SUM(J19) AS J19,"
                    sStr = sStr & "        SUM(K19) AS K19,"
                    sStr = sStr & "        SUM(J20) AS J20,"
                    sStr = sStr & "        SUM(K20) AS K20,"
                    sStr = sStr & "        SUM(J21) AS J21,"
                    sStr = sStr & "        SUM(K21) AS K21"
            
            sStr = sStr & "    FROM ("
            '---------------------------------------------------------------------------- 전체학생 조회 START
            sStr = sStr & "          SELECT *"
            sStr = sStr & "            FROM CLSTD01TB"
            sStr = sStr & "           WHERE ACID = '" & Trim(basModule.SchCD) & "'"
            sStr = sStr & "             AND EXMID > ' ' "
            
    If Trim(Right(kaeyol, 30)) <> "ALL" Then
            sStr = sStr & "             AND KAEYOL = '" & Trim(Right(kaeyol, 30)) & "'"
    End If
    
    '<< 기간설정 >>
    If day1 <> "" Or day2 <> "" Then
        If Trim(day1) <> "" And Trim(day2) <> "" Then
            sStr = sStr & "                 AND REGDATE BETWEEN TO_DATE('" & Trim(day1) & "000000', 'YYYYMMDDHH24MISS') "
            sStr = sStr & "                                 AND TO_DATE('" & Trim(day2) & "235959', 'YYYYMMDDHH24MISS') "
        ElseIf Trim(day1) = "" And Trim(day2) <> "" Then
            sStr = sStr & "                 AND REGDATE BETWEEN TO_DATE('19000101000000', 'YYYYMMDDHH24MISS') "
            sStr = sStr & "                                 AND TO_DATE('" & Trim(day2) & "235959', 'YYYYMMDDHH24MISS') "
        ElseIf Trim(day1) <> "" And Trim(day2) = "" Then
            sStr = sStr & "                 AND REGDATE BETWEEN TO_DATE('" & Trim(day1) & "000000', 'YYYYMMDDHH24MISS') "
            sStr = sStr & "                                 AND TO_DATE('99991231235959', 'YYYYMMDDHH24MISS') "
        End If
        
    End If
    
            sStr = sStr & "             AND BIGO2 IS NULL "
            sStr = sStr & "          UNION ALL"
            '---------------------------------------------------------------------------- 전체학생 조회 END
            '---------------------------------------------------------------------------- 합격자 조회 START
            sStr = sStr & "          SELECT *"
            sStr = sStr & "            From CLSTD01TB"
            sStr = sStr & "           WHERE (PASS1 = '" & Trim(basModule.SchCD) & "'" & " OR"
            sStr = sStr & "                  PASS2 = '" & Trim(basModule.SchCD) & "'" & " OR"
            sStr = sStr & "                  PASS3 = '" & Trim(basModule.SchCD) & "'" & " OR"
            sStr = sStr & "                  PASS4 = '" & Trim(basModule.SchCD) & "'" & " )"
            sStr = sStr & "             AND EXMID > ' ' "
    If Trim(Right(kaeyol, 30)) <> "ALL" Then
            sStr = sStr & "             AND KAEYOL = '" & Trim(Right(kaeyol, 30)) & "'"
    End If
    
    
    '<< 기간설정 >>
    If day1 <> "" Or day2 <> "" Then
        If Trim(day1) <> "" And Trim(day2) <> "" Then
            sStr = sStr & "                 AND REGDATE BETWEEN TO_DATE('" & Trim(day1) & "000000', 'YYYYMMDDHH24MISS') "
            sStr = sStr & "                                 AND TO_DATE('" & Trim(day2) & "235959', 'YYYYMMDDHH24MISS') "
        ElseIf Trim(day1) = "" And Trim(day2) <> "" Then
            sStr = sStr & "                 AND REGDATE BETWEEN TO_DATE('19000101000000', 'YYYYMMDDHH24MISS') "
            sStr = sStr & "                                 AND TO_DATE('" & Trim(day2) & "235959', 'YYYYMMDDHH24MISS') "
        ElseIf Trim(day1) <> "" And Trim(day2) = "" Then
            sStr = sStr & "                 AND REGDATE BETWEEN TO_DATE('" & Trim(day1) & "000000', 'YYYYMMDDHH24MISS') "
            sStr = sStr & "                                 AND TO_DATE('99991231235959', 'YYYYMMDDHH24MISS') "
        End If
        
    End If
            sStr = sStr & "             AND BIGO2 IS NULL "
            
            sStr = sStr & "          ) A, "
            
            sStr = sStr & "               ("
        
                sStr = sStr & "         SELECT SCHNO,"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '37', DECODE(SUB_NUM,'X',0, SUB_NUM), 0)    AS J01,    /* 언어                  */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '37', DECODE(SUB_BAK,'X',0, SUB_BAK), 0)    AS K01,    /* 백분위  언어          */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '38', DECODE(SUB_NUM,'X',0, SUB_NUM), 0)    AS J02,    /* 수리가형              */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '38', DECODE(SUB_BAK,'X',0, SUB_BAK), 0)    AS K02,    /* 백분위  수리가형      */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '39', DECODE(SUB_NUM,'X',0, SUB_NUM), 0)    AS J03,    /* 외국어                */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '39', DECODE(SUB_BAK,'X',0, SUB_BAK), 0)    AS K03,    /* 백분위  외국어        */"
                
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(0) & "', DECODE(SUB_NUM,'X',0, SUB_NUM), '51', DECODE(SUB_NUM, 'X',0, SUB_NUM), 0) AS J04,    /* 사탐-" & constSatams(0) & "       , 과탐-물리1             */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(0) & "', DECODE(SUB_BAK,'X',0, SUB_BAK), '51', DECODE(SUB_BAK, 'X',0, SUB_BAK), 0) AS K04,    /* 백분위  사탐-" & constSatams(0) & "        , 과탐-물리1     */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(1) & "', DECODE(SUB_NUM,'X',0, SUB_NUM), '52', DECODE(SUB_NUM, 'X',0, SUB_NUM), 0) AS J05,    /* 사탐-" & constSatams(1) & "        , 과탐-화학1             */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(1) & "', DECODE(SUB_BAK,'X',0, SUB_BAK), '52', DECODE(SUB_BAK, 'X',0, SUB_BAK), 0) AS K05,    /* 백분위  사탐-" & constSatams(1) & "        , 과탐-화학1     */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(2) & "', DECODE(SUB_NUM,'X',0, SUB_NUM), '53', DECODE(SUB_NUM, 'X',0, SUB_NUM), 0) AS J06,    /* 사탐-" & constSatams(2) & "        , 과탐-생명과학1             */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(2) & "', DECODE(SUB_BAK,'X',0, SUB_BAK), '53', DECODE(SUB_BAK, 'X',0, SUB_BAK), 0) AS K06,    /* 백분위  사탐-" & constSatams(2) & "        , 과탐-생명과학1     */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(3) & "', DECODE(SUB_NUM,'X',0, SUB_NUM), '54', DECODE(SUB_NUM, 'X',0, SUB_NUM), 0) AS J07,    /* 사탐-" & constSatams(3) & "  , 과탐-지구과학1         */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(3) & "', DECODE(SUB_BAK,'X',0, SUB_BAK), '54', DECODE(SUB_BAK, 'X',0, SUB_BAK), 0) AS K07,    /* 백분위  사탐-" & constSatams(3) & "  , 과탐-지구과학1 */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(4) & "', DECODE(SUB_NUM,'X',0, SUB_NUM), '55', DECODE(SUB_NUM, 'X',0, SUB_NUM), 0) AS J08,    /* 사탐-" & constSatams(4) & "      , 과탐-물리2             */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(4) & "', DECODE(SUB_BAK,'X',0, SUB_BAK), '55', DECODE(SUB_BAK, 'X',0, SUB_BAK), 0) AS K08,    /* 백분위  사탐-" & constSatams(4) & "      , 과탐-물리2     */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(5) & "', DECODE(SUB_NUM,'X',0, SUB_NUM), '56', DECODE(SUB_NUM, 'X',0, SUB_NUM), 0) AS J09,    /* 사탐-" & constSatams(5) & "    , 과탐-화학2             */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(5) & "', DECODE(SUB_BAK,'X',0, SUB_BAK), '56', DECODE(SUB_BAK, 'X',0, SUB_BAK), 0) AS K09,    /* 백분위  사탐-" & constSatams(5) & "    , 과탐-화학2     */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(6) & "', DECODE(SUB_NUM,'X',0, SUB_NUM), '57', DECODE(SUB_NUM, 'X',0, SUB_NUM), 0) AS J10,      /* 사탐-" & constSatams(6) & "    , 과탐-생명과학2           */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(6) & "', DECODE(SUB_BAK,'X',0, SUB_BAK), '57', DECODE(SUB_BAK, 'X',0, SUB_BAK), 0) AS K10,      /* 백분위 사탐-" & constSatams(6) & "    , 과탐-생명과학2    */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(7) & "', DECODE(SUB_NUM,'X',0, SUB_NUM), '58', DECODE(SUB_NUM, 'X',0, SUB_NUM), 0) AS J11,    /* 사탐-" & constSatams(7) & "        , 과탐-지구과학2         */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(7) & "', DECODE(SUB_BAK,'X',0, SUB_BAK), '58', DECODE(SUB_BAK, 'X',0, SUB_BAK), 0) AS K11,    /* 백분위  사탐-" & constSatams(7) & "        , 과탐-지구과학2 */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(8) & "', DECODE(SUB_NUM,'X',0, SUB_NUM), 0)    AS J12,    /* 사탐-" & constSatams(8) & "         */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(8) & "', DECODE(SUB_BAK,'X',0, SUB_BAK), 0)    AS K12,    /* 백분위  사탐-" & constSatams(8) & " */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(9) & "', DECODE(SUB_NUM,'X',0, SUB_NUM), 0)    AS J13,    /* 사탐-" & constSatams(9) & "         */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(9) & "', DECODE(SUB_BAK,'X',0, SUB_BAK), 0)    AS K13,    /* 백분위  사탐-" & constSatams(9) & " */"
                sStr = sStr & " '' AS K14, "
                sStr = sStr & " '' AS J14, "
                
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '31', DECODE(SUB_NUM,'X',0, SUB_NUM), '81', DECODE(SUB_NUM, 'X',0, SUB_NUM), 0) AS J15,    /* 독어             , 미적분                 */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '31', DECODE(SUB_BAK,'X',0, SUB_BAK), '81', DECODE(SUB_BAK, 'X',0, SUB_BAK), 0) AS K15,    /* 백분위  독어             , 미적분         */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '32', DECODE(SUB_NUM,'X',0, SUB_NUM), '82', DECODE(SUB_NUM, 'X',0, SUB_NUM), 0) AS J16,    /* 일어             , 이산수학               */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '32', DECODE(SUB_BAK,'X',0, SUB_BAK), '82', DECODE(SUB_BAK, 'X',0, SUB_BAK), 0) AS K16,    /* 백분위  일어             , 이산수학       */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '33', DECODE(SUB_NUM,'X',0, SUB_NUM), '83', DECODE(SUB_NUM, 'X',0, SUB_NUM), 0) AS J17,    /* 에스파냐         , 확률통계               */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '33', DECODE(SUB_BAK,'X',0, SUB_BAK), '83', DECODE(SUB_BAK, 'X',0, SUB_BAK), 0) AS K17,    /* 백분위  에스파냐         , 확률통계       */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '34', DECODE(SUB_NUM,'X',0, SUB_NUM), '43', DECODE(SUB_NUM, 'X',0, SUB_NUM), 0) AS J18,    /* 불어             , 수리나형               */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '34', DECODE(SUB_BAK,'X',0, SUB_BAK), '43', DECODE(SUB_BAK, 'X',0, SUB_BAK), 0) AS K18,    /* 백분위  불어             , 수리나형       */"
                
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '35', DECODE(SUB_NUM,'X',0, SUB_NUM), 0)    AS J19,    /* 중국어                */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '35', DECODE(SUB_BAK,'X',0, SUB_BAK), 0)    AS K19,    /* 백분위  중국어        */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '36', DECODE(SUB_NUM,'X',0, SUB_NUM), 0)    AS J20,    /* 한문                  */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '36', DECODE(SUB_BAK,'X',0, SUB_BAK), 0)    AS K20,    /* 백분위  한문          */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '42', DECODE(SUB_NUM,'X',0, SUB_NUM), 0)    AS J21,    /* 아랍어                */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '42', DECODE(SUB_BAK,'X',0, SUB_BAK), 0)    AS K21     /* 백분위  아랍어        */"
                sStr = sStr & "           FROM CLSTD03TB"
        
        sStr = sStr & "                ) B"
        sStr = sStr & "        WHERE A.SCHNO = B.SCHNO(+)"
            
            sStr = sStr & "   GROUP BY A.SCHNO"
            '---------------------------------------------------------------------------- 합격자 조회 END
    
    sStr = sStr & "    ) "
    sStr = sStr & " ORDER BY EXMID "
    
    Get_StdExcuteSqlToExcel_N = sStr
End Function


'학생 엑셀 저장 쿼리문 (노량진,송파 이외에)
Public Function Get_StdExcuteSqlToExcel(kaeyol As String, Optional day1 As String, Optional day2 As String) As String
    
    Dim sStr         As String
    Dim ni           As Long
    ni = 0
    
    sStr = ""
    sStr = sStr & "  SELECT  "
    sStr = sStr & "         ACID  AS 학원   , "
    sStr = sStr & "         EXMID AS 수험번호, STDNM AS 학생,"
    sStr = sStr & "         birth_ymd AS 생년월일, "
    sStr = sStr & "         DECODE(EXMTYPE,'0','무시험','1','유시험') AS 시험형태, "
    sStr = sStr & "         DECODE(KAEYOL,'01','인문',"
    sStr = sStr & "                       '02','자연',"
'<< 계열 >> : 2008.01.09
    If Trim(basModule.SchCD) = "N" Then
        sStr = sStr & "                   '03','예체',"
        sStr = sStr & "                   '04','수리(나)',"
        sStr = sStr & "                   '05','인문수능',"
        sStr = sStr & "                   '06','자연수능',"
        
        sStr = sStr & "                   '07','신설인문',"
        sStr = sStr & "                   '08','신설자연',"
        sStr = sStr & "                   '09','신설수능인문',"
        sStr = sStr & "                   '10','신설수능자연',"
        
        sStr = sStr & "                   '11','편)인문',"
        sStr = sStr & "                   '12','편)자연',"
        sStr = sStr & "                   '13','편)예체',"
        sStr = sStr & "                   '14','편)수리(나)',"
        sStr = sStr & "                   '15','편)인문수능',"
        sStr = sStr & "                   '16','편)자연수능',"
        sStr = sStr & "                   '21','서울대인문',"
        sStr = sStr & "                   '22','서울대자연',"
    End If
'<< 계열 >> : 2008.01.10
    If Trim(basModule.SchCD) = "K" Or Trim(basModule.SchCD) = "W" Or Trim(basModule.SchCD) = "Q" Then
        sStr = sStr & "                   '04','주말법대',"
        sStr = sStr & "                   '05','주말의대',"
        sStr = sStr & "                   '06','야간법대',"
        sStr = sStr & "                   '07','야간의대',"
    
        sStr = sStr & "                   '11','선착순인문',"
        sStr = sStr & "                   '12','선착순자연',"
        
        sStr = sStr & "                   '16','선착순인문16',"
        sStr = sStr & "                   '17','선착순자연17',"
        
        sStr = sStr & "                   '19','내신우수자인문',"
        sStr = sStr & "                   '20','내신우수자자연',"
    End If
'<< 계열 >> : 2008.02.15
    If Trim(basModule.SchCD) = "S" Then
        sStr = sStr & "                   '03','예체능',"
        'sStr = sStr & "                   '04','특별자연',"
        
        sStr = sStr & "                   '05','수능인문',"
        sStr = sStr & "                   '06','수능자연',"
        
        sStr = sStr & "                   '11','신설인문',"
        sStr = sStr & "                   '12','신설자연',"
        
        sStr = sStr & "                   '18','인문프리미엄',"
        sStr = sStr & "                   '19','자연프리미엄',"
        
    End If
'<< 계열 >> : 2008.02.15
    If Trim(basModule.SchCD) = "P" Then         '< 마송
        sStr = sStr & "                   '03','특별인문',"
        sStr = sStr & "                   '04','특별자연',"
    End If
    
    If Trim(basModule.SchCD) = "J" Then         '< 양재
        sStr = sStr & "                   '11','신설인문',"
        sStr = sStr & "                   '12','신설자연',"
        
        sStr = sStr & "                   '18','인문프리미엄',"
        sStr = sStr & "                   '19','자연프리미엄',"
    End If
    
'<< 계열 >> : 2009.01.09
    If Trim(basModule.SchCD) = "B" Then         '< 부산7
        sStr = sStr & "                   '05','선행인문',"
        sStr = sStr & "                   '06','선행자연',"
        sStr = sStr & "                   '07','연고대인문',"
        sStr = sStr & "                   '08','연고대자연',"
        sStr = sStr & "                   '09','심화인문',"
        sStr = sStr & "                   '10','심화자연',"
    End If
    
    sStr = sStr & "                       '','기타') AS 계열,"
    
    sStr = sStr & "     /* 사탐, 과탐 분리 */"
    For ni = 0 To SATAM_COUNT - 1
        sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'" & constSatamCodes(ni) & "|') > 0 THEN          /* 사탐-" & constSatams(ni) & " */"
        sStr = sStr & "             '" & constSatams(ni) & "'"
        sStr = sStr & "         ELSE "

        If ni < GWATAM_COUNT - 1 Then
            sStr = sStr & "         CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'" & constGwatamCodes(ni) & "|') > 0 THEN     /* 과탐-" & constGwatams(ni) & " */"
            sStr = sStr & "             '" & constGwatams(ni) & "'"
            sStr = sStr & "         ELSE"
            sStr = sStr & "             ' '"
            sStr = sStr & "         END "
        Else
            sStr = sStr & "         ' '"
        End If
        sStr = sStr & "         END AS 탐구" & CStr(ni) & ", "

    Next ni
    
    
    If basModule.SchCD = "J" Then
        sStr = sStr & "         CASE WHEN SEL1 > ' ' AND INSTR(SEL1,'" & TGANG_CODE & "|') > 0 THEN          /* 사탐-특강 */"
        sStr = sStr & "             '특강'"
        sStr = sStr & "         ELSE "
        sStr = sStr & "             CASE WHEN SEL3 > ' ' AND INSTR(SEL3,'" & TGANG_CODE & "|') > 0 THEN     /* 과탐-특강*/"
        sStr = sStr & "                '특강'"
        sStr = sStr & "             ELSE"
        sStr = sStr & "                 ' '"
        sStr = sStr & "             END "
        sStr = sStr & "         END AS 탐구11, "
    End If
    

    sStr = sStr & "  "
    sStr = sStr & "      /* 제2외국어 & 수리 */"
    sStr = sStr & "              CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'31|') > 0 THEN '독어'"
    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'32|') > 0 THEN '일어'"
    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'33|') > 0 THEN '에파'"
    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'34|') > 0 THEN '불어'"
    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'35|') > 0 THEN '중어'"
    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'36|') > 0 THEN '한문'"
    
    '<< 송파 >> : 2008.01.09
    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'37|') > 0 THEN '언어'"
    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'38|') > 0 THEN '수리'"
    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'39|') > 0 THEN '영어'"
    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'40|') > 0 THEN '세계사'"
    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'41|') > 0 THEN '세지'"
    sStr = sStr & "         ELSE CASE WHEN SEL1 > ' ' AND INSTR(SEL2,'42|') > 0 THEN '아랍어'"
    
    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL4,'81|') > 0 THEN '미적'"
    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL4,'82|') > 0 THEN '이산'"
    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL4,'83|') > 0 THEN '확률'"
    sStr = sStr & "         ELSE CASE WHEN SEL3 > ' ' AND INSTR(SEL4,'84|') > 0 THEN '나형'"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             ' '"
    sStr = sStr & "         END END END END END END END END END END END END END END END END 제2선택,"
    sStr = sStr & "  "
    sStr = sStr & "      /* 논술 */"
    sStr = sStr & "         CASE WHEN INSTR(SEL5,'91|') > 0 THEN         /* 언어 */"
    sStr = sStr & "             '언어'"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             ' '"
    sStr = sStr & "         END 언어논술,"
    sStr = sStr & "         CASE WHEN INSTR(SEL5,'92|') > 0 THEN         /* 수리 */"
    sStr = sStr & "             '수리'"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             ' '"
    sStr = sStr & "         END 수리논술,"
    sStr = sStr & "         CASE WHEN INSTR(SEL5,'93|') > 0 THEN         /* 외국어 */"      '< 변경
    sStr = sStr & "             '외국어'"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             ' '"
    sStr = sStr & "         END 사탐논술,"
    sStr = sStr & "         CASE WHEN INSTR(SEL5,'94|') > 0 THEN         /*  */"            '< 변경
    sStr = sStr & "             ' '"
    sStr = sStr & "         ELSE"
    sStr = sStr & "             ' '"
    sStr = sStr & "         END 과탐논술,"
    sStr = sStr & "  "
    sStr = sStr & "         CY_ACNT AS 가상계좌, TOT_AMT AS 전체금액    ,"
    sStr = sStr & "         NVL(BASE_AMT1    ,0) AS 기본금액1  ,"
    sStr = sStr & "         NVL(BASE_AMT2    ,0) AS 기본금액2  ,"
    sStr = sStr & "         NVL(BASE_AMT3    ,0) AS 기본금액3  ,"
    sStr = sStr & "         NVL(BASE_AMT4    ,0) AS 기본금액4  ,"
    sStr = sStr & "         NVL(BASE_AMT5    ,0) AS 기본금액5  ,"
    sStr = sStr & "         NVL(BASE_AMT6    ,0) AS 기본금액6  ,"
    sStr = sStr & "         NVL(BASE_AMT7    ,0) AS 기본금액7  ,"
    sStr = sStr & "         NVL(BASE_AMT8    ,0) AS 기본금액8  ,"
    sStr = sStr & "         NVL(TAMGU_AMT1   ,0) AS 탐구영역금액1 ,"
    sStr = sStr & "         NVL(TAMGU_AMT2   ,0) AS 탐구영역금액2 ,"
    sStr = sStr & "         NVL(TAMGU_AMT3   ,0) AS 탐구영역금액3 ,"
    sStr = sStr & "         NVL(TAMGU_AMT4   ,0) AS 탐구영역금액4 ,"
    sStr = sStr & "         NVL(TAMGU_AMT5   ,0) AS 탐구영역금액5 ,"
    sStr = sStr & "         NVL(TAMGU_AMT6   ,0) AS 탐구영역금액6 ,"
    sStr = sStr & "         NVL(TAMGU_AMT7   ,0) AS 탐구영역금액7 ,"
    sStr = sStr & "         NVL(TAMGU_AMT8   ,0) AS 탐구영역금액8 ,"
    sStr = sStr & "         NVL(TAMGU_AMT9   ,0) AS 탐구영역금액9 ,"
    sStr = sStr & "         NVL(TAMGU_AMT10  ,0) AS 탐구영역금액10,"
    sStr = sStr & "         NVL(TAMGU_AMT11  ,0) AS 탐구영역금액11,"
    
    sStr = sStr & "         K_NUM AS 언어점수, M_NUM AS 수학점수, E_NUM AS 영어점수, "
    sStr = sStr & "         (NVL(K_NUM,0)+NVL(M_NUM,0)+NVL(E_NUM,0)) AS 전체점수, N_NUM AS 내신등급, "
    
    
    sStr = sStr & "         DECODE(SEL1_SCH,'N','노량진','K','강남','S','송파','P','송파마이맥','M','강남마이맥', 'W', '주말법의대','Q','야간법의대','Y','양재', 'B','부산') AS 제1지망,"
    sStr = sStr & "         DECODE(SEL2_SCH,'N','노량진','K','강남','S','송파','P','송파마이맥','M','강남마이맥', 'W', '주말법의대','Q','야간법의대','Y','양재', 'B','부산') AS 제2지망,"
    
    sStr = sStr & "         DECODE(PASS1,'N','노량진','K','강남','S','송파','P','송파마이맥','M','강남마이맥', 'W', '주말법의대','Q','야간법의대','Y','양재', 'B','부산') AS 합격1   ,"
    sStr = sStr & "         DECODE(PASS2,'N','노량진','K','강남','S','송파','P','송파마이맥','M','강남마이맥', 'W', '주말법의대','Q','야간법의대','Y','양재', 'B','부산') AS 합격2   ,"
    sStr = sStr & "         DECODE(PASS3,'N','노량진','K','강남','S','송파','P','송파마이맥','M','강남마이맥', 'W', '주말법의대','Q','야간법의대','Y','양재', 'B','부산') AS 합격3   ,"
    sStr = sStr & "         DECODE(PASS4,'N','노량진','K','강남','S','송파','P','송파마이맥','M','강남마이맥', 'W', '주말법의대','Q','야간법의대','Y','양재', 'B','부산') AS 합격4   ,"
    
    
    sStr = sStr & "         DECODE(SEX,'M','남','F','여') AS 성별        , "
    sStr = sStr & "         SUBSTR(ZIP,1,3)||'-'||SUBSTR(ZIP,4,3) AS 우편번호, ADDR1 AS 우편주소      , ADDR2 AS 상세주소     ,"
    sStr = sStr & "         TEL AS 전화번호, CEL AS 핸드폰        , EMAIL AS 이메일     ,"
    sStr = sStr & "         HIGH_SCH AS 고등학교 , GRADE_YEAR AS 졸업년도 ,"
    sStr = sStr & "         PRNT_NM AS 학부모명 , DECODE(PRNT_RLTN,'1','부','2','모','3','기타') AS 학부모관계, "
    sStr = sStr & "         SUBSTR(PRNT_ZIP,1,3)||'-'||SUBSTR(PRNT_ZIP,4,3) AS 학부모_우편번호, PRNT_ADDR1 AS 학부모_우편주소 , PRNT_ADDR2 AS 학부모_상세주소,"
    sStr = sStr & "         PRNT_TEL AS 학부모_전화번호  , PRNT_CEL AS 학부모_핸드폰   , PRNT_JOB AS 학부모_직업   , PRNT_W_TEL AS 학부모_직장전화 ,"
    sStr = sStr & "         PHOTO_PATH AS 사진저장장소, "
    sStr = sStr & "         DECODE(R_WAY,'1','학원등록','2','인터넷등록','3','학원등록') AS 등록번호, "
    sStr = sStr & "         ORD_NO AS 주문번호, "
    sStr = sStr & "         ACID||EXMID AS 이미지파일명, "
    'sStr = sStr & "         DECODE(ACID,'" & Trim(basModule.SchCD) & "','',ACID) AS WANT_ACID "
    'sStr = sStr & "         DECODE(ACID,'" & Trim(basModule.SchCD) & "','" & Trim(basModule.SchCD) & "',ACID) AS WANT_ACID, "       '< TEST
    sStr = sStr & "         REGDATE AS 등록일자, GET_PAYGUBN(ORD_NO) AS 결재방법, CASH_BILL_NUM AS 현금영수증,"
    sStr = sStr & "         DECODE(MU_TYPE,'1','수능평가','2','6월 평가원','3','9월 평가원','4','6월 평가원','9','내신등급','5','9월 평가원','') AS 등급, "
    sStr = sStr & "         CL_CLOSE AS 완료년월 ,"
    
    Select Case Trim(basModule.SchCD)
        Case "S"
            'sStr = sStr & " DECODE(PTS_SEL,'1','수능','2','6월 평가원','3','9월 평가원','4','6월 평가원','5','9월 평가원','') AS 구분, "
            sStr = sStr & " DECODE(PTS_SEL,'1','가형','2','나형','') AS 구분, "
        Case "P"
            sStr = sStr & " DECODE(PTS_SEL,'8','수능','9','2010 평가','6','3등급','','') AS 구분, "
        Case Else
            sStr = sStr & " '' AS 구분,"
    End Select
    
        sStr = sStr & "        J01 AS 언어          ,"
        sStr = sStr & "        K01 AS 언어_백       ,"
        sStr = sStr & "        J02 AS 수리가        ,"
        sStr = sStr & "        K02 AS 수리가형_백   ,"
        sStr = sStr & "        J03 AS 외국어        ,"
        sStr = sStr & "        K03 AS 외국어_백     ,"
                                   
        sStr = sStr & "        J04 AS " & constSatams(0) & "_" & constGwatams(0) & "      ,"
        sStr = sStr & "        K04 AS " & constSatams(0) & "_" & constGwatams(0) & "_백   ,"
        sStr = sStr & "        J05 AS " & constSatams(1) & "_" & constGwatams(1) & "      ,"
        sStr = sStr & "        K05 AS " & constSatams(1) & "_" & constGwatams(1) & "_백   ,"
        sStr = sStr & "        J06 AS " & constSatams(2) & "_" & constGwatams(2) & "      ,"
        sStr = sStr & "        K06 AS " & constSatams(2) & "_" & constGwatams(2) & "_백   ,"
        sStr = sStr & "        J07 AS " & constSatams(3) & "_" & constGwatams(3) & "      ,"
        sStr = sStr & "        K07 AS " & constSatams(3) & "_" & constGwatams(3) & "_백   ,"
        sStr = sStr & "        J08 AS " & constSatams(4) & "_" & constGwatams(4) & "      ,"
        sStr = sStr & "        K08 AS " & constSatams(4) & "_" & constGwatams(4) & "_백   ,"
        sStr = sStr & "        J09 AS " & constSatams(5) & "_" & constGwatams(5) & "      ,"
        sStr = sStr & "        K09 AS " & constSatams(5) & "_" & constGwatams(5) & "_백   ,"
        sStr = sStr & "        J10 AS " & constSatams(6) & "_" & constGwatams(6) & "      ,"
        sStr = sStr & "        K10 AS " & constSatams(6) & "_" & constGwatams(6) & "_백   ,"
        sStr = sStr & "        J11 AS " & constSatams(7) & "_" & constGwatams(7) & "      ,"
        sStr = sStr & "        K11 AS " & constSatams(7) & "_" & constGwatams(7) & "_백   ,"
                                   
        sStr = sStr & "        J12 AS " & constSatams(8) & "          ,"
        sStr = sStr & "        K12 AS " & constSatams(8) & "_백       ,"
        sStr = sStr & "        J13 AS " & constSatams(9) & "          ,"
        sStr = sStr & "        K13 AS " & constSatams(9) & "_백       ,"
        sStr = sStr & " '' AS J14, "
        sStr = sStr & " '' AS K14, "
                                           
        sStr = sStr & "        J15 AS 독어_미적     ,"
        sStr = sStr & "        K15 AS 독어_미적_백  ,"
        sStr = sStr & "        J16 AS 일어_이산     ,"
        sStr = sStr & "        K16 AS 일어_이산_백  ,"
        sStr = sStr & "        J17 AS 에파_확통     ,"
        sStr = sStr & "        K17 AS 에파_확통_백  ,"
        sStr = sStr & "        J18 AS 불어_수리나   ,"
        sStr = sStr & "        K18 AS 불어_수리나_백,"
                                   
        sStr = sStr & "        J19 AS 중어          ,"
        sStr = sStr & "        K19 AS 중어_백       ,"
        sStr = sStr & "        J20 AS 한문          ,"
        sStr = sStr & "        K20 AS 한문_백       ,"
        sStr = sStr & "        J21 AS 아랍어        ,"
        sStr = sStr & "        K21 AS 아랍어_백     ,"
        
        ' 노량진 요청에 의한 지원단대... 그러나 한욱씨께서 위에다 추가해놓으셨다.. 그래서 엑셀이 안되었었다.
        ' 밑에다가로.. 변경..
        sStr = sStr & "        D_UNIVCD AS 지원대학, D_MAJORCD AS 지원단대 "
        
        sStr = sStr & "    FROM ( "
    
            sStr = sStr & "  SELECT A.SCHNO           ,"
            sStr = sStr & "         MAX(ACID      ) AS ACID       ,"
            sStr = sStr & "         MAX(EXMID     ) AS EXMID      ,"
            sStr = sStr & "         MAX(STDNM     ) AS STDNM      ,"
            sStr = sStr & "         MAX(D_UNIVCD     ) AS D_UNIVCD      ,"
            sStr = sStr & "         MAX(D_MAJORCD     ) AS D_MAJORCD      ,"
            sStr = sStr & "         MAX(birth_ymd     ) AS birth_ymd      ,"
            sStr = sStr & "         MAX(EXMTYPE   ) AS EXMTYPE    , MAX(KAEYOL    ) AS KAEYOL     ,"
            sStr = sStr & "         MAX(SEL1      ) AS SEL1       , MAX(SEL2      ) AS SEL2       , MAX(SEL3      ) AS SEL3      , MAX(SEL4      ) AS SEL4      , MAX(SEL5      ) AS  SEL5      ,"
            sStr = sStr & "         MAX(K_NUM     ) AS K_NUM      , MAX(M_NUM     ) AS M_NUM      , MAX(E_NUM     ) AS E_NUM     , MAX(TOT_NUM   ) AS TOT_NUM   , MAX(N_NUM   ) AS N_NUM   ,"
            sStr = sStr & "         MAX(SEL1_SCH  ) AS SEL1_SCH   , MAX(SEL2_SCH  ) AS SEL2_SCH   ,"
            sStr = sStr & "         MAX(PASS1     ) AS PASS1      , MAX(PASS2     ) AS PASS2      , MAX(PASS3     ) AS PASS3     , MAX(PASS4     ) AS PASS4     , MAX(CL_CLOSE  ) AS  CL_CLOSE  ,"
            sStr = sStr & "         MAX(CY_ACNT   ) AS CY_ACNT    , MAX(TOT_AMT   ) AS TOT_AMT    ,"
            sStr = sStr & "         MAX(BASE_AMT1 ) AS BASE_AMT1  , MAX(BASE_AMT2 ) AS BASE_AMT2  , MAX(BASE_AMT3 ) AS BASE_AMT3 , MAX(BASE_AMT4 ) AS BASE_AMT4 ,"
            sStr = sStr & "         MAX(BASE_AMT5 ) AS BASE_AMT5  , MAX(BASE_AMT6 ) AS BASE_AMT6  , MAX(BASE_AMT7 ) AS BASE_AMT7 , MAX(BASE_AMT8 ) AS BASE_AMT8 ,"
            sStr = sStr & "         MAX(TAMGU_AMT1) AS TAMGU_AMT1 , MAX(TAMGU_AMT2) AS TAMGU_AMT2 , MAX(TAMGU_AMT3) AS TAMGU_AMT3, MAX(TAMGU_AMT4) AS TAMGU_AMT4, MAX(TAMGU_AMT5) AS  TAMGU_AMT5,"
            sStr = sStr & "         MAX(TAMGU_AMT6) AS TAMGU_AMT6 , MAX(TAMGU_AMT7) AS TAMGU_AMT7 , MAX(TAMGU_AMT8) AS TAMGU_AMT8, MAX(TAMGU_AMT9) AS TAMGU_AMT9, MAX(TAMGU_AMT10) AS TAMGU_AMT10, MAX(TAMGU_AMT11) AS TAMGU_AMT11,"
            sStr = sStr & "         MAX(SEX       ) AS SEX        ,"
            sStr = sStr & "         MAX(ZIP       ) AS ZIP        , MAX(ADDR1     ) AS ADDR1      , MAX(ADDR2     ) AS ADDR2     ,"
            sStr = sStr & "         MAX(TEL       ) AS TEL        , MAX(CEL       ) AS CEL        , MAX(EMAIL     ) AS EMAIL     ,"
            sStr = sStr & "         MAX(HIGH_SCH  ) AS HIGH_SCH   , MAX(GRADE_YEAR) AS GRADE_YEAR ,"
            sStr = sStr & "         MAX(PRNT_NM   ) AS PRNT_NM    , MAX(PRNT_RLTN ) AS PRNT_RLTN  ,"
            sStr = sStr & "         MAX(PRNT_ZIP  ) AS PRNT_ZIP   , MAX(PRNT_ADDR1) AS PRNT_ADDR1 , MAX(PRNT_ADDR2) AS PRNT_ADDR2,"
            sStr = sStr & "         MAX(PRNT_TEL  ) AS PRNT_TEL   , MAX(PRNT_CEL  ) AS PRNT_CEL   , MAX(PRNT_JOB  ) AS PRNT_JOB  , MAX(PRNT_W_TEL) AS PRNT_W_TEL,"
            sStr = sStr & "         MAX(PHOTO_PATH) AS PHOTO_PATH , MAX(R_WAY     ) AS R_WAY      , MAX(ORD_NO    ) AS ORD_NO    , "
            sStr = sStr & "         MAX(TO_CHAR(REGDATE,'YYYY-MM-DD HH24:MI:SS')) AS REGDATE      , MAX(PTS_SEL   ) AS PTS_SEL   , MAX(MU_TYPE   ) AS MU_TYPE   , MAX(CASH_BILL_NUM) AS CASH_BILL_NUM"
            
                    sStr = sStr & " , "
                    sStr = sStr & "        SUM(J01) AS J01,"
                    sStr = sStr & "        SUM(K01) AS K01,"
                    sStr = sStr & "        SUM(J02) AS J02,"
                    sStr = sStr & "        SUM(K02) AS K02,"
                    sStr = sStr & "        SUM(J03) AS J03,"
                    sStr = sStr & "        SUM(K03) AS K03,"
                    
                    sStr = sStr & "        SUM(J04) AS J04,"
                    sStr = sStr & "        SUM(K04) AS K04,"
                    sStr = sStr & "        SUM(J05) AS J05,"
                    sStr = sStr & "        SUM(K05) AS K05,"
                    sStr = sStr & "        SUM(J06) AS J06,"
                    sStr = sStr & "        SUM(K06) AS K06,"
                    sStr = sStr & "        SUM(J07) AS J07,"
                    sStr = sStr & "        SUM(K07) AS K07,"
                    sStr = sStr & "        SUM(J08) AS J08,"
                    sStr = sStr & "        SUM(K08) AS K08,"
                    sStr = sStr & "        SUM(J09) AS J09,"
                    sStr = sStr & "        SUM(K09) AS K09,"
                    sStr = sStr & "        SUM(J10) AS J10,"
                    sStr = sStr & "        SUM(K10) AS K10,"
                    sStr = sStr & "        SUM(J11) AS J11,"
                    sStr = sStr & "        SUM(K11) AS K11,"
                    
                    sStr = sStr & "        SUM(J12) AS J12,"
                    sStr = sStr & "        SUM(K12) AS K12,"
                    sStr = sStr & "        SUM(J13) AS J13,"
                    sStr = sStr & "        SUM(K13) AS K13,"
                    sStr = sStr & " '' AS K14,"
                    sStr = sStr & " '' AS J14,"
                    
                    sStr = sStr & "        SUM(J15) AS J15,"
                    sStr = sStr & "        SUM(K15) AS K15,"
                    sStr = sStr & "        SUM(J16) AS J16,"
                    sStr = sStr & "        SUM(K16) AS K16,"
                    sStr = sStr & "        SUM(J17) AS J17,"
                    sStr = sStr & "        SUM(K17) AS K17,"
                    sStr = sStr & "        SUM(J18) AS J18,"
                    sStr = sStr & "        SUM(K18) AS K18,"
                    
                    sStr = sStr & "        SUM(J19) AS J19,"
                    sStr = sStr & "        SUM(K19) AS K19,"
                    sStr = sStr & "        SUM(J20) AS J20,"
                    sStr = sStr & "        SUM(K20) AS K20,"
                    sStr = sStr & "        SUM(J21) AS J21,"
                    sStr = sStr & "        SUM(K21) AS K21"
            
            sStr = sStr & "    FROM ("
            '---------------------------------------------------------------------------- 전체학생 조회 START
            sStr = sStr & "          SELECT *"
            sStr = sStr & "            FROM CLSTD01TB"
            sStr = sStr & "           WHERE ACID = '" & Trim(basModule.SchCD) & "'"
            sStr = sStr & "             AND EXMID > ' ' "
            
    If Trim(Right(kaeyol, 30)) <> "ALL" Then
            sStr = sStr & "             AND KAEYOL = '" & Trim(Right(kaeyol, 30)) & "'"
    End If
    
    '<< 기간설정 >>
    If day1 <> "" Or day2 <> "" Then
        If Trim(day1) <> "" And Trim(day2) <> "" Then
            sStr = sStr & "             AND REGDATE BETWEEN TO_DATE('" & Trim(day1) & "000000', 'YYYYMMDDHH24MISS') "
            sStr = sStr & "                             AND TO_DATE('" & Trim(day2) & "235959', 'YYYYMMDDHH24MISS') "
        ElseIf Trim(day1) = "" And Trim(day2) <> "" Then
            sStr = sStr & "             AND REGDATE BETWEEN TO_DATE('19000101000000', 'YYYYMMDDHH24MISS') "
            sStr = sStr & "                             AND TO_DATE('" & Trim(day2) & "235959', 'YYYYMMDDHH24MISS') "
        ElseIf Trim(day1) <> "" And Trim(day2) = "" Then
            sStr = sStr & "             AND REGDATE BETWEEN TO_DATE('" & Trim(day1) & "000000', 'YYYYMMDDHH24MISS') "
            sStr = sStr & "                             AND TO_DATE('99991231235959', 'YYYYMMDDHH24MISS') "
        End If

    End If
            sStr = sStr & "             AND BIGO2 IS NULL "
            sStr = sStr & "          UNION ALL"
            '---------------------------------------------------------------------------- 전체학생 조회 END
            '---------------------------------------------------------------------------- 합격자 조회 START
            sStr = sStr & "          SELECT *"
            sStr = sStr & "            From CLSTD01TB"
            sStr = sStr & "           WHERE (PASS1 = '" & Trim(basModule.SchCD) & "'" & " OR"
            sStr = sStr & "                  PASS2 = '" & Trim(basModule.SchCD) & "'" & " OR"
            sStr = sStr & "                  PASS3 = '" & Trim(basModule.SchCD) & "'" & " OR"
            sStr = sStr & "                  PASS4 = '" & Trim(basModule.SchCD) & "'" & " )"
            sStr = sStr & "             AND EXMID > ' ' "
    If Trim(Right(kaeyol, 30)) <> "ALL" Then
            sStr = sStr & "             AND KAEYOL = '" & Trim(Right(kaeyol, 30)) & "'"
    End If
    
    '<< 기간설정 >>
    If day1 <> "" Or day2 <> "" Then
        If Trim(day1) <> "" And Trim(day2) <> "" Then
            sStr = sStr & "             AND REGDATE BETWEEN TO_DATE('" & Trim(day1) & "000000', 'YYYYMMDDHH24MISS') "
            sStr = sStr & "                             AND TO_DATE('" & Trim(day2) & "235959', 'YYYYMMDDHH24MISS') "
        ElseIf Trim(day1) = "" And Trim(day2) <> "" Then
            sStr = sStr & "             AND REGDATE BETWEEN TO_DATE('19000101000000', 'YYYYMMDDHH24MISS') "
            sStr = sStr & "                             AND TO_DATE('" & Trim(day2) & "235959', 'YYYYMMDDHH24MISS') "
        ElseIf Trim(day1) <> "" And Trim(day2) = "" Then
            sStr = sStr & "             AND REGDATE BETWEEN TO_DATE('" & Trim(day1) & "000000', 'YYYYMMDDHH24MISS') "
            sStr = sStr & "                             AND TO_DATE('99991231235959', 'YYYYMMDDHH24MISS') "
        End If

    End If
            sStr = sStr & "             AND BIGO2 IS NULL "
            
            sStr = sStr & "          ) A, "
            
            sStr = sStr & "               ("
        
                sStr = sStr & "         SELECT SCHNO,"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '37', DECODE(SUB_NUM,'X',0, SUB_NUM), 0)    AS J01,    /* 언어                  */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '37', DECODE(SUB_BAK,'X',0, SUB_BAK), 0)    AS K01,    /* 백분위  언어          */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '38', DECODE(SUB_NUM,'X',0, SUB_NUM), 0)    AS J02,    /* 수리가형              */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '38', DECODE(SUB_BAK,'X',0, SUB_BAK), 0)    AS K02,    /* 백분위  수리가형      */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '39', DECODE(SUB_NUM,'X',0, SUB_NUM), 0)    AS J03,    /* 외국어                */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '39', DECODE(SUB_BAK,'X',0, SUB_BAK), 0)    AS K03,    /* 백분위  외국어        */"
                
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(0) & "', DECODE(SUB_NUM,'X',0, SUB_NUM), '51', DECODE(SUB_NUM, 'X',0, SUB_NUM), 0) AS J04,    /* 사탐-" & constSatams(0) & "       , 과탐-물리1             */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(0) & "', DECODE(SUB_BAK,'X',0, SUB_BAK), '51', DECODE(SUB_BAK, 'X',0, SUB_BAK), 0) AS K04,    /* 백분위  사탐-" & constSatams(0) & "        , 과탐-물리1     */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(1) & "', DECODE(SUB_NUM,'X',0, SUB_NUM), '52', DECODE(SUB_NUM, 'X',0, SUB_NUM), 0) AS J05,    /* 사탐-" & constSatams(1) & "        , 과탐-화학1             */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(1) & "', DECODE(SUB_BAK,'X',0, SUB_BAK), '52', DECODE(SUB_BAK, 'X',0, SUB_BAK), 0) AS K05,    /* 백분위  사탐-" & constSatams(1) & "        , 과탐-화학1     */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(2) & "', DECODE(SUB_NUM,'X',0, SUB_NUM), '53', DECODE(SUB_NUM, 'X',0, SUB_NUM), 0) AS J06,    /* 사탐-" & constSatams(2) & "        , 과탐-생명과학1             */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(2) & "', DECODE(SUB_BAK,'X',0, SUB_BAK), '53', DECODE(SUB_BAK, 'X',0, SUB_BAK), 0) AS K06,    /* 백분위  사탐-" & constSatams(2) & "        , 과탐-생명과학1     */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(3) & "', DECODE(SUB_NUM,'X',0, SUB_NUM), '54', DECODE(SUB_NUM, 'X',0, SUB_NUM), 0) AS J07,    /* 사탐-" & constSatams(3) & "  , 과탐-지구과학1         */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(3) & "', DECODE(SUB_BAK,'X',0, SUB_BAK), '54', DECODE(SUB_BAK, 'X',0, SUB_BAK), 0) AS K07,    /* 백분위  사탐-" & constSatams(3) & "  , 과탐-지구과학1 */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(4) & "', DECODE(SUB_NUM,'X',0, SUB_NUM), '55', DECODE(SUB_NUM, 'X',0, SUB_NUM), 0) AS J08,    /* 사탐-" & constSatams(4) & "      , 과탐-물리2             */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(4) & "', DECODE(SUB_BAK,'X',0, SUB_BAK), '55', DECODE(SUB_BAK, 'X',0, SUB_BAK), 0) AS K08,    /* 백분위  사탐-" & constSatams(4) & "      , 과탐-물리2     */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(5) & "', DECODE(SUB_NUM,'X',0, SUB_NUM), '56', DECODE(SUB_NUM, 'X',0, SUB_NUM), 0) AS J09,    /* 사탐-" & constSatams(5) & "    , 과탐-화학2             */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(5) & "', DECODE(SUB_BAK,'X',0, SUB_BAK), '56', DECODE(SUB_BAK, 'X',0, SUB_BAK), 0) AS K09,    /* 백분위  사탐-" & constSatams(5) & "    , 과탐-화학2     */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(6) & "', DECODE(SUB_NUM,'X',0, SUB_NUM), '57', DECODE(SUB_NUM, 'X',0, SUB_NUM), 0) AS J10,      /* 사탐-" & constSatams(6) & "    , 과탐-생명과학2           */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(6) & "', DECODE(SUB_BAK,'X',0, SUB_BAK), '57', DECODE(SUB_BAK, 'X',0, SUB_BAK), 0) AS K10,      /* 백분위 사탐-" & constSatams(6) & "    , 과탐-생명과학2    */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(7) & "', DECODE(SUB_NUM,'X',0, SUB_NUM), '58', DECODE(SUB_NUM, 'X',0, SUB_NUM), 0) AS J11,    /* 사탐-" & constSatams(7) & "        , 과탐-지구과학2         */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(7) & "', DECODE(SUB_BAK,'X',0, SUB_BAK), '58', DECODE(SUB_BAK, 'X',0, SUB_BAK), 0) AS K11,    /* 백분위  사탐-" & constSatams(7) & "        , 과탐-지구과학2 */"
                
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(8) & "', DECODE(SUB_NUM,'X',0, SUB_NUM), 0)    AS J12,    /* 사탐-" & constSatams(8) & "         */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(8) & "', DECODE(SUB_BAK,'X',0, SUB_BAK), 0)    AS K12,    /* 백분위  사탐-" & constSatams(8) & " */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(9) & "', DECODE(SUB_NUM,'X',0, SUB_NUM), 0)    AS J13,    /* 사탐-" & constSatams(9) & "         */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '" & constSatamCodes(9) & "', DECODE(SUB_BAK,'X',0, SUB_BAK), 0)    AS K13,    /* 백분위  사탐-" & constSatams(9) & " */"
                sStr = sStr & " '' AS K14,"
                sStr = sStr & " '' AS J14,"
                
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '31', DECODE(SUB_NUM,'X',0, SUB_NUM), '81', DECODE(SUB_NUM, 'X',0, SUB_NUM), 0) AS J15,    /* 독어             , 미적분                 */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '31', DECODE(SUB_BAK,'X',0, SUB_BAK), '81', DECODE(SUB_BAK, 'X',0, SUB_BAK), 0) AS K15,    /* 백분위  독어             , 미적분         */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '32', DECODE(SUB_NUM,'X',0, SUB_NUM), '82', DECODE(SUB_NUM, 'X',0, SUB_NUM), 0) AS J16,    /* 일어             , 이산수학               */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '32', DECODE(SUB_BAK,'X',0, SUB_BAK), '82', DECODE(SUB_BAK, 'X',0, SUB_BAK), 0) AS K16,    /* 백분위  일어             , 이산수학       */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '33', DECODE(SUB_NUM,'X',0, SUB_NUM), '83', DECODE(SUB_NUM, 'X',0, SUB_NUM), 0) AS J17,    /* 에스파냐         , 확률통계               */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '33', DECODE(SUB_BAK,'X',0, SUB_BAK), '83', DECODE(SUB_BAK, 'X',0, SUB_BAK), 0) AS K17,    /* 백분위  에스파냐         , 확률통계       */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '34', DECODE(SUB_NUM,'X',0, SUB_NUM), '43', DECODE(SUB_NUM, 'X',0, SUB_NUM), 0) AS J18,    /* 불어             , 수리나형               */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '34', DECODE(SUB_BAK,'X',0, SUB_BAK), '43', DECODE(SUB_BAK, 'X',0, SUB_BAK), 0) AS K18,    /* 백분위  불어             , 수리나형       */"
                
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '35', DECODE(SUB_NUM,'X',0, SUB_NUM), 0)    AS J19,    /* 중국어                */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '35', DECODE(SUB_BAK,'X',0, SUB_BAK), 0)    AS K19,    /* 백분위  중국어        */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '36', DECODE(SUB_NUM,'X',0, SUB_NUM), 0)    AS J20,    /* 한문                  */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '36', DECODE(SUB_BAK,'X',0, SUB_BAK), 0)    AS K20,    /* 백분위  한문          */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '42', DECODE(SUB_NUM,'X',0, SUB_NUM), 0)    AS J21,    /* 아랍어                */"
                sStr = sStr & "                DECODE(TRIM(SUB_ID), '42', DECODE(SUB_BAK,'X',0, SUB_BAK), 0)    AS K21     /* 백분위  아랍어        */"
                sStr = sStr & "           FROM CLSTD03TB"
        
        sStr = sStr & "                ) B"
        sStr = sStr & "        WHERE A.SCHNO = B.SCHNO(+)"
            
            sStr = sStr & "   GROUP BY A.SCHNO"
            '---------------------------------------------------------------------------- 합격자 조회 END
    
    sStr = sStr & "    ) "
    sStr = sStr & " ORDER BY EXMID "
    
    Get_StdExcuteSqlToExcel = sStr
End Function


'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'유틸
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Public Function Cbo_Val(ByRef cboControl As Object, ByVal length As Long) As String
    Cbo_Val = Trim(Right(cboControl.Text, 30))
End Function


Public Function Get_IndexByChk(ByRef Control As Object) As Long
    
    Dim i As Long
    Dim selIndex As Long
    
    selIndex = -1
    For i = 0 To Control.count - 1
        If Control(i).value = True Then: selIndex = i
    Next
    
    Get_IndexByChk = selIndex
End Function
