Attribute VB_Name = "basGwamok"
Option Explicit


Public SATAM_COUNT
Public GWATAM_COUNT
Public ENG2_COUNT
Public MATH2_COUNT
Public TAMGOO_COUNT
Public CLINIC_L_COUNT
Public CLINIC_M_COUNT
Public CLINIC_E_COUNT

'>> 등급 이다..
Public Const CLINIC_L_CLASS = 101
Public Const CLINIC_M_CLASS = 111
Public Const CLINIC_E_CLASS = 121
'>> 아래 등급은 잘못설정했다. 원래 21,31,51등으로 되었어야했는데.. 귀차나.
Public Const SATAM_CLASS = 20
Public Const GWATAM_CLASS = 30
Public Const ENG2_CLASS = 50
Public Const MATH2_CLASS = 80
Public Const TAMGOO_CLASS = 90

Public constSatams() As String
Public constGwatams() As String
Public constEng2s() As String
Public constMaths() As String
Public constTamgoos() As String

Public constSatamCodes() As String
Public constGwatamCodes() As String
Public constEng2Codes() As String
Public constMathCodes() As String
Public constTamgooCodes() As String


Public g_sClinic_LCodes() As String
Public g_sClinic_MCodes() As String
Public g_sClinic_ECodes() As String

Public g_sClinic_Ls() As String
Public g_sClinic_Ms() As String
Public g_sClinic_Es() As String

Public Const TGANG_CODE = "95"





'목적 : 과목과 코드가 변경이 이뤄졌을경우. basGwamok파일만 변경하면 되도록.
'>>     *********중요 ************** 1. frm파일에서 과목명들을 모두 걷어낸다.
'>>     100%걷어내지못한다. (에파 ->에스파냐, 세사->세계사)


'상황 : 자료관리는 이파일기준으로 DB처럼 사용하면 되나
'>>     소스코드를 보면 uExcel_StdData.SATAM3(엑셀파일시트의SATAM3컬럼)과같이 하드코딩되어잇는부분이 있기떄문에
'>>     어쩔수없이 과목이 사용되는부분 전체를 잘살펴보아야 하지만
'>>     수정해야하는부분을 많이 줄여주었다.

'>>     프로그램 소스코드의 복잡성이 상당하므로.. private 으로 되어있는것을 public으로 바꾸어서 무분별하게
'>>     다른파일에서 사용하지 않도록한다.
'>>     전체리펙토링을 하지 않는이상... 더이상 소스복잡해지지 않도록 조심하자.

'>>     학원별로 과목이 틀리기 때문에 학원 코드에 따른 배열들을 세팅해주어야한다.

'DB대신에 배열을 사용.
'언제든지 DB를 이용하여 과목정보를 이용할수있또록
'과목변경하면서 소스구조도 변경함.



'>>     다필요없고. 어줍잖게 부분적으로 수정하다가 더 복잡해졌다. (코드의 일관성 파괴)
'>>     노가다 할건 하고. 그냥 최대한 데이터 중심으로만 바꾸려고 노력했다.


'마강
Private Function setConstant_M()
    
End Function
'마송
Private Function setConstant_P()
    
End Function
'부산
Private Function setConstant_B()
   
End Function
'송파
Private Function setConstant_S()
    '>>클리닉
    '국어
    g_sClinic_Ls(0) = "(심화)어법&문학 개념어"
    g_sClinic_Ls(1) = "(심화)고난도 취약유형 연습"
    g_sClinic_Ls(2) = "(기본)어법&문학 개념어"
    g_sClinic_Ls(3) = "(심화)비문학 취약유형 연습"
    
    g_sClinic_LCodes(0) = "101"
    g_sClinic_LCodes(1) = "102"
    g_sClinic_LCodes(2) = "103"
    g_sClinic_LCodes(3) = "104"
    
    '수학
    g_sClinic_Ms(0) = "(기본)수능에 꼭 필요한 고1 수학"
    g_sClinic_Ms(1) = "(심화)도형과 함수"
    g_sClinic_Ms(2) = "(기본)함수의 극한&미분"
    g_sClinic_Ms(3) = "(심화)공간도형과 벡터"
    
    g_sClinic_MCodes(0) = "111"
    g_sClinic_MCodes(1) = "112"
    g_sClinic_MCodes(2) = "113"
    g_sClinic_MCodes(3) = "114"
    
    '영어
    g_sClinic_Es(0) = "(기본)핵심문법과 구문"
    g_sClinic_Es(1) = "(심화)취약유형 독해연습"
    g_sClinic_Es(2) = "(기본)기초문법 및 구문"
    g_sClinic_Es(3) = "(심화)취약유형 독해연습"
    
    g_sClinic_ECodes(0) = "121"
    g_sClinic_ECodes(1) = "122"
    g_sClinic_ECodes(2) = "123"
    g_sClinic_ECodes(3) = "124"
End Function
'노량진
Private Function setConstant_N()
    '>>클리닉
    '국어
    g_sClinic_Ls(0) = "(심화)문학 독해"
    g_sClinic_Ls(1) = "(심화)비문학 독해"
    g_sClinic_Ls(2) = "(기본)문학 독해의 정석"
    g_sClinic_Ls(3) = "(기본)비문학 독해의 정석"
    
    g_sClinic_LCodes(0) = "101"
    g_sClinic_LCodes(1) = "102"
    g_sClinic_LCodes(2) = "103"
    g_sClinic_LCodes(3) = "104"
    
    '수학
    g_sClinic_Ms(0) = "(심화)도형과 함수"
    g_sClinic_Ms(1) = "(기본)수능에 꼭 필요한 고1 수학"
    g_sClinic_Ms(2) = "(심화)공간도형과 벡터"
    g_sClinic_Ms(3) = "(기본)수학1 핵심유형 총정리"
    
    g_sClinic_MCodes(0) = "111"
    g_sClinic_MCodes(1) = "112"
    g_sClinic_MCodes(2) = "113"
    g_sClinic_MCodes(3) = "114"
    
    '영어
    g_sClinic_Es(0) = "(심화)고난도 구문&독해"
    g_sClinic_Es(1) = "(심화)어법"
    g_sClinic_Es(2) = "(기본)독해 문제 풀이법"
    g_sClinic_Es(3) = "(기본)어법"
    
    g_sClinic_ECodes(0) = "121"
    g_sClinic_ECodes(1) = "122"
    g_sClinic_ECodes(2) = "123"
    g_sClinic_ECodes(3) = "124"
End Function
'강남
Private Function setConstant_K()

End Function
'야법
Private Function setConstant_Q()

End Function

'주법
Private Function setConstant_W()

End Function

'양재
Private Function setConstant_J()
    
'    SATAM_COUNT = 11
'    GWATAM_COUNT = 9
'
'
'    ReDim Preserve constSatams(SATAM_COUNT - 1)
'    ReDim Preserve constSatamCodes(SATAM_COUNT - 1)
'    ReDim Preserve constGwatams(GWATAM_COUNT - 1)
'    ReDim Preserve constGwatamCodes(GWATAM_COUNT - 1)
'
'
'
'    constSatams(10) = "특강"
'    constSatamCodes(10) = "95"
'    constGwatams(8) = "특강"
'    constGwatamCodes(8) = "95"
    
End Function



'DB배열
Function setConstant()

    SATAM_COUNT = 10
    GWATAM_COUNT = 8
    ENG2_COUNT = 12
    MATH2_COUNT = 4
    TAMGOO_COUNT = 3
    CLINIC_L_COUNT = 4
    CLINIC_M_COUNT = 4
    CLINIC_E_COUNT = 4
    
    
    ReDim constSatams(SATAM_COUNT - 1)
    ReDim constGwatams(GWATAM_COUNT - 1)
    ReDim constEng2s(ENG2_COUNT - 1)
    ReDim constMaths(MATH2_COUNT - 1)
    ReDim constTamgoos(TAMGOO_COUNT - 1)
    
    ReDim constSatamCodes(SATAM_COUNT - 1)
    ReDim constGwatamCodes(GWATAM_COUNT - 1)
    ReDim constEng2Codes(ENG2_COUNT - 1)
    ReDim constMathCodes(MATH2_COUNT - 1)
    ReDim constTamgooCodes(TAMGOO_COUNT - 1)

    
    ReDim g_sClinic_LCodes(CLINIC_L_COUNT - 1)
    ReDim g_sClinic_MCodes(CLINIC_M_COUNT - 1)
    ReDim g_sClinic_ECodes(CLINIC_E_COUNT - 1)
    
    ReDim g_sClinic_Ls(CLINIC_E_COUNT - 1)
    ReDim g_sClinic_Ms(CLINIC_E_COUNT - 1)
    ReDim g_sClinic_Es(CLINIC_E_COUNT - 1)



    ' 사회탐구
    constSatams(0) = "한국사"
    constSatams(1) = "세계사"
    constSatams(2) = "동아시아사"
    constSatams(3) = "한국지리"
    constSatams(4) = "세계지리"
    constSatams(5) = "생활과윤리"
    constSatams(6) = "윤리사상"
    constSatams(7) = "법과정치"
    constSatams(8) = "경제"
    constSatams(9) = "사회문화"
    
    
    constSatamCodes(0) = "21"
    constSatamCodes(1) = "22"
    constSatamCodes(2) = "23"
    constSatamCodes(3) = "24"
    constSatamCodes(4) = "25"
    constSatamCodes(5) = "26"
    constSatamCodes(6) = "27"
    constSatamCodes(7) = "28"
    constSatamCodes(8) = "29"
    constSatamCodes(9) = "30"
    
    '과학탐구
    constGwatams(0) = "물리1"
    constGwatams(1) = "화학1"
    constGwatams(2) = "생명과학1"
    constGwatams(3) = "지구과학1"
    constGwatams(4) = "물리2"
    constGwatams(5) = "화학2"
    constGwatams(6) = "생명과학2"
    constGwatams(7) = "지구과학2"
    
    constGwatamCodes(0) = "51"
    constGwatamCodes(1) = "52"
    constGwatamCodes(2) = "53"
    constGwatamCodes(3) = "54"
    constGwatamCodes(4) = "55"
    constGwatamCodes(5) = "56"
    constGwatamCodes(6) = "57"
    constGwatamCodes(7) = "58"
    
    ' 제2외국어
    constEng2s(0) = "독어"
    constEng2s(1) = "일어"
    constEng2s(2) = "에스파냐어"
    constEng2s(3) = "불어"
    constEng2s(4) = "중국어"
    constEng2s(5) = "한문"
    constEng2s(6) = "언어"
    constEng2s(7) = "수리"
    constEng2s(8) = "영어"
    constEng2s(9) = "세계사"
    constEng2s(10) = "세계지리"
    constEng2s(11) = "아랍어"
    
     
    constEng2Codes(0) = "31"
    constEng2Codes(1) = "32"
    constEng2Codes(2) = "33"
    constEng2Codes(3) = "34"
    constEng2Codes(4) = "35"
    constEng2Codes(5) = "36"
    constEng2Codes(6) = "37"
    constEng2Codes(7) = "38"
    constEng2Codes(8) = "39"
    constEng2Codes(9) = "40"
    constEng2Codes(10) = "41"
    constEng2Codes(11) = "42"
    
    '탐구선택
    constMaths(0) = "미적분"
    constMaths(1) = "이산수학"
    constMaths(2) = "확률통계"
    constMaths(3) = "수리나형"
    
    constMathCodes(0) = "81"
    constMathCodes(1) = "82"
    constMathCodes(2) = "83"
    constMathCodes(3) = "84"
    
    
    '탐구선택
    constTamgoos(0) = "언어"
    constTamgoos(1) = "수리"
    constTamgoos(2) = "외국어"
    
    constTamgooCodes(0) = "91"
    constTamgooCodes(1) = "92"
    constTamgooCodes(2) = "93"
    
    

    
    
    '위의 코드들이 기본이고 나머지 바뀌는것들은 아래의 학원별 설정 함수에서 세팅함.
    Select Case basModule.SchCD
        Case "M":   Call setConstant_M
        Case "P":   Call setConstant_P
        Case "B":   Call setConstant_B
        Case "J":   Call setConstant_J
        Case "S":   Call setConstant_S
        Case "N":   Call setConstant_N
        Case "K":   Call setConstant_K
        Case "Q":   Call setConstant_Q
        Case "W":   Call setConstant_W
    End Select
    
    
End Function

'과목명 - > 과목코드
Function Get_GwaMokCodeByName(subject As String) As String
    
    Dim i As Integer
    
    For i = 0 To SATAM_COUNT - 1
        If constSatams(i) = subject Then
            Get_GwaMokCodeByName = constSatamCodes(i)
            Exit Function
        End If
    Next i
    
    For i = 0 To ENG2_COUNT - 1
        If constEng2Codes(i) = subject Then
            Get_GwaMokCodeByName = constEng2Codes(i)
            Exit Function
        End If
    Next i
    
    For i = 0 To GWATAM_COUNT - 1
        If constGwatams(i) = subject Then
            Get_GwaMokCodeByName = constGwatamCodes(i)
            Exit Function
        End If
    Next i
    
    For i = 0 To MATH2_COUNT - 1
        If constMaths(i) = subject Then
            Get_GwaMokCodeByName = constMathCodes(i)
            Exit Function
        End If
    Next i
    For i = 0 To TAMGOO_COUNT - 1
        If constTamgoos(i) = subject Then
            Get_GwaMokCodeByName = constTamgooCodes(i)
            Exit Function
        End If
    Next i

End Function



'필드에 해당하는 과목코드목록들을 리턴해준다.
'SEL1~5는 DB 컬럼값.
Function Get_GwaMokCodes(fieldName As String) As String()

    Select Case fieldName
        Case "SEL1": Get_GwaMokCodes = constSatamCodes  '사탐
        Case "SEL2": Get_GwaMokCodes = constEng2Codes   '제2외국어
        Case "SEL3": Get_GwaMokCodes = constGwatamCodes '과탐
        Case "SEL4": Get_GwaMokCodes = constMathCodes   '수리
        Case "SEL5": Get_GwaMokCodes = constTamgooCodes '탐구
        
    End Select

End Function

'필드에 해당하는 과목코드목록들을 리턴해준다.
Function Get_GwaMokNames(fieldName As String) As String()

    Select Case fieldName
        Case "SEL1": Get_GwaMokNames = constSatams
        Case "SEL2": Get_GwaMokNames = constEng2s
        Case "SEL3": Get_GwaMokNames = constGwatams
        Case "SEL4": Get_GwaMokNames = constMaths
        Case "SEL5": Get_GwaMokNames = constTamgoos
        
    End Select

End Function

' 범위 안에 과목들을 리턴해준다.
Function Get_StrGwaMokRange(codes As String, rangeStart As Long, rangeEnd As Long) As String

    'codes를 split
    Dim arrTmp() As String
    Dim count As Long
    Dim i As Long
    Dim code As String
    
    Dim sReturnVal As String
    
    
    sReturnVal = ""
    
    arrTmp = Split(Trim(codes), "|", -1, vbTextCompare)
    
    count = UBound(arrTmp)
    
    For i = 0 To count - 1
    
        code = arrTmp(i)
        If code >= rangeStart And code <= rangeEnd Then
        
            sReturnVal = sReturnVal & Get_StrGwaMokByCode(code)
        
        End If
        
    Next i
    
    Get_StrGwaMokRange = sReturnVal

End Function

'과목 코드 ->  과목명
Function Get_StrGwaMokByCode(gwamokCode As String) As String

   Dim i As Integer
    
    For i = 0 To SATAM_COUNT - 1
        If constSatamCodes(i) = gwamokCode Then
            Get_StrGwaMokByCode = constSatams(i)
            Exit Function
        End If
    Next i
    
    For i = 0 To ENG2_COUNT - 1
        If constEng2Codes(i) = gwamokCode Then
            Get_StrGwaMokByCode = constEng2s(i)
            Exit Function
        End If
    Next i
    
    For i = 0 To GWATAM_COUNT - 1
        If constGwatamCodes(i) = gwamokCode Then
            Get_StrGwaMokByCode = constGwatams(i)
            Exit Function
        End If
    Next i
    
    For i = 0 To MATH2_COUNT - 1
        If constMathCodes(i) = gwamokCode Then
            Get_StrGwaMokByCode = constMaths(i)
            Exit Function
        End If
    Next i
    
    For i = 0 To TAMGOO_COUNT - 1
        If constTamgooCodes(i) = gwamokCode Then
            Get_StrGwaMokByCode = constTamgoos(i)
            Exit Function
        End If
    Next i
    
    For i = 0 To CLINIC_L_COUNT - 1
        If g_sClinic_LCodes(i) = gwamokCode Then
            Get_StrGwaMokByCode = g_sClinic_Ls(i)
            Exit Function
        End If
    Next i
    
    For i = 0 To CLINIC_M_COUNT - 1
        If g_sClinic_MCodes(i) = gwamokCode Then
            Get_StrGwaMokByCode = g_sClinic_Ms(i)
            Exit Function
        End If
    Next i
    
    For i = 0 To CLINIC_E_COUNT - 1
        If g_sClinic_ECodes(i) = gwamokCode Then
            Get_StrGwaMokByCode = g_sClinic_Es(i)
            Exit Function
        End If
    Next i
    
End Function

'스프레드에서 과목11번째 컬럼을 삭제한다.
'디자인 모드에서 과목11 컬럼을 쉽게 삭제하는방법을 모르겠다.
Public Sub Spread_DelCel(ss As Control, colNum As Long)
    ss.Col = colNum
    ss.Action = 4
End Sub
