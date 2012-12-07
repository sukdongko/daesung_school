Attribute VB_Name = "basDataBase"
'################################################################################################################
'   시 스 템  명 : 대성학원 입학사정, 반배정 & 시간표 프로그램
' 5  서브시스템명 :
'   모   듈   명 : BASDATABASE
'   모 듈  목 적 : 데이터베이스 접속
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

'>> ODBC API 설정
    Private Declare Function SQLConfigDataSource Lib "ODBCCP32.DLL" _
                   (ByVal hwndParent As Long, _
                    ByVal fRequest As Long, _
                    ByVal lpszDriver As String, _
                    ByVal lpszAttributes As String) As Long
                    
    Private Const ODBC_ADD_DSN_USER = 1     ' Add data source
    Private Const ODBC_ADD_DSN_SYS = 4      ' Add data source
    Private Const ODBC_CONFIG_DSN = 2       ' Configure (edit) data source
    Private Const ODBC_REMOVE_DSN = 3       ' Remove data source
    Private Const vbAPINull As Long = 0&    ' NULL Pointer

    Private Const ODBCDrv = "Oracle ODBC Driver"
    Private Const ODBCDrv92 = "Oracle in OraHome92"             ' 노량진
    Private Const ODBCDrv11g = "{Microsoft ODBC for Oracle}"    ' 11g
    
    Private OracleVer       As String                           ' 오라클 버젼
    
    
    Public DBConn As ADODB.Connection
    
    
'>> REGISTRY 에 관여하는 DECLARE STATEMENT
    Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
    Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
    Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
    
    Public Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long
    Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
    
    Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
    Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long

    ' ROOT 키를 처리
    Public Const ERROR_SUCCESS = 0&
    
    Public Const REG_OPTION_BACKUP_RESTORE = 4                      '백업이나 복구를 위해 필요한 엑세스로 열린다.
    Public Const REG_OPTION_VOLATILE = 1                            '휘발성 데이터가 아니므로 시스템 재시동시 손실되지 않는다.
    Public Const REG_OPTION_NON_VOLATILE = &O0                      '데이터는 메모리에만 쓰여지고, 디스크 상에는 기록되지 않는다.
    
    Public Const REG_BINARY = 3                                     '이진데이터
    Public Const REG_DWORD = 4                                      '32 BIT NUMBER
    Public Const REG_DWORD_LITTLE_ENDIAN = 4                        '32 BIT NUMBER 최상위 바이트가 하위바이트
    Public Const REG_DWORD_BIG_ENDIAN = 5                           '32 BIT NUMBER 최상위 바이트가 상위바이트
    Public Const REG_EXPAND_SZ = 2                                  '환경변수에 대해 확장 되지 않은 참조형태의 문자열 (EX : %PATH%)
    
    Public Const REG_LINK = 6                                       '유니코드 심볼릭 연결
    Public Const REG_MULTI_SZ = 7                                   '널로 끝나는 문자열의 리스트
    Public Const REG_NONE = 0                                       '정의되지 않은 유형
    Public Const REG_RESOURCE_LIST = 8                              '디바이스 드라이버 자원 목록
    Public Const REG_SZ = 1                                         'NULL로 끝나는 문자열
    
    Public Const HKEY_CLASSES_ROOT = &H80000000
    Public Const HKEY_CURRENT_USER = &H80000001
    Public Const HKEY_LOCAL_MACHINE = &H80000002
    Public Const HKEY_USERS = &H80000003
    Public Const HKEY_CURRENT_CONFIG = &H80000005
    Public Const HKEY_DYN_DATA = &H80000006
    
    Public Const STANDARD_RIGHT_ALL = &H1F0000
    Public Const KEY_QUERY_VALUE = &H1
    Public Const KEY_SET_VALUE = &H2
    
    Public Const KEY_ENUMERATE_SUB_KEYS = &H8
    Public Const KEY_NOTIFY = &H10
    
    Public Const KEY_CREATE_LINK = &H20
    Public Const KEY_CREATE_SUB_KEY = &H4
    
    Public Const SYNCHRONIZE = &H100000
    Public Const KEY_ALL_ACCESS = ( _
                    STANDARD_RIGHT_ALL Or _
                    KEY_QUERY_VALUE Or _
                    KEY_SET_VALUE Or _
                    KEY_CREATE_SUB_KEY Or _
                    KEY_ENUMERATE_SUB_KEYS Or _
                    KEY_NOTIFY Or _
                    KEY_CREATE_LINK _
                    ) And _
                    (Not SYNCHRONIZE)
    
    Public Type SECURITY_ATTRIBUTES
        nLength                 As Long
        lpSecurityDescriptor    As Long
        bInHeritHandle          As Long
    End Type
    
    Public Type ACL
        AclRevision             As Byte
        Sbz1                    As Byte
        AclSize                 As Integer
        AceCount                As Integer
        Sbz2                    As Integer
    End Type
    
    Public Type SECURITY_DESCRIPTOR
        Revision                As Byte
        Sbz1                    As Byte
        Control                 As Long
        Owner                   As Long
        Group                   As Long
        Sacl                    As ACL
        Dacl                    As ACL
    End Type
    
'>> SERVER 설정
    Public Const PORT = "15800"
    Public Const Dev_LoginHost = "dmdev.mimacstudy.com"
    Public Const Mimac_LoginHost = "ms.mimacstudy.com"
    'Public Const PassWord = "sybaQ#12"

    Public Const hKey = HKEY_LOCAL_MACHINE
'    Public Const SubKey = "SOFTWARE\ORACLE\KEY_OraClient11g_home1\"
    
    
    Private Oracle_Pass         As String                   '<< oracle pass
    
    Public Const TNS_Path1 = "C:\oracle\instantclient_11_2_0_3\network\ADMIN\tnsnames.ora"
'    Public Const TNS_Path2 = "C:\ORACLE\instantclient_11_2\network\ADMIN\tnsnames.ora"
'    Public Const TNS_Path3 = "C:\oracle\product\11.2.0\client_1\Network\Admin\tnsnames.ora"
'    Public Const TNS_Path4 = "C:\oracle\ora81\network\ADMIN\tnsnames.ora"
'    Public Const TNS_Path5 = "C:\oracle\ora92\network\ADMIN\tnsnames.ora"
    
    

Public Function DataBase_Connection() As Boolean

    Dim bSuccess As Boolean
    Dim sTns As String
    Dim DB_Name As String
    Dim tns_Path As String
    
    On Error GoTo Error1                     'error 처리
    
    '>>>>>>>>>> tnsnames.ora파일 경로 가져오기
    tns_Path = Get_TNSNames_Path()
    If "" = tns_Path Then
        DataBase_Connection = False
        Exit Function
    End If
    
    
    '>>>>>>>>>>> tnsnames.ora파일에 DB접속정보 세팅.
    'tnsname.ora파일에 추가할 DB접속정보
    Select Case UCase(Trim(basModule.connDB))
        Case "MIMAC"
            sTns = "MI2_CLASS= (DESCRIPTION = (ADDRESS_LIST = (ADDRESS = (PROTOCOL = TCP)(HOST = ms.mimacstudy.com)(PORT = 30022)))(CONNECT_DATA = (SERVICE_NAME = DS)(INSTANCE_NAME = DS2)))"
            DB_Name = "MI2_CLASS"
        Case Else
            sTns = "DMDB =(DESCRIPTION =(ADDRESS = (PROTOCOL = TCP)(HOST = 172.16.0.5)(PORT = 1521))    (CONNECT_DATA =      (SERVER = DEDICATED)      (SERVICE_NAME = dm)    )  )"
            DB_Name = "DMDB"
    End Select
    
    bSuccess = writeOraTnsName(tns_Path, sTns, DB_Name)
    
    
    
    '>>>>>>>>>> DB 접속
    If bSuccess Then
        DataBase_Connection = connectionOledb()           '<< 최종 DB접속
    Else
        MsgBox "tnsnames.ora파일에 DB접속정보 세팅중에 에러발생", vbCritical + vbOKOnly, "데이터 접속"
    End If
    
    DataBase_Connection = True
    Exit Function
    
    
Error1:
    DataBase_Connection = False
    MsgBox "DatabaseConnection Error", "데이터베이스 접속"

End Function



Public Function connectionOledb() As Boolean
    
    On Error GoTo ErrorADODB                     'error 처리
    
    Dim strDB As String
    
    strDB = Chr(13) & Chr(10)
    
    Set DBConn = New ADODB.Connection
    'READGAME  READGAME/eoqkrskfk7
    'MSDAORA.1.1
    'OraOLEDB.Oracle
    Select Case UCase(Trim(basModule.connDB))
            Case "MIMAC"
               strDB = "Provider    =OraOLEDB.Oracle;" & _
                        "Data Source =MI2_CLASS;" & _
                        "User Id     =DSHW;" & _
                        "Password    =sybaQ#12;"
            Case Else
                strDB = "Provider    =OraOLEDB.Oracle;" & _
                        "Data Source =DMDB;" & _
                        "User Id     =DSHW;" & _
                        "Password    =sybaQ#12;"
        End Select
        
    ' Data Source = 접속할 데이터베이스
    DBConn.ConnectionString = strDB       '데이터베이스와 연결을 시도합니다.
    DBConn.ConnectionTimeout = 5          '제한 시간내에 연결이 되지 않으면 자동으로 끊습니다.
    'DB.Properties("Prompt") = adPromptNever   '이것은 ADO에서 기본 프롬프트 모드입니다.
    'DB.CursorLocation = adUseClient           '커서위치를 Client 쪽에 넣습니다.
        
    DBConn.Open                                   '데이터베이스를 엽니다.
    'MsgBox "연결 성공"
    DoEvents
'    Do While DB.State And adStateConnecting
'        DoEvents
'    Loop

    connectionOledb = True

    Exit Function
    
ErrorADODB:
    MsgBox "connectionOledb시 에러가 발생하였습니다." & vbCrLf & _
           Trim(CStr(Err.Number)) & ":" & Err.Description, vbCritical + vbOKOnly, "학생조회"

    'End
End Function





'>> 우선순위 1. DAESUNG.INI파일에 PATH_ORACLE_TNS
'>> 우선순위 2. 환경변수 Path   (레지스트리에서 읽음)
'>> TNS경로에 tnsnames.ora파일이 없을경우 생성   (\network\admin폴더가 없을경우 실패)
Public Function Get_TNSNames_Path() As String

    Dim tns_Path As String
    
    '>>>>>>>>> 오라클 경로 가져오기
    ' 환경변수 Path의 레지스트리값을 읽어와서 오라클경로를 가져옴.
    tns_Path = Get_TnsFile_Path_Registry()
    
    
    '>>>>>>>>>> DAESUNG.INI파일에서 오라클 경로 가져오기
    ' DAESUNG.INI파일에서 tns경로(PATH_ORACLE_TNS)가 있으면 그걸로 하고 없으면. Path에 잡혀있는걸로 한다.
    Dim sData               As String * 255
    Dim sTmp                As String
    
    sData = ""
    Call basModule.GetPrivateProfileString("SCHOOL", "PATH_ORACLE_TNS", "", sData, 255, App.Path & "\DAESUNG.INI")             '>> 오라클 인스턴스 경로
    sTmp = Trim(Replace(sData, Chr(0), "", 1, -1, vbTextCompare))
    If "" <> sTmp Then
        tns_Path = sTmp
    End If
    
    If "" = tns_Path Then
        MsgBox "오라클 경로설정에 문제가 있습니다. " & Chr(13) & "직접 INI파일에 PATH_ORACLE_TNS를 추가할것을 권장합니다.", vbCritical + vbOKOnly, "TNS경로 가져오기"
        Get_TNSNames_Path = ""
        Exit Function
    End If
    
    
    '>>>>>>>>>> tnsnames.ora파일 생성
    ' tns_Path경로에 tnsnames.ora파일이 없을경우 생성
    ' \network\admin 폴더가 없을경우 파일생성 실패
    If "" = Dir(tns_Path) Then
    
        If False = Create_Tnsnames(tns_Path) Then
            Get_TNSNames_Path = ""
            MsgBox "tnsnames.ora파일 생성실패", "Create_Tnsnames"
            Exit Function
        End If
    End If
    
    Get_TNSNames_Path = tns_Path
End Function



'RETURN VALLUE : 찾을경우 경로 , 못찾으면 ""
'tnsnames.ora파일 경로 가져오기 : 환경변수 PATH의 ORACLE 경로(레지스트리에서 읽어옴)를 참조
Private Function Get_TnsFile_Path_Registry() As String

    Dim RetStr      As String
    Dim hSubKey     As Long
    Dim Rtn         As Long
    Dim strLength   As Long
    Dim dType       As Long
    Dim sReturn     As String
    Dim bFindPath   As Boolean
    
    
    '>> Registry Open
    Rtn = RegOpenKeyEx(hKey, "SYSTEM\CurrentControlSet\Control\Session Manager\Environment\", 0, KEY_ALL_ACCESS, hSubKey)
    
    '하위 키값을 얻는다.
    strLength = 256
    RetStr = String(strLength, 0)

    '>> oracle path
    Rtn = RegQueryValueEx(hSubKey, "Path", 0, dType, ByVal RetStr, strLength)
    
    
    '>> Path키가 존재할경우
    bFindPath = False
    If Rtn = ERROR_SUCCESS Then
    
        '>> path안의 문자열들은 ";"로 구분되어있음
        Dim strPaths
        Dim Path
        
        sReturn = Left(RetStr, strLength - 1)   '뒤에 따라오는 문자열을 제거한다
        strPaths = Split(sReturn, ";")
        
        'PATH안의 문자들중에 oracle경로가 있는지 확인. '(network\admin 폴더가 있으면 OK)
        For Each Path In strPaths
            
            If "" <> Dir(Path & "\" & "network\admin", vbDirectory) Then
                '경로가 있음.
                sReturn = Path & "\network\admin\tnsnames.ora"
                bFindPath = True
                Exit For
            End If
        Next
        
        If bFindPath = False Then
            'MsgBox "환경변수 PATH에서 PATH\network\admin 폴더가 존재 하지 않습니다"
         End If
    Else
        'MsgBox "레지스트리값 읽어오기 실패 " & vbCrLf & "rtn : " & Rtn & "  dType :" & dType & Left(RetStr, strLength - 1)
    End If
    
    Rtn = RegCloseKey(hSubKey)
    
    If False = bFindPath Then
        sReturn = ""
    End If

    Get_TnsFile_Path_Registry = sReturn
    
End Function


'>>>>>>>>>>>> tnsnames.ora파일이 없을경우 파일 생성
Private Function Create_Tnsnames(tns_Path As String) As Boolean

    Dim strFolder
    Dim strFile
    
    strFile = Right(tns_Path, 12)
    If strFile <> "tnsnames.ora" Then
        MsgBox tns_Path & " 는 tnsnames.ora파일이 아닙니다", vbCritical + vbOKOnly, "tnsnames 생성"
        Create_Tnsnames = False
        Exit Function
    End If
    
    ' 폴더가 존재하는지
    strFolder = Split(UCase(tns_Path), "NETWORK\ADMIN\")
    strFolder(0) = strFolder(0) & "NETWORK\ADMIN\"
    
    If "" = Dir(strFolder(0), vbDirectory) Then
    
        ' 폴더 없음 FALSE(에러) 반환
        MsgBox tns_Path & "가 존재하지 않습니다."
        Create_Tnsnames = False
        Exit Function
        
    Else
        ' 폴더는 있는데 파일이 없음. >> 파일생성
        Dim FileNumber As Integer
        FileNumber = FreeFile
        
        Open tns_Path For Output As FileNumber
            ' 밑에 FileStream.ReadAll할때 파일안에 아무 TEXT도 없으면 "파일 끝을 넘어가는 입력입니다" 라는 에러발생
            Print #FileNumber, " "
        Close FileNumber
    End If
    
    Create_Tnsnames = True
End Function



'########################################
'# ORACLE Tnsnames 자동 추가
'########################################
Public Function writeOraTnsName(tns_Path As String, sTns As String, DB_Name As String) As Boolean

    Dim FS, FileStream, OutStream
    Dim strTxt As String, arrTxt() As String
    Dim i As Integer
    Dim tns_nm_flag As Boolean
    
    
    On Error GoTo err_rtn
    
    
    '>>>>>>>>>> 파일에 DBName이 없을경우 추가해서 파일을 덮어씌운다. <
    Set FS = CreateObject("Scripting.FileSystemObject")
    
    If tns_Path <> "" Then
        
         Set FileStream = FS.OpenTextFile(tns_Path)
        
         strTxt = FileStream.ReadAll
         arrTxt = Split(strTxt, vbCrLf)
         FileStream.Close
        
         tns_nm_flag = False
         For i = 0 To UBound(arrTxt)
             If UCase(Left(arrTxt(i), Len(DB_Name))) = UCase(DB_Name) Then
                 arrTxt(i) = sTns
                 tns_nm_flag = True
             End If
         Next

         If tns_nm_flag = False Then
            ReDim Preserve arrTxt(UBound(arrTxt) + 1)
            arrTxt(UBound(arrTxt)) = sTns
         End If
         
         strTxt = Join(arrTxt, vbCrLf)
        
         Set OutStream = FS.OpenTextFile(tns_Path, 2, True)
         OutStream.Write strTxt
         OutStream.Close
         
         writeOraTnsName = True       '정상적 처리
    Else
    
         Set OutStream = FS.OpenTextFile(tns_Path, 2, True)
         OutStream.Write sTns
         OutStream.Close
    
    End If

    Set FS = Nothing
   
   writeOraTnsName = True                '정상적 처리
   Exit Function
   
err_rtn:
   Set FS = Nothing
   writeOraTnsName = False               '비정상적 처리 (경로가 존재하지 않는 경우.)
   
End Function




'Public Function Find_DB_Tnsnames_Location() As String
'    '>> Registry OPEN
'    Dim hSubKey     As Long
'    Dim dType       As Long
'    Dim strLength   As Long
'    Dim RetStr      As String
'    Dim Rtn         As Long
'
'    '>> Registry Save
'    Dim iSecurity   As SECURITY_ATTRIBUTES
'    Dim strSubkey   As String
'    Dim KeyRet      As Long
'    Dim dPosition   As Long
'    Dim iRet        As Long
'
'    Dim sTmp        As String
'    Dim sReturn     As String
'
'
'    sReturn = ""
'
'    '>> registry open
'
'    Rtn = RegOpenKeyEx(hKey, "SYSTEM\CurrentControlSet\Control\Session Manager\Environment\", 0, KEY_ALL_ACCESS, hSubKey)
'    If Rtn <> ERROR_SUCCESS Then
'        '레지스트리를 생성시켜주어야 한다.
'        iRet = RegCreateKeyEx(hKey, SubKey, 0, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, iSecurity, KeyRet, dPosition)
'        If iRet <> ERROR_SUCCESS Then
'            MsgBox "오라클을 설치하여 주십시요.", vbExclamation + vbOKOnly, "DB Connection"
'            Rtn = RegCloseKey(hSubKey)
'
'            Find_DB_Tnsnames_Location = sReturn
'            Exit Function
'        End If
'    End If
'
'    '하위 키값을 얻는다.
'    strLength = 256
'    RetStr = String(strLength, 0)
'
'
'    '>> oracle path
'        Rtn = RegQueryValueEx(hSubKey, "Path", 0, dType, ByVal RetStr, strLength)
'
'        If Rtn = ERROR_SUCCESS And dType = REG_SZ Then
'            '뒤에 따라오는 문자열을 제거한다
'            sReturn = Left(RetStr, strLength - 1)
'            sReturn = sReturn & "\network\admin\tnsnames.ora"
'
'            Rtn = RegQueryValueEx(hSubKey, "ORACLE_HOME_NAME", 0, dType, ByVal RetStr, strLength)
'
'            OracleVer = "ORAHOME8"
'            If strLength > 0 Then OracleVer = Trim(Left(RetStr, strLength - 1))
'
'            If InStr(1, UCase(OracleVer), "ORAHOME8", vbTextCompare) > 0 Then OracleVer = "ORA8"
'            If InStr(1, UCase(OracleVer), "ORAHOME9", vbTextCompare) > 0 Then OracleVer = "ORA9"
'
'        Else
'            MsgBox "오라클 경로설정에 문제가 있습니다.", vbExclamation + vbOKOnly, "DB Connection"
'
'        End If
'        Rtn = RegCloseKey(hSubKey)
'
'        Find_DB_Tnsnames_Location = sReturn
'
'
'End Function
