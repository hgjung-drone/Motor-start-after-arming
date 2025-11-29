Option Compare Database
Option Explicit

' === 보고서용 쿼리 생성 모듈 ===

' 수료증 발급용 쿼리 생성 함수
Public Sub CreateCertificateQuery()
    On Error Resume Next
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim sql As String

    Set db = CurrentDb

    ' 1. 수료증 발급용 쿼리 (Q_수료증발급)
    ' 주의: 0-교육생명단 테이블에 [영문이름], [주소] 필드가 있어야 합니다.

    sql = "SELECT " & _
          "T.이름 AS 성명_한글, " & _
          "T.영문이름 AS 성명_영문, " & _
          "T.생년월일, " & _
          "T.주소, " & _
          "C.과정명, " & _
          "TC.수료증ID AS 발급번호, " & _
          "TC.학과시작, TC.학과끝, TC.학과시간, " & _
          "TC.모의시작, TC.모의끝, TC.모의시간, " & _
          "TC.실기시작, TC.실기끝, TC.실기시간, " & _
          "TC.수료일자, " & _
          "TC.수료증발급일 " & _
          "FROM ([교육내용] AS TC " & _
          "INNER JOIN [0-교육생명단] AS T ON TC.교육생ID = T.ID) " & _
          "LEFT JOIN [교육과정] AS C ON TC.교육과정ID = C.ID " & _
          "WHERE TC.수료증ID IS NOT NULL " & _
          "ORDER BY TC.수료증ID DESC;"

    ' 기존 쿼리가 있다면 삭제하고 재생성
    db.QueryDefs.Delete "Q_수료증발급"
    Err.Clear

    Set qdf = db.CreateQueryDef("Q_수료증발급", sql)

    If Err.Number = 0 Then
        MsgBox "쿼리 생성 완료: Q_수료증발급" & vbCrLf & vbCrLf & _
               "이제 [만들기] -> [보고서 마법사]에서 이 쿼리를 선택하여 수료증을 디자인하세요.", vbInformation
    Else
        MsgBox "쿼리 생성 실패: " & Err.Description, vbCritical
    End If
End Sub

' 비행기록부 출력용 테이블 생성 함수 (보고서 버그 해결용)
Public Sub MakeFlightLogTable()
    On Error GoTo ErrorHandler
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim rsCourse As DAO.Recordset
    Dim sql As String
    Dim traineeID As Long
    Dim flightSeq As Long

    Set db = CurrentDb

    ' 1. 기존 출력용 테이블 삭제
    DoCmd.SetWarnings False
    On Error Resume Next
    DoCmd.DeleteObject acTable, "비행기록부_출력용"
    Err.Clear
    On Error GoTo ErrorHandler

    ' 2. [1단계] 기본 데이터로 테이블 생성
    ' '최근안전성인증일'을 포함한 핵심 필드로 테이블을 우선 생성합니다.
    sql = "SELECT " & _
          "A.기종, A.종류, A.형식, A.신고번호, A.자체중량, A.최대이륙중량, A.[최근안전성인증일], " & _
          "T.이름 AS 교육생성명, T.ID AS 교육생내부ID, F.교육생ID, " & _
          "F.교육일 AS 일자, '전남 나주' AS 비행장소, " & _
          "F.교육시작시간 AS 이륙시간, F.교육종료시간 AS 착륙시간, " & _
          "F.교육시간 AS 비행시간_분, Round(F.교육시간 / 60, 1) AS 비행시간_시간, " & _
          "S.성명 AS 교관성명, S.자격번호 AS 교관자격번호, " & _
          "(SELECT Count(*) FROM [비행기록부] AS F2 WHERE F2.교육생ID = F.교육생ID AND (F2.교육일 < F.교육일 OR (F2.교육일 = F.교육일 AND F2.교육시작시간 <= F.교육시작시간))) AS 비행순서 " & _
          "INTO [비행기록부_출력용] " & _
          "FROM (([비행기록부] AS F " & _
          "INNER JOIN [기체정보] AS A ON F.기체ID = A.ID) " & _
          "INNER JOIN [0-교육생명단] AS T ON F.교육생ID = T.출퇴근ID) " & _
          "LEFT JOIN [직원명단] AS S ON F.교관ID = S.직원ID"

    DoCmd.RunSQL sql

    ' 3. [2단계] 계산용 필드 추가
    DoCmd.RunSQL "ALTER TABLE [비행기록부_출력용] ADD COLUMN 교육과정ID LONG"
    DoCmd.RunSQL "ALTER TABLE [비행기록부_출력용] ADD COLUMN 과정명 TEXT(255)"
    DoCmd.RunSQL "ALTER TABLE [비행기록부_출력용] ADD COLUMN 목표비행횟수 INTEGER"
    DoCmd.RunSQL "ALTER TABLE [비행기록부_출력용] ADD COLUMN 비행목적 TEXT(255)"
    DoCmd.RunSQL "ALTER TABLE [비행기록부_출력용] ADD COLUMN 기장 TEXT(10)"
    DoCmd.RunSQL "ALTER TABLE [비행기록부_출력용] ADD COLUMN 훈련 TEXT(10)"
    DoCmd.RunSQL "ALTER TABLE [비행기록부_출력용] ADD COLUMN 교관 TEXT(10)"
    DoCmd.RunSQL "ALTER TABLE [비행기록부_출력용] ADD COLUMN 소계 TEXT(10)"

    ' 4. [3단계] 추가된 필드 값 계산 및 업데이트
    ' 4-1. 비행목적, 시간 등 간단한 계산 업데이트 (한 번에 실행)
    DoCmd.RunSQL "UPDATE [비행기록부_출력용] SET " & _
                 "비행목적 = Switch(비행순서=1 OR 비행순서=4 OR 비행순서=7, '이착륙', 비행순서=2 OR 비행순서=5, '전진 및 후진', 비행순서=3 OR 비행순서=6, '삼각비행', True, '종합비행'), " & _
                 "기장 = IIf(비행순서 = 7, '0.1', IIf(비행순서 >= 8, '0.3', NULL)), " & _
                 "훈련 = IIf(비행순서 <= 6, '0.3', IIf(비행순서 = 7, '0.2', NULL)), " & _
                 "소계 = '0.3'"

    ' 4-2. 교육과정 정보 업데이트 (복잡한 로직은 VBA Recordset으로 순회하며 처리)
    ' 이 방법이 복잡한 서브쿼리를 포함한 단일 UPDATE문보다 훨씬 안정적입니다.
    Set rs = db.OpenRecordset("SELECT * FROM [비행기록부_출력용] ORDER BY 교육생내부ID, 비행순서")
    If Not (rs.BOF And rs.EOF) Then
        Do While Not rs.EOF
            traineeID = rs!교육생내부ID
            flightSeq = rs!비행순서

            ' 해당 비행순서에 맞는 교육과정 정보를 찾는 쿼리
            sql = "SELECT TOP 1 TC.교육과정ID, C.과정명, C.목표비행횟수 " & _
                  "FROM [교육내용] AS TC INNER JOIN [교육과정] AS C ON TC.교육과정ID = C.ID " & _
                  "WHERE TC.교육생ID = " & traineeID & " AND " & flightSeq & " <= " & _
                  "(SELECT Sum(C2.목표비행횟수) FROM [교육내용] AS TC2 INNER JOIN [교육과정] AS C2 ON TC2.교육과정ID = C2.ID WHERE TC2.교육생ID = " & traineeID & " AND TC2.ID <= TC.ID) " & _
                  "ORDER BY TC.ID"

            Set rsCourse = db.OpenRecordset(sql)
            If Not (rsCourse.BOF And rsCourse.EOF) Then
                rs.Edit
                rs!교육과정ID = rsCourse!교육과정ID
                rs!과정명 = rsCourse!과정명
                rs!목표비행횟수 = rsCourse!목표비행횟수
                rs.Update
            End If
            rsCourse.Close

            rs.MoveNext
        Loop
    End If
    rs.Close

    ' 5. 테이블 정렬
    ' 보고서에서 정렬할 수도 있지만, 최종 테이블을 정렬된 상태로 두면 디버깅에 용이합니다.
    ' 임시 테이블을 만들고 데이터를 복사한 후, 원본을 삭제하고 임시 테이블 이름을 변경합니다.
    DoCmd.RunSQL "SELECT * INTO [비행기록부_출력용_Sorted] FROM [비행기록부_출력용] ORDER BY 기종, 교육생성명, 일자, 이륙시간;"
    DoCmd.DeleteObject acTable, "비행기록부_출력용"
    DoCmd.Rename "비행기록부_출력용_Sorted", acTable, "비행기록부_출력용"

    DoCmd.SetWarnings True

    MsgBox "출력용 테이블 생성 완료: 비행기록부_출력용" & vbCrLf & vbCrLf & _
           "이제 보고서의 레코드 원본을 '비행기록부_출력용'으로 변경하세요.", vbInformation
    Exit Sub

ErrorHandler:
    DoCmd.SetWarnings True
    If Not rs Is Nothing Then rs.Close
    If Not rsCourse Is Nothing Then rsCourse.Close
    Set rs = Nothing
    Set rsCourse = Nothing
    Set db = Nothing
    MsgBox "테이블 생성 실패 (MakeFlightLogTable): " & Err.Description & " (" & Err.Number & ")", vbCritical
End Sub

' === 비행경력증명서 관련 함수 ===

' 비행목적 문자열 합치기 함수 (중복 제거)
Public Function GetFlightPurposeString(traineeID As Long, flightDate As Date) As String
    On Error Resume Next
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String
    Dim purposeList As String
    Dim currentPurpose As String

    Set db = CurrentDb

    ' 해당 교육생, 해당 일자의 비행목적 조회 (중복 제거)
    sql = "SELECT DISTINCT 비행목적 FROM 비행기록부_출력용 " & _
          "WHERE 교육생내부ID = " & traineeID & " AND 일자 = #" & Format(flightDate, "yyyy-mm-dd") & "#"

    Set rs = db.OpenRecordset(sql)

    purposeList = ""
    Do While Not rs.EOF
        currentPurpose = Nz(rs!비행목적, "")
        If currentPurpose <> "" Then
            If purposeList = "" Then
                purposeList = currentPurpose
            Else
                purposeList = purposeList & ", " & currentPurpose
            End If
        End If
        rs.MoveNext
    Loop

    rs.Close
    GetFlightPurposeString = purposeList
End Function

' 비행경력증명서 출력용 테이블 생성 함수
Public Sub MakeFlightExperienceTable()
    On Error GoTo ErrorHandler
    Dim db As DAO.Database
    Dim sql As String

    Set db = CurrentDb

    ' 1. 비행기록부_출력용 테이블이 최신인지 확인 (재생성)
    MakeFlightLogTable

    ' 2. 기존 비행경력증명서_출력용 테이블 삭제
    On Error Resume Next
    DoCmd.DeleteObject acTable, "비행경력증명서_출력용"
    On Error GoTo ErrorHandler
    Err.Clear

    ' 3. [1단계] 기본 집계 테이블 생성 (교육내용 조인 제외, 비행목적 제외)
    ' 복잡도를 낮추기 위해 가장 기본적인 조인만 수행
    sql = "SELECT " & _
          "T.교육생내부ID, T.교육생성명, S.생년월일, S.주소, " & _
          "T.교육과정ID, T.과정명, " & _
          "T.일자, " & _
          "First(T.기종) AS 기종, First(T.형식) AS 형식, First(T.신고번호) AS 신고번호, " & _
          "First(T.[최근안전성인증일]) AS 최근안전성인증일, First(T.자체중량) AS 자체중량, First(T.최대이륙중량) AS 최대이륙중량, " & _
          "First(T.비행장소) AS 비행장소, " & _
          "Count(*) AS 비행회수, " & _
          "Round(Sum(T.비행시간_분)/60, 1) AS 비행시간_합계, " & _
          "Round(Sum(Val(Nz(T.기장,'0'))), 1) AS 기장시간, " & _
          "Round(Sum(Val(Nz(T.훈련,'0'))), 1) AS 훈련시간, " & _
          "Round(Sum(Val(Nz(T.교관,'0'))), 1) AS 교관시간, " & _
          "Round(Sum(Val(Nz(T.소계,'0'))), 1) AS 소계, " & _
          "First(T.교관성명) AS 교관성명, First(T.교관자격번호) AS 교관자격번호 " & _
          "INTO 비행경력증명서_출력용 " & _
          "FROM (비행기록부_출력용 AS T " & _
          "INNER JOIN [0-교육생명단] AS S ON T.교육생내부ID = S.ID) " & _
          "GROUP BY T.교육생내부ID, T.교육생성명, S.생년월일, S.주소, T.교육과정ID, T.과정명, T.일자 " & _
          "ORDER BY T.교육생성명, T.일자;"

    ' 3061/3920 오류 방지를 위해 DoCmd.RunSQL 사용
    DoCmd.SetWarnings False
    DoCmd.RunSQL sql

    ' 1단계 성공

    ' 4. [2단계] 컬럼 추가 (비행목적, 발급일, 발급번호)
    ' TEXT(255)는 Short Text, DATETIME은 날짜/시간
    ' DDL 문도 RunSQL로 실행 (일관성)
    DoCmd.RunSQL "ALTER TABLE 비행경력증명서_출력용 ADD COLUMN 비행목적 TEXT(255)"
    DoCmd.RunSQL "ALTER TABLE 비행경력증명서_출력용 ADD COLUMN 발급일 DATETIME"
    DoCmd.RunSQL "ALTER TABLE 비행경력증명서_출력용 ADD COLUMN 발급번호 TEXT(50)"

    ' 5. [3단계] 비행목적 업데이트 (VBA 함수 사용)
    sql = "UPDATE 비행경력증명서_출력용 SET 비행목적 = Trim(GetFlightPurposeString(교육생내부ID, 일자))"
    ' db.Execute는 VBA 함수를 인식하지 못할 수 있으므로 DoCmd.RunSQL 사용
    DoCmd.SetWarnings False
    DoCmd.RunSQL sql
    DoCmd.SetWarnings True

    ' 6. [4단계] 발급일/발급번호 업데이트 (교육내용 테이블 조인)
    sql = "UPDATE 비행경력증명서_출력용 AS T " & _
          "INNER JOIN [교육내용] AS C ON (T.교육생내부ID = C.교육생ID AND T.교육과정ID = C.교육과정ID) " & _
          "SET T.발급일 = C.수료증발급일, T.발급번호 = C.수료증ID"
    DoCmd.RunSQL sql
    DoCmd.SetWarnings True

    MsgBox "비행경력증명서 출력용 테이블 생성 완료: 비행경력증명서_출력용" & vbCrLf & vbCrLf & _
           "일자별 집계 및 데이터 병합이 완료되었습니다.", vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "데이터 업데이트 실패 (MakeFlightExperienceTable): " & Err.Description & " (" & Err.Number & ")", vbCritical
End Sub
