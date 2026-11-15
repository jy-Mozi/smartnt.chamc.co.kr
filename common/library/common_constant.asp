<%
'//=================================================================
'작  성 자 : @전륜기
'작  성 일 : @2007-11-15
'파일 내용 : @사이트 전역변수 처럼(?) 사용할 변수모음
'=================================================================//

Const gScripDomain = "FMS.CHAMC.CO.KR"
Const gOnSiteManagerCode = "F0010004"
'기타 상수
CONST MAX_SIZE		= 10485760 '업로드 최대 용량
                      
'CONST USER_PATH		= "/upload/"			'업로드 폴더의 최상단 루트

'Const gYearsSett = year(now()) & "-03-01"

Dim gFileUpPath	'사이트 업로드 경로
'gFileUpPath = Server.MapPath("/") & "\upload\"
gFileUpPath = "D:\wwwroot\fms.chamc.co.kr\upload\"

%>