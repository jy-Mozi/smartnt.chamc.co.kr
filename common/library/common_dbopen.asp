<!-- METADATA TYPE="typelib" FILE="c:\program files\common files\system\ado\msado15.dll"-->
<%
'//=================================================================
'작  성 자 : @전륜기
'작  성 일 : @2007-11-15
'파일 내용 : @DB 연결, DB명령 함수 모음
'=================================================================//

%>
<OBJECT RUNAT="SERVER" PROGID="ADODB.CONNECTION" ID="DB"></OBJECT>

<%

'OLEDB 접속정보
DB.CursorLocation = adUseClient		'Client Cursor
DB.CommandTimeOut = 300
DB.Open("file name=D:\dbconn\smartnt_db.udl")



'//=================================================================
'* 기능 : 	DB와 연결을 확인하고 SQL문을 명령한다.
Function setQueryDB(fSQL)
	DB.Execute fSQL
End Function

'* 기능 : DB와 연결을 확인하고 SQL문을 명령한다. (Disconnected Recordset)
Function getQueryDB(fSQL)
	Set getQueryDB = Server.CreateObject("ADODB.RecordSet")
	getQueryDB.CursorLocation = adUseClient
	getQueryDB.Open fSQL, DB, adOpenForwardOnly, adLockReadOnly
	Set getQueryDB.ActiveConnection = Nothing
End Function

'* 기   능 : 쿼리문의 조건절 만들기
Function getQuery(fQuery)
	IF isNone(fQuery) Then
		getQuery = "WHERE "& fQuery
	Else
		getQuery = fQuery &" AND "
	End if
End Function

'* 기   능 : DB의 데이터 가져오기
'* 사용법 : Set vRs = getSelectQuery("테이블이름","가져올 컬럼","조건절")
Function getSelectQuery(fTblName,fFieldNM,fWQuery)
	Dim fQuery
	fQuery = "SELECT "& fFieldNM &" FROM "& fTblName &" "& fWQuery
	Set getSelectQuery = getQueryDB(fQuery)
End Function

'* 기   능 : 리스트 기본 적인 쿼리
'* 사용법 : Set vRs = getListQuery("테이블이름","가져올 컬럼","조건절","게시물 수","현재페이지")
Function getListQuery(fTblName,fFieldNM,fWQuery,fListSize,fPage)
	fSWQuery = Replace(fWQuery,"WHERE","AND")
	Dim fQuery,fSWQuery,fSAQuery
	fQuery = "SELECT TOP "& fListSize &" "& fFieldNM &" FROM "& fTblName
	fQuery = fQuery &" WHERE NUM NOT IN"
	fQuery = fQuery &" (SELECT TOP "&  CStr(fListSize * (fPage - 1)) &  " NUM FROM "& fTblName
	fQuery = fQuery &" "& fWQuery &" ORDER BY NUM DESC) "& fSWQuery
	fQuery = fQuery &" ORDER BY NUM DESC"
	Set getListQuery = getQueryDB(fQuery)
End Function

'* 기   능 :리스트 PK를 지정할 수 있는 쿼리
'* 사용법 : Set vRs = getListQuery("테이블이름","PK","가져올 컬럼","조건절","정렬","게시물 수","현재페이지")
Function getListQueryEtc(fTblName,fIdx,fFieldNM,fWQuery,fOrder,fListSize,fPage)
	fSWQuery = Replace(fWQuery,"WHERE","AND")
	Dim fQuery,fSWQuery,fSAQuery
	fQuery = "SELECT TOP "& fListSize &" "& fFieldNM &" FROM "& fTblName
	fQuery = fQuery &" WHERE "& fIdx &" NOT IN"
	fQuery = fQuery &" (SELECT TOP "&  CStr(fListSize * (fPage - 1)) &  " "& fIdx &" FROM "& fTblName
	fQuery = fQuery &" "& fWQuery &" ORDER BY "& fOrder &") "& fSWQuery
	fQuery = fQuery &" ORDER BY "& fOrder
	Set getListQueryEtc = getQueryDB(fQuery)
End Function

'* 기   능 : 게시물의 총 수
'* 사용법 : Set vRs = getSelectQuery("테이블이름","조건절")
Function getTotalCnt(fTblName,fWQuery,fCnt)
	Dim fQuery,fSWQuery
	fQuery = "SELECT COUNT(*) AS TOTAL_CNT, CEILING(CAST(COUNT(*) AS FLOAT) / "& fCnt &") AS TOTAL_PAGE FROM "& fTblName
	fQuery = fQuery &" "& fWQuery
	Set getTotalCnt = getQueryDB(fQuery)
End Function

'* 기능 : 배열값을 조합하여 UPDATE 쿼리문을 만들고 쿼리문을 실행한다.
'* 사용법 : Call setUpdateQuery("테이블명","배열값","배열값","조건문")
Sub setUpdateQuery(fTblName,fFieldNM,fData,fWQuery)
	IF isArray(fFieldNM) and isArray(fData) Then
		Dim fQuery,fI,fStr()
		For fI = 0 To UBound(fFieldNM)
			ReDim Preserve fStr(fI)
			fStr(fI) = fFieldNM(fI) &" = '"& fData(fI) &"'"
		Next

		'쿼리문 작성
		fQuery = "UPDATE "& fTblName &" SET "& Join(fStr,", ") &" "& fWQuery

		setQueryDB(fQuery)
	End if
End Sub

'* 기능 : 배열값을 조합하여 INSERT 쿼리문을 만들고 쿼리문을 실행한다.
'* 사용법 : Call setInsertQuery("테이블명","배열값","배열값")
Sub setInsertQuery(fTblName,fFieldNM,fData)
	IF isArray(fFieldNM) and isArray(fData) Then
		Dim fQuery
		'쿼리문 작성
		fQuery = "INSERT INTO "& fTblName &" ("& Join(fFieldNM,",") &") VALUES ('"& Join(fData,"','") &"')"
		setQueryDB(fQuery)
	End if
End Sub

'* 기능 : DELETE 쿼리문을 만들고 쿼리문을 실행한다.
'* 사용법 : Call setDeleteQuery("테이블명","조건문")
Sub setDeleteQuery(fTblName,fWQuery)
	Dim fQuery
	fQuery = "DELETE FROM "& fTblName &" "& fWQuery
	setQueryDB(fQuery)
End Sub

'* 기능 : 	열려있는 개체를 닫아준다.
Function RsClose(fObj)
	fObj.Close
	Set fObj = Nothing
End Function

'* 기능 : SP를 실행한다. 레코드셋이 없을때 사용
'* 사용법 : Call setSP("SP이름", "배열")
Sub setSP(fSP, fParam)
	On Error Resume Next
	Dim fCMD, fI
	Set fCMD = Server.CreateObject("ADODB.Command")
	With fCMD
		.ActiveConnection = DB
		.CommandText = fSP
		.CommandType = adCmdStoredProc
		.CommandTimeOut = 0
		IF isArray(fParam) Then
			For fI = 0 To UBound(fParam)
				.Parameters(fI + 1) = fParam(fI)
			Next
		End if

		.Execute, , adExecuteNoRecords

	End With
	Set fCMD = Nothing

	IF Err Then
		%>
		<script type="text/javascript">
			alert("진행중 Database 에러가 발생하였습니다. \n다시 한번 시도하여 주세요.\n\n에러SP<%=fSP%>\n<%=Err.Description%>");
		</script>
		<%
		RW("에러SP	 : "& fSP &"<br/>")
		RW("에러번호 : "& Err.Number &"<br/>")
		RW("에러설명 : "& Err.Description &"<br/>")
		RW("에러소스 : "& Err.Source &"<br/>")
		RW("에러라인 : "& Err.Line &"<br/>")
		RE()
	End if
End Sub

'* 기능 : SP를 실행하고 레코드셋을 받아온다.
'* 사용법 : Set vRs = getSP("SP이름", "배열")
Function getSP(fSP, fParam)
	On Error Resume Next
	Dim fCMD, fI
	Set fCMD = Server.CreateObject("ADODB.Command")
	With fCMD
		.ActiveConnection = DB
		.CommandText = fSP
		.CommandType = adCmdStoredProc
		.CommandTimeOut = 0
		IF isArray(fParam) Then
			For fI = 0 To UBound(fParam)
				.Parameters(fI + 1) = fParam(fI)
			Next
		End if
		Set getSP = .Execute
	End With
	Set fCMD = Nothing

	IF Err Then
		%>
		<script type="text/javascript">
			alert("진행중 Database 에러가 발생하였습니다. \n다시 한번 시도하여 주세요.\n\n에러SP<%=fSP%>\n<%=Err.Description%>");
		</script>
		<%
		RW("에러SP	 : "& fSP &"<br/>")
		RW("에러번호 : "& Err.Number &"<br/>")
		RW("에러설명 : "& Err.Description &"<br/>")
		RW("에러소스 : "& Err.Source &"<br/>")
		RW("에러라인 : "& Err.Line &"<br/>")
		RE()
	End if
End Function


'* 기능 : SP를 실행한다. 레코드셋이 없을때 사용
'* 사용법 : Call setSP("SP이름", "배열")
Sub setSPROC(fSP, fParam)
	Dim fCMD, fI
	Set fCMD = Server.CreateObject("ADODB.Command")
	With fCMD
		.ActiveConnection = DB
		.CommandText = fSP
		.CommandType = adCmdStoredProc
		.CommandTimeOut = 0
		IF isArray(fParam) Then
			For fI = 0 To UBound(fParam)
				.Parameters(fI + 1) = fParam(fI)
			Next
		End if

		.Execute, , adExecuteNoRecords
	End With
	Set fCMD = Nothing
End Sub

'* 기능 : SP를 실행하고 output을 받는다.
'* 사용법 : Set vRs = getSP("SP이름", "배열", "OUTPUT으로 받을 배열")
Sub setSPOP(fSP, fParams, fOParams)
	On Error Resume Next
	Dim fCMD, fI, fJ, fCount
	Set fCMD = Server.CreateObject("ADODB.Command")
	With fCMD
		.ActiveConnection = DB
		.CommandText = fSP
		.CommandType = adCmdStoredProc
		.CommandTimeOut = 0

		fCount = 1	'매개변수의 총 카운터
		IF isArray(fParams) Then	'input
			For fI = 0 To UBound(fParams)
				.Parameters(fCount) = fParams(fI)
				fCount = fCount + 1
			Next
		End if

		IF isArray(fOParams) Then	'output
			For fJ = 0 To UBound(fOParams)
				.Parameters(fCount) = fOParams(fJ)
				fCount = fCount + 1
			Next
		End if

		.Execute, , adExecuteNoRecords	'Execute

		IF isArray(fOParams) Then	'output을 매개변수 배열에 다시 셋팅
			fCount = fCount - (UBound(fOParams) + 1)
			For fJ = 0 To UBound(fOParams)
				fOParams(fJ) = .Parameters(fCount)
				fCount = fCount + 1
			Next
		End if

	End With
	Set fCMD = Nothing

	IF Err Then
		RW("에러SP	 : "& fSP &"<br/>")
		RW("에러번호 : "& Err.Number &"<br/>")
		RW("에러설명 : "& Err.Description &"<br/>")
		RW("에러소스 : "& Err.Source &"<br/>")
		RW("에러라인 : "& Err.Line &"<br/>")
		RE()
	End if
End Sub

'* 미완성 output을 못 받고 있음!!!!
'* 기능 : SP를 실행하고 output을 받는다.
'* 사용법 : Set vRs = getSPOP("SP이름", "배열", "OUTPUT으로 받을 배열")
Function getSPOP(fSP, fParams, fOParams)
	On Error Resume Next
	Dim fCMD,fI, fJ, fCount
	Set fCMD = Server.CreateObject("ADODB.Command")
	With fCMD
		.ActiveConnection = DB
		.CommandText = fSP
		.CommandType = adCmdStoredProc
		.CommandTimeOut = 0

		fCount = 1	'매개변수의 총 카운터
		IF isArray(fParams) Then	'input
			For fI = 0 To UBound(fParams)
				.Parameters(fCount) = fParams(fI)
				fCount = fCount + 1
			Next
		End if

		IF isArray(fOParams) Then	'output
			For fJ = 0 To UBound(fOParams)
				.Parameters(fCount) = fOParams(fJ)
				fCount = fCount + 1
			Next
		End if

		Set getSPOP = .Execute	'Execute

		IF isArray(fOParams) Then	'output을 매개변수 배열에 다시 셋팅
			fCount = fCount - (UBound(fOParams) + 1)
			For fJ = 0 To UBound(fOParams)
				fOParams(fJ) = .Parameters(fCount)
				fCount = fCount + 1
			Next
		End if

	End With
	Set fCMD = Nothing

	IF Err Then
		RW("에러SP	 : "& fSP &"<br/>")
		RW("에러번호 : "& Err.Number &"<br/>")
		RW("에러설명 : "& Err.Description &"<br/>")
		RW("에러소스 : "& Err.Source &"<br/>")
		RW("에러라인 : "& Err.Line &"<br/>")
		RE()
	End if
End Function

'* 기능 : SP내용을 화면에 찍어준다.
'* 사용법 : Call getText("저장프로시져")
Sub getText(fProc)
	Dim fRs, fData
	Set fRs = getQueryDB("sp_helptext "& fProc)
	IF Not(fRs.Bof or fRs.Eof) Then
		fData = fRs.GetRows()
	End if
	RsClose(fRs)

	Dim fI, fJ
	IF isArray(fData) Then
		For fI = 0 To UBound(fData, 2)
			For fJ = 0 To UBound(fData, 1)
				RW(fData(fJ, fI) &"<br>")
			Next
		Next
	End if
End Sub

'해당 테이블의 데이터 가져오기
Sub getData(fTop, fTable)
	'테이블 컬럼명 가져오기
	Dim fQuery, fRs, fTData
	fQuery = "SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS"
	fQuery = fQuery &" WHERE TABLE_NAME = '"& fTable &"' ORDER BY ORDINAL_POSITION ASC"
	Set fRs = getQueryDB(fQuery)
	IF Not(fRs.Bof or fRs.Eof) Then
		fTData = fRs.GetRows()
	End if
	RsClose(fRs)

	'테이블 저장된 데이터 가져오기
	Dim fData
	fQuery = "SELECT "& fTop &" * FROM "& fTable
	fQuery = fQuery &" ORDER BY 1 DESC"
	Set fRs = getQueryDB(fQuery)
	IF Not(fRs.Bof or fRs.Eof) Then
		fData = fRs.GetRows()
	End if
	RsClose(fRs)

	Dim fI, fG, fStr
	fStr = ""
	fStr = "<table border=""1"" bordercolor=""#000"" cellpadding=""1"" cellspacing=""2"" >"
	IF isArray(fTData) Then
		fStr = fStr &"<tr align=""center"">"
		For fG = 0 To UBound(fTData, 2)
			fStr = fStr &"<td style=""padding:2px 5px;"">"& nullFillWith(fTData(0, fG), "&nbsp;", False) &"</td>"
		Next
		fStr = fStr &"</tr>"
	End if

	IF isArray(fData) Then
		For fG = 0 To UBound(fData, 2)
			fStr = fStr &"<tr align=""center"">"
			For fI = 0 To UBound(fData, 1)
				fStr = fStr &"<td style=""padding:2px 5px;"">"& nullFillWith(fData(fI, fG), "&nbsp;", False) &"</td>"
			Next
			fStr = fStr &"</tr>"
		Next
	End if
	fStr = fStr &"</table>"
	RWE(fStr)
End Sub
'=================================================================//
%>