<%@CodePage=65001%>
<%
'==============================
' 작 성 자 : 노광옹
' 작 성 일 : 2024-09-04
' 파일내용 : 근로정보 관리용 프로세스
'==============================
Option Explicit
%>
<!-- #include virtual = "/common/library/common.asp"-->
<%
Execute("Call "& sCommand &"_Process()")    '함수 호츨

Dim sCommand, sLevel, sId, sNm, sPw, sNum,sIdx, sDept, sPage, sPosition, sCompany, sHospitalType
sCommand		= nullFillWith(Request("command"), "", False)
sLevel			= nullFillWith(Request("level"), "", False)
sId				= nullFillWith(Request("id"), "", False)
sIdx			= nullFillWith(Request("idx"), "", False)
sNm				= nullFillWith(Request("nm"), "", False)
sPw				= nullFillWith(Request("pw"), "", False)
sNum			= nullFillWith(Request("num"), "", False)
sDept			= nullFillWith(Request("dept"), "", False)
sPage			= nullFillWith(Request("page"), "", False)
sPosition		= nullFillWith(Request("position"), "", False)
sCompany		= nullFillWith(Request("company"), "", False)
sHospitalType	= nullFillWith(Request("hospitalType"), "", False)


Sub USER_DATA_Process()	
	Dim sql , rs, OP_DATA_FLAG , OP_DATA_IDX
	sql = "EXEC USP_AMIN_DATA_PROC '{P1}','{P2}','{P3}','{P4}','{P5}','{P6}','{P7}','{P8}','{P9}','{P10}';"
	sql = Replace(sql , "{P1}" , sIdx , vbTextCompare)
	sql = Replace(sql , "{P2}" , sId , vbTextCompare)
	sql = Replace(sql , "{P3}" , sNm , vbTextCompare)
	sql = Replace(sql , "{P4}" , sPw , vbTextCompare)
	sql = Replace(sql , "{P5}" , sLevel , vbTextCompare)
	sql = Replace(sql , "{P6}" , sNum , vbTextCompare)
	sql = Replace(sql , "{P7}" , sDept , vbTextCompare)
	sql = Replace(sql , "{P8}" , sCompany , vbTextCompare)
	sql = Replace(sql , "{P9}" , sPosition , vbTextCompare)
	sql = Replace(sql , "{P10}" , sHospitalType , vbTextCompare)
	
	Set rs = getQueryDB(sql)
	If Not(rs.Eof Or rs.Bof) Then 
		OP_DATA_FLAG	= rs("DATA_FLAG")
		OP_DATA_IDX		= rs("DATA_IDX")
	End If 
	rsClose(rs)
	
	If OP_DATA_FLAG = "EXISTS_ID" Then 
		Call onlyAlert("warning","중복된 아이디 입니다!\n다른 아이디를 입력해 주세요!")
	Else 
		Call gotoAlert("success" , "처리완료되었습니다.","m" , "./user_list.asp?page=" & sPage & Session("Param"))
	End If 
End Sub 

Sub USER_REMOVE_Process()
	Dim sql , rs, OP_DATA_FLAG , OP_DATA_IDX
	sql = "EXEC USP_AMIN_DATA_DEL '{P1}';"
	sql = Replace(sql , "{P1}" , sIdx , vbTextCompare)
	
	Set rs = getQueryDB(sql)
	If Not(rs.Eof Or rs.Bof) Then 
		OP_DATA_FLAG	= rs("DATA_FLAG")
		OP_DATA_IDX		= rs("DATA_IDX")
	End If 
	rsClose(rs)

	Call gotoAlert("success" , "처리완료되었습니다.","m" , "./user_list.asp?page=" & sPage & Session("Param"))
End Sub 
%>