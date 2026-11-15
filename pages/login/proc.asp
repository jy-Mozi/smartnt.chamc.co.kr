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
Dim sCommand, sPassword, sId
sCommand	= nullFillWith(Request("command"), "", False)
sPassword	= nullFillWith(Request("password"), "", False)
sId			= nullFillWith(Request("id"), "", False)


Execute("Call "& sCommand &"_Process()")    '함수 호츨

Sub LoginChk_Process()
	Dim sql , rs 
	sql = "EXEC USP_LOGIN_CHK '{P1}','{P2}';"
	sql = Replace(sql , "{P1}" , sId , vbTextCompare)
	sql = Replace(sql , "{P2}" , sPassword , vbTextCompare)
	Set rs = getQueryDB(sql)

	Dim OP_LOGIN_CHK, OP_ADMIN_ID, OP_EMPLOYEE_NUM, OP_ADMIN_LEVEL, OP_ADMIN_NM, currIp
	Dim OP_ADMIN_COMPANY , OP_ADMIN_POSITION, OP_HOSPITAL_TYPE, OP_STR_POSITION
	If Not(rs.Eof Or rs.Bof) Then 
		OP_LOGIN_CHK		= rs("LOGIN_YN")
		OP_ADMIN_ID			= rs("ADMIN_ID")
		OP_ADMIN_NM			= rs("ADMIN_NM")
		OP_EMPLOYEE_NUM		= rs("EMPLOYEE_NUM")
		OP_ADMIN_LEVEL		= rs("ADMIN_LEVEL")
		OP_ADMIN_COMPANY	= rs("ADMIN_COMPANY") 
		OP_ADMIN_POSITION	= rs("ADMIN_POSITION")
		OP_STR_POSITION		= rs("STR_POSITION")
		OP_HOSPITAL_TYPE	= rs("HOSPITAL_TYPE")
		OP_HOSPITAL_TYPE	= nullFillWith(OP_HOSPITAL_TYPE , "" , False)
	End If 
	rsClose(rs)


	currIp	= Request.ServerVariables("REMOTE_ADDR")
	sql = "EXEC USP_LOGIN_LOG_ADD '{P1}','{P2}','{P3}';"
	sql = Replace(sql , "{P1}", OP_ADMIN_ID , vbTextCompare)
	sql = Replace(sql , "{P2}", currIp , vbTextCompare)
	sql = Replace(sql , "{P3}", OP_LOGIN_CHK , vbTextCompare)
	
	Call SetQueryDB(sql)


	If OP_LOGIN_CHK = "Y" Then 

		Session("SNT_ADMIN_ID")		= OP_ADMIN_ID
		Session("SNT_EMPLOYEE_NUM")	= OP_EMPLOYEE_NUM
		Session("SNT_ADMIN_LEVEL")	= OP_ADMIN_LEVEL
		Session("SNT_ADMIN_NM")	= OP_ADMIN_NM
		Session("SNT_ADMIN_COMPANY")	= OP_ADMIN_COMPANY
		Session("SNT_ADMIN_POSITION")	= OP_ADMIN_POSITION
		Session("SNT_ADMIN_POSITION_STR")	= OP_STR_POSITION		
		Session("SNT_HOSPITAL_TYPE")    = OP_HOSPITAL_TYPE
		Session.Timeout = 360
		Call errorGotoParent("/pages/checklist/meal_cart.asp")
	Else 
		Session.Contents.Remove("SNT_ADMIN_ID")
		Session.Contents.Remove("SNT_EMPLOYEE_NUM")
		Session.Contents.Remove("SNT_ADMIN_LEVEL")
		Session.Contents.Remove("SNT_ADMIN_NM")
		Session.Contents.Remove("SNT_ADMIN_COMPANY")
		Session.Contents.Remove("SNT_ADMIN_POSITION")
		Session.Contents.Remove("SNT_ADMIN_POSITION_STR")		
		Session.Contents.Remove("SNT_HOSPITAL_TYPE")
		Session.Contents.RemoveAll()
		Session.Abandon()
		Call onlyAlert("error","입력한 정보가 올바르지 않습니다.\n입력하신 내용을 다시 확인해주세요.")
	End If 
End Sub 

Sub Logout_Chk_PROCESS()
	Session.Contents.Remove("SNT_ADMIN_ID")
	Session.Contents.Remove("SNT_EMPLOYEE_NUM")
	Session.Contents.Remove("SNT_ADMIN_LEVEL")
	Session.Contents.Remove("SNT_ADMIN_NM")
	Session.Contents.RemoveAll()
	Session.Abandon()
	Call errorGotoParent("/pages/login/")
End Sub 

%>