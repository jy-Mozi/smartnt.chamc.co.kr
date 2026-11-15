<%@CodePage=65001%>
<%
'==============================
' 작 성 자 : 송 쟁 영
' 작 성 일 : 2025-03-21 오후 2:29
' 파일내용 : 결재 승인 프로세스
'==============================
Option Explicit
%>
<!-- #include virtual = "/common/library/common.asp"-->
<%

Execute("Call "& sCommand &"_Process()")    '함수 호츨

Sub Approve_Process()
	Dim sCommand, sFLAG, sApproveGrade, sSelDate, sHospitalType, sHospitalRoom
	Dim sql , sVal
	sCommand		= nullFillWith(Request("command"), "", False)
	sFLAG			= nullFillWith(Request("FLAG"), "", False)
	sApproveGrade	= nullFillWith(Request("approveGrade"), "", False)
	sSelDate		= nullFillWith(Request("selDate"), "", False)
	sHospitalType		= nullFillWith(Request("hospitalType"), "", False)
	sHospitalRoom   = nullFillWith(Request("hospitalRoom"), "", False)

	If sFLAG = "MEAL" Then 
		sVal = sHospitalType
	Else 
		sVal = ""
	End If 

	sql = "EXEC AMD_APPROVE_PROC '{P1}','{P2}','{P3}','{P4}','{P5}','{P6}';" 
	sql = Replace(sql , "{P1}" , Replace(sSelDate,"-","") , vbTextCompare)
	sql = Replace(sql , "{P2}" , sFLAG , vbTextCompare)
	sql = Replace(sql , "{P3}" , sHospitalType , vbTextCompare)
	sql = Replace(sql , "{P4}" , Session("SNT_ADMIN_ID") , vbTextCompare)
	sql = Replace(sql , "{P5}" , sApproveGrade , vbTextCompare)
	sql = Replace(sql , "{P6}" , sVal , vbTextCompare)

	Call setQueryDB(sql)
	Call gotoAlert("success","결재 처리 되었습니다.","r","")
End Sub 


Sub ApproveCancel_Process()
	Dim sCommand, sFLAG, sApproveGrade, sHospitalType, sSelDate, sHospitalRoom
	Dim sql , sVal
	
	sCommand		= nullFillWith(Request("command"), "", False)
	sFLAG			= nullFillWith(Request("FLAG"), "", False)
	sApproveGrade	= nullFillWith(Request("approveGrade"), "", False)
	sHospitalType		= nullFillWith(Request("hospitalType"), "", False)
	sSelDate		= nullFillWith(Request("selDate"), "", False)
	sHospitalRoom   = nullFillWith(Request("hospitalRoom"), "", False)

	If sFLAG = "MEAL" Then 
		sVal = sHospitalType
	Else 
		sVal = ""
	End If 

	sql = "EXEC AMD_APPROVE_CANCEL_PROC '{P1}','{P2}','{P3}','{P4}','{P5}','{P6}';" 
	sql = Replace(sql , "{P1}" , Replace(sSelDate,"-","") , vbTextCompare)
	sql = Replace(sql , "{P2}" , sFLAG , vbTextCompare)
	sql = Replace(sql , "{P3}" , sHospitalType , vbTextCompare)
	sql = Replace(sql , "{P4}" , Session("SNT_ADMIN_ID") , vbTextCompare)
	sql = Replace(sql , "{P5}" , sApproveGrade , vbTextCompare)
	sql = Replace(sql , "{P6}" , sHospitalRoom , vbTextCompare)
	'rwe(sql)
	Call setQueryDB(sql)
	Call gotoAlert("success","결재 취소 되었습니다.","r","")
End Sub 
%>