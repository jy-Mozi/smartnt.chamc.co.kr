<%@CodePage=65001%>
<%
'==============================
' 작 성 자 : 노광옹
' 작 성 일 : 2024-10-29
' 파일내용 : 점검일지 등록 > 배선카트
'==============================
Option Explicit
%>
<!-- #include virtual = "/common/library/common.asp"-->
<!-- #include virtual = "/common/include/RequestValue.asp"-->
<%


Execute("Call "& sCommand &"_Process()")    '함수 호츨

Sub Daily_Meal_Data_Process()

	Dim sCommand, sConfirm, sIdx, sPage, sToTime, sCurrDate, sItemIdx, sFmTime, sHospitalRoom 
	Dim sCartStaff, sStaff, sItemCode, sItemResult, sEtc, sEtcCode , sHospitalType
	sCommand = nullFillWith(Request("command"), "", False)
	sConfirm = nullFillWith(Request("confirm"), "", False)
	sIdx = nullFillWith(Request("idx"), "", False)
	sPage = nullFillWith(Request("page"), "", False)
	sToTime = nullFillWith(Request("toTime"), "", False)
	sCurrDate = nullFillWith(Request("currDate"), "", False)
	sItemIdx = nullFillWith(Request("itemIdx"), "", False)
	sFmTime = nullFillWith(Request("fmTime"), "", False)
	sHospitalRoom = nullFillWith(Request("hospitalRoom"), "", False)
	sHospitalType = nullFillWith(Request("hospitalType"), "", False)	
	sCartStaff = nullFillWith(Request("cartStaff"), "", False)
	sStaff = nullFillWith(Request("staff"), "", False)
	sItemCode = nullFillWith(Request("itemCode"), "", False)
	sItemResult = nullFillWith(Request("itemResult"), "", False)
	sEtc = nullFillWith(Request("etc"), "", False)
	sEtcCode = nullFillWith(Request("etcCode"), "" , False)

	Dim arrItemIdx, arrItemCode, arrItemResult, i
	Dim sql , rs
	arrItemIdx		= Split(sItemIdx, ",")
	arrItemCode		= Split(sItemCode, ",")
	arrItemResult	= Split(sItemResult, ",")

	For i = 0 To Ubound(arrItemIdx)
		sql = "EXEC USP_FOOD_LOTATION_RECORD_DATA_PROC '{P1}','{P2}','{P3}','{P4}','{P5}','{P6}','{P7}','{P8}','{P9}','{P10}','{P11}','{P12}';"
		sql = Replace(sql , "{P1}"  , Trim(arrItemIdx(i)) , vbTextCompare)
		sql = Replace(sql , "{P2}"  , sCurrDate , vbTextCompare)
		sql = Replace(sql , "{P3}"  , sStaff , vbTextCompare)
		sql = Replace(sql , "{P4}"  , sHospitalType , vbTextCompare)
		sql = Replace(sql , "{P5}"  , sHospitalRoom , vbTextCompare)
		sql = Replace(sql , "{P6}"  , sFmTime , vbTextCompare)
		sql = Replace(sql , "{P7}"  , sToTime , vbTextCompare)
		sql = Replace(sql , "{P8}"  , Trim(arrItemCode(i)) , vbTextCompare)
		sql = Replace(sql , "{P9}"  , Trim(arrItemResult(i)) , vbTextCompare)
		sql = Replace(sql , "{P10}"  , sEtcCode , vbTextCompare)
		sql = Replace(sql , "{P11}" , sEtc , vbTextCompare)
		sql = Replace(sql , "{P12}" , sCartStaff , vbTextCompare)
		rw(sql & "<BR/>")
		Call setQueryDB(sql)
	Next 


	Call onlyAlert("success","처리완료 되었습니다.")
End Sub 

Sub Daily_Meal_Approve_Process()
	Dim sql 
	sql = "EXEC AMD_APPROVE_PROC '{P1}','{P2}','{P3}','{P4}';" 
	sql = Replace(sql , "{P1}" , Replace(sCurrDate,"-","") , vbTextCompare)
	sql = Replace(sql , "{P2}" , "MEAL" , vbTextCompare)
	sql = Replace(sql , "{P3}" , Session("SNT_ADMIN_ID") , vbTextCompare)
	sql = Replace(sql , "{P4}" , sConfirm , vbTextCompare)
	Call setQueryDB(sql)
	Call gotoAlert("success","처리완료 되었습니다.","r","")
End Sub 

Sub Daily_Meal_Confirm_Process()
	Dim sql
	sql = "EXEC USP_FOOD_LOTATION_RECORD_DATA_COMFIRM '{P1}','{P2}','{P3}';"
	sql = Replace(sql , "{P1}" , sCurrDate , vbTextCompare)
	sql = Replace(sql , "{P2}" , sUserId , vbTextCompare)
	sql = Replace(sql , "{P3}" , sConfirm , vbTextCompare)
	Call setQueryDB(sql)
	Call gotoAlert("success","처리완료 되었습니다.","r","")
End Sub 

Sub Daily_Meal_Remove_Process()
	Dim sql
	sql = "EXEC USP_FOOD_LOCATION_RECORD_DATA_DEL '{P1}','{P2}';"
	sql = Replace(sql , "{P1}" , sCurrDate , vbTextCompare)
	sql = Replace(sql , "{P2}" , sHospitalType , vbTextCompare)
	Call setQueryDB(sql)
	Call gotoAlert("success","처리완료 되었습니다.","m","./daily_meal.asp")
End Sub 
%>