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
<!-- #include virtual = "/common/include/requestValue.asp"-->
<%
Dim sCommand, sDinner, sIdx, sManagerId, sPage, sCurrDate, sItemIdx 
Dim sAreaCode, sTakeAction, sItemCode,Breakfast, sLunch, sEtc1, sEtc2, sEtc3, sConfirm
Dim arrEtc , arrTake , arrManagerId
Dim sBrackfastC, sLunchC, sDinnerC, sBrackfastH, sLunchH, sDinnerH

sCommand = nullFillWith(Request("command"), "", False)
sDinner = nullFillWith(Request("dinner"), "", False)
sIdx = nullFillWith(Request("idx"), "", False)
sManagerId = nullFillWith(Request("managerId"), "", False)
sPage = nullFillWith(Request("page"), "", False)
sCurrDate = nullFillWith(Request("currDate"), "", False)
sItemIdx = nullFillWith(Request("itemIdx"), "", False)
sAreaCode = nullFillWith(Request("selAreaCode"), "", False)
sTakeAction = nullFillWith(Request("takeAction"), "", False)
sItemCode = nullFillWith(Request("itemCode"), "", False)
sBreakfast = nullFillWith(Request("breakfast"), "", False)
sLunch = nullFillWith(Request("lunch"), "", False)
sEtc = nullFillWith(Request("etc"), "", False)
sConfirm  = nullFillWith(Request("confirm"), "", False)	

sBrackfastC	= nullFillWith(Request("brackfastC"), "", False)
sLunchC		= nullFillWith(Request("lunchC"), "", False)
sDinnerC	= nullFillWith(Request("dinnerC"), "", False)
sBrackfastH	= nullFillWith(Request("brackfastH"), "", False)
sLunchH		= nullFillWith(Request("lunchH"), "", False)
sDinnerH	= nullFillWith(Request("dinnerH"), "", False)



arrEtc			= Split(sEtc , ",")
arrTake			= Split(sTakeAction , ",")
arrManagerId	= Split(sManagerId , ",")



Execute("Call "& sCommand &"_Process()")    '함수 호츨

Sub machine_celsius_Data_Process()
	
	Dim arrItemIdx, arrItemCode, arrBreakfast, arrLunch, arrDinner
	Dim sql , rs, i
	arrItemIdx		= Split(sItemIdx, ",")
	arrItemCode		= Split(sItemCode, ",")
	arrBreakFast	= Split(sBreakfast, ",")
	arrLunch		= Split(sLunch, ",")
	arrDinner		= Split(sDinner, ",")

	For i = 0 To Ubound(arrItemIdx)
		'조식
		sql = "EXEC USP_MACHINE_CELSIUS_RECORD_DATA_MEAL_PROC '{P1}','{P2}','{P3}','{P4}','{P5}','{P6}','{P7}','{P8}','{P9}','{P10}','{P11}';"
		sql = Replace(sql , "{P1}" , Trim(arrItemIdx(i)) , vbTextCompare)
		sql = Replace(sql , "{P2}" , sCurrDate , vbTextCompare)
		sql = Replace(sql , "{P3}" , sAreaCode , vbTextCompare)
		sql = Replace(sql , "{P4}" , "1" , vbTextCompare)
		sql = Replace(sql , "{P5}" , sBrackfastC , vbTextCompare)
		sql = Replace(sql , "{P6}" , sBrackfastH , vbTextCompare)
		sql = Replace(sql , "{P7}" , Trim(arrItemCode(i)) , vbTextCompare)
		sql = Replace(sql , "{P8}" , Trim(arrBreakFast(i)) , vbTextCompare)
		sql = Replace(sql , "{P9}" , Trim(arrTake(0)) , vbTextCompare)
		sql = Replace(sql , "{P10}" , Trim(arrEtc(0)) , vbTextCompare)
		sql = Replace(sql , "{P11}" , Trim(arrManagerId(0)) , vbTextCompare)

		Call setQueryDB(sql)
		rw(sql)

		rw("<br>")
		'중식
		sql = "EXEC USP_MACHINE_CELSIUS_RECORD_DATA_MEAL_PROC '{P1}','{P2}','{P3}','{P4}','{P5}','{P6}','{P7}','{P8}','{P9}','{P10}','{P11}';"
		sql = Replace(sql , "{P1}" , Trim(arrItemIdx(i)) , vbTextCompare)
		sql = Replace(sql , "{P2}" , sCurrDate , vbTextCompare)
		sql = Replace(sql , "{P3}" , sAreaCode , vbTextCompare)
		sql = Replace(sql , "{P4}" , "2" , vbTextCompare)
		sql = Replace(sql , "{P5}" , sLunchC , vbTextCompare)
		sql = Replace(sql , "{P6}" , sLunchH , vbTextCompare)
		sql = Replace(sql , "{P7}" , Trim(arrItemCode(i)) , vbTextCompare)
		sql = Replace(sql , "{P8}" , Trim(arrLunch(i)) , vbTextCompare)
		sql = Replace(sql , "{P9}" , Trim(arrTake(1)) , vbTextCompare)
		sql = Replace(sql , "{P10}" , Trim(arrEtc(1)) , vbTextCompare)
		sql = Replace(sql , "{P11}" , Trim(arrManagerId(1)) , vbTextCompare)
		Call setQueryDB(sql)
		rw(sql)
		rw("<br>")

		'석식
		sql = "EXEC USP_MACHINE_CELSIUS_RECORD_DATA_MEAL_PROC '{P1}','{P2}','{P3}','{P4}','{P5}','{P6}','{P7}','{P8}','{P9}','{P10}','{P11}';"
		sql = Replace(sql , "{P1}" , Trim(arrItemIdx(i)) , vbTextCompare)
		sql = Replace(sql , "{P2}" , sCurrDate , vbTextCompare)
		sql = Replace(sql , "{P3}" , sAreaCode , vbTextCompare)
		sql = Replace(sql , "{P4}" , "3" , vbTextCompare)
		sql = Replace(sql , "{P5}" , sDinnerC , vbTextCompare)
		sql = Replace(sql , "{P6}" , sDinnerH , vbTextCompare)
		sql = Replace(sql , "{P7}" , Trim(arrItemCode(i)) , vbTextCompare)
		sql = Replace(sql , "{P8}" , Trim(arrDinner(i)) , vbTextCompare)
		sql = Replace(sql , "{P9}" , Trim(arrTake(2)) , vbTextCompare)
		sql = Replace(sql , "{P10}" , Trim(arrEtc(2)) , vbTextCompare)
		sql = Replace(sql , "{P11}" , Trim(arrManagerId(2)) , vbTextCompare)
		Call setQueryDB(sql)
		rw(sql)
		rw("<br>")
	Next 
	Call onlyAlert("success","처리완료 되었습니다.")
End Sub 

Sub Cart_Temperature_Approve_Process()
	Dim sql 
	sql = "EXEC AMD_APPROVE_PROC '{P1}','{P2}','{P3}','{P4}';" 
	sql = Replace(sql , "{P1}" , Replace(sCurrDate,"-","") , vbTextCompare)
	sql = Replace(sql , "{P2}" , "TEMPERATURE" , vbTextCompare)
	sql = Replace(sql , "{P3}" , Session("SNT_ADMIN_ID") , vbTextCompare)
	sql = Replace(sql , "{P4}" , sConfirm , vbTextCompare)
	Call setQueryDB(sql)
	Call gotoAlert("success","처리완료 되었습니다.","r","")
End Sub 


Sub Machine_Celsius_Confirm_Process()
	Dim sql
	sql = "EXEC USP_MACHINE_CELSIUS_RECORD_DATA_COMFIRM '{P1}','{P2}','{P3}';"
	sql = Replace(sql , "{P1}" , sCurrDate , vbTextCompare)
	sql = Replace(sql , "{P2}" , sAreaCode , vbTextCompare)
	sql = Replace(sql , "{P3}" , sConfirm , vbTextCompare)
	Call setQueryDB(sql)
	Call gotoAlert("success","처리완료 되었습니다.","r","")
End Sub 

Sub cart_temperature_Remove_Process()
	Dim sql 
	sql = "EXEC USP_MACHINE_CELSIUS_RECORD_DATA_DEL '{P1}','{P2}';"
	sql = Replace(sql , "{P1}" , sCurrDate , vbTextCompare)
	sql = Replace(sql , "{P2}" , sAreaCode , vbTextCompare)
	Call setQueryDB(sql)
	Call gotoAlert("success","처리완료 되었습니다.","m","./cart_temperature.asp")
End Sub 
%>