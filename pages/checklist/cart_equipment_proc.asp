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


Execute("Call "& sCommand &"_Process()")    '함수 호츨

Sub machine_equipment_Data_Process()
	Dim sCommand, sConfirm, sSelRegId, sIdx, sSelResult, sSelEtc, sPage, sCurrDate, sItemIdx, sAreaCode, sItemCode, sEtc
	Dim sql , rs , i 
	sCommand	= nullFillWith(Request("command"), "", False)
	sConfirm	= nullFillWith(Request("confirm"), "", False)
	sSelRegId	= nullFillWith(Request("selRegId"), "", False)
	sIdx		= nullFillWith(Request("idx"), "", False)
	sSelResult	= nullFillWith(Request("selResult"), "", False)
	sSelEtc		= nullFillWith(Request("selEtc"), "", False)
	sPage		= nullFillWith(Request("page"), "", False)
	sCurrDate	= nullFillWith(Request("currDate"), "", False)
	sItemIdx	= nullFillWith(Request("itemIdx"), "", False)
	sAreaCode	= nullFillWith(Request("areaCode"), "", False)
	sItemCode	= nullFillWith(Request("itemCode"), "", False)
	sEtc		= nullFillWith(Request("etc"), "", False)
	Dim arrItemIdx , arrItemCode , arrItemResult
	sItemIdx	= Replace(sItemIdx , Space(1) , "")
	sItemCode	= Replace(sItemCode , Space(1) , "")
	sSelResult	= Replace(sSelResult , Space(1) , "")

	arrItemIdx		= Split(sItemIdx,",")	
	arrItemCode		= Split(sItemCode,",")
	arrItemResult	= Split(sSelResult,",")

	For i = 0 To UBound(arrItemIdx)
		sql = "EXEC USP_MACHINE_FACILITY_RECORD_DATA_MEAL_PROC '{P1}','{P2}','{P3}','{P4}','{P5}','{P6}','{P7}','{P8}';"
		sql = Replace(sql , "{P1}" , Trim(arrItemIdx(i)) , vbTextCompare)
		sql = Replace(sql , "{P2}" , sCurrDate , vbTextCompare)
		sql = Replace(sql , "{P3}" , sAreaCode , vbTextCompare)
		sql = Replace(sql , "{P4}" , Trim(arrItemCode(i)) , vbTextCompare)
		sql = Replace(sql , "{P5}" , Trim(arrItemResult(i)) , vbTextCompare)
		sql = Replace(sql , "{P6}" , sSelEtc , vbTextCompare)
		sql = Replace(sql , "{P7}" , sEtc , vbTextCompare)
		sql = Replace(sql , "{P8}" , sSelRegId , vbTextCompare)

		Call setQueryDB(sql)
   	Next 
		
	Call onlyAlert("success","처리완료 되었습니다.")
End Sub 

Sub Machine_Equipment_Approve_Process()
	Dim sql 
	sql = "EXEC AMD_APPROVE_PROC '{P1}','{P2}','{P3}','{P4}';" 
	sql = Replace(sql , "{P1}" , Replace(sCurrDate,"-","") , vbTextCompare)
	sql = Replace(sql , "{P2}" , "MACHINE" , vbTextCompare)
	sql = Replace(sql , "{P3}" , Session("SNT_ADMIN_ID") , vbTextCompare)
	sql = Replace(sql , "{P4}" , sConfirm , vbTextCompare)
	Call setQueryDB(sql)
	Call gotoAlert("success","처리완료 되었습니다.","r","")
End Sub 

Sub Machine_Equipment_Confirm_Process()
	Dim sql
	sql = "EXEC USP_MACHINE_FACILITY_RECORD_DATA_COMFIRM '{P1}','{P2}','{P3}';"
	sql = Replace(sql , "{P1}" , sCurrDate , vbTextCompare)
	sql = Replace(sql , "{P2}" , sAreaCode , vbTextCompare)
	sql = Replace(sql , "{P3}" , sConfirm , vbTextCompare)
	Call setQueryDB(sql)
	Call gotoAlert("success","처리완료 되었습니다.","r","")
End Sub 

Sub cart_equipment_Remove_Process()
	Dim sql
	sql = "EXEC USP_MACHINE_FACILITY_RECORD_DATA_DEL '{P1}','{P2}';"
	sql = Replace(sql , "{P1}" , sCurrDate , vbTextCompare)
	sql = Replace(sql , "{P2}" , sAreaCode , vbTextCompare)
	Call setQueryDB(sql)
	Call gotoAlert("success","처리완료 되었습니다.","m","./cart_equipment.asp")
End Sub 
%>