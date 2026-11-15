<meta charset="UTF-8" />
<meta http-equiv="X-UA-Compatible" content="IE=edge" />
<meta name="viewport" content="width=device-width, initial-scale=1.0" />
<meta name="description" content="CHA, Biotech, 차바이오텍, 분당차병원, 환자관리어드민" />
<meta name="keywords" content="CHA, Biotech, 차바이오텍, 분당차병원, 환자관리어드민" />
<title></title>

<!-- Fonts -->
<link rel="preload" href="/assets/fonts/Pretendard/Pretendard-Black.woff2" as="font" type="font/woff2" crossorigin />
<link rel="preload" href="/assets/fonts/Pretendard/Pretendard-Bold.woff2" as="font" type="font/woff2" crossorigin />
<link rel="preload" href="/assets/fonts/Pretendard/Pretendard-ExtraBold.woff2" as="font" type="font/woff2" crossorigin />
<link rel="preload" href="/assets/fonts/Pretendard/Pretendard-ExtraLight.woff2" as="font" type="font/woff2" crossorigin />
<link rel="preload" href="/assets/fonts/Pretendard/Pretendard-Light.woff2" as="font" type="font/woff2" crossorigin />
<link rel="preload" href="/assets/fonts/Pretendard/Pretendard-Medium.woff2" as="font" type="font/woff2" crossorigin />
<link rel="preload" href="/assets/fonts/Pretendard/Pretendard-Regular.woff2" as="font" type="font/woff2" crossorigin />
<link rel="preload" href="/assets/fonts/Pretendard/Pretendard-SemiBold.woff2" as="font" type="font/woff2" crossorigin />
<link rel="preload" href="/assets/fonts/Pretendard/Pretendard-Thin.woff2" as="font" type="font/woff2" crossorigin />

<!-- StyleSheet -->
<link rel="stylesheet" href="/assets/styles/index.css" type="text/css" />

<!-- JavaScript -->
<script defer src="/assets/js/index.js" type="module"></script>
<script src="/assets/js/jquery-1.7.2.min.js"></script>

<!-- SweetAlert -->
<script src="/js/sweetalert2.min.js"></script>
<link rel="stylesheet" href="/js/sweetalert2.min.css">

<script src="/js/jquery-1.7.2.min.js"></script>
<script src="/js/common.js?a=1"></script>
<style type="text/css">
	.LEDB {background-color:lime;display:block;}
</style>
<%
Dim vScriptName , vArrScriptName , vArrScript1 , vArrScript2 , vArrScript3 , vMenuChk  , vArrMenuChk
vScriptName = Request.ServerVariables("SCRIPT_NAME")

Select Case CInt(getStrCnt(vScriptName,"/"))
	Case 0  :   vScriptName = "////" & vScriptName
	Case 1  :   vScriptName = "///" & vScriptName 
	Case 2  :   vScriptName = "//" & vScriptName 
	Case 3  :   vScriptName = "/" & vScriptName 
End Select 

vArrScriptName  = Split(vScriptName,"/")   

vArrScript1     = vArrScriptName(2)     '메인 폴더
vArrScript2     = vArrScriptName(3)     '하위 폴더
vArrScript3     = vArrScriptName(4)     '파일 이름

'Response.write "<br/>vArrScript1:" & vArrScript1
'Response.write "<br/>vArrScript2:" & vArrScript2
'Response.write "<br/>vArrScript3:" & vArrScript3

Dim vAdminCompany , vAdminPosition

Select Case Session("SNT_ADMIN_COMPANY")
	Case "P09001"	: vAdminCompany = "CHA"
	Case "P09002"	: vAdminCompany = "HD"
End Select 

Select Case Session("SNT_ADMIN_POSITION")
	Case "P10001"	: vAdminPosition = "MANAGER"
	Case "P10002"	: vAdminPosition = "STAFF"
End Select 

%>

<script>
	
	$(document).ready(function() {

		var res = /^[0-9]+$/; 
		var resFloat = /^[0-9-.]+$/; 
		var resTime  = /^[0-9-:]+$/; 
		$(".number").keyup(function (e)
		{			
			if ($(this).val() !="")
			{
				if(!res.test($(this).val())) 
				{
					$(this).val("");
				}
			}							
		});	   

		$(".numberFloat").keyup(function (e)
		{			
			if ($(this).val() !="")
			{
				if(!resFloat.test($(this).val())) 
				{
					$(this).val("");
				}
			}							
		});	   

		$(".numberTime").keyup(function (e)
		{			
			if ($(this).val() !="")
			{
				if(!resTime.test($(this).val())) 
				{
					$(this).val("");
				}
			}							
		});	 



		//console.log("<%=vArrScript2%>" + ", " + "<%=vArrScript3%>");

		if ("<%=vArrScript2%>" == "setting") {
			if ("<%=vArrScript3%>" == "code.asp" || "<%=vArrScript3%>" == "user_add.asp" || "<%=vArrScript3%>" == "user_list.asp") {
				if ("<%=vArrScript3%>" == "user_list.asp" || "<%=vArrScript3%>" == "user_add.asp" ) {
					$(".gnb-list .gnb-item").eq(0).find(".lnb-list .lnb-item").eq(0).find("a").addClass("active");
				} else if ("<%=vArrScript3%>" == "code.asp") {
					$(".gnb-list .gnb-item").eq(0).find(".lnb-list .lnb-item").eq(1).find("a").addClass("active");
				}				
			}
			$(".gnb-list .gnb-item").eq(0).find("button").addClass("active");
		} else if ("<%=vArrScript2%>" == "checklist_setting") {
			if ("<%=vArrScript3%>" == "meal_cart.asp" || "<%=vArrScript3%>" == "meal_cart_register.asp") {
				$(".gnb-list .gnb-item").eq(1).find(".lnb-list .lnb-item").eq(0).find("a").addClass("active");
			} else if ("<%=vArrScript3%>" == "daily_meal.asp" || "<%=vArrScript3%>" == "daily_meal_register.asp") {
				$(".gnb-list .gnb-item").eq(1).find(".lnb-list .lnb-item").eq(1).find("a").addClass("active");
			} else if ("<%=vArrScript3%>" == "daily_meal_room.asp" || "<%=vArrScript3%>" == "daily_meal_room_register.asp") {
				$(".gnb-list .gnb-item").eq(1).find(".lnb-list .lnb-item").eq(2).find("a").addClass("active");
			} else if ("<%=vArrScript3%>" == "cart_temperature.asp" || "<%=vArrScript3%>" == "cart_temperature_register.asp") {
				$(".gnb-list .gnb-item").eq(1).find(".lnb-list .lnb-item").eq(3).find("a").addClass("active");
			} else if ("<%=vArrScript3%>" == "cart_equipment.asp" || "<%=vArrScript3%>" == "cart_equipment_register.asp") {
				$(".gnb-list .gnb-item").eq(1).find(".lnb-list .lnb-item").eq(4).find("a").addClass("active");
			}
			$(".gnb-list .gnb-item").eq(1).find("button").addClass("active");
		} else if ("<%=vArrScript2%>" == "checklist") {
			if ("<%=vArrScript3%>" == "meal_cart.asp" || "<%=vArrScript3%>" == "meal_cart_register.asp") {
				$(".gnb-list .gnb-item").eq(2).find(".lnb-list .lnb-item").eq(0).find("a").addClass("active");
			} else if ("<%=vArrScript3%>" == "daily_meal.asp" || "<%=vArrScript3%>" == "daily_meal_register.asp") {
				$(".gnb-list .gnb-item").eq(2).find(".lnb-list .lnb-item").eq(1).find("a").addClass("active");
			} else if ("<%=vArrScript3%>" == "cart_temperature.asp" || "<%=vArrScript3%>" == "cart_temperature_register.asp") {
				$(".gnb-list .gnb-item").eq(2).find(".lnb-list .lnb-item").eq(2).find("a").addClass("active");
			} else if ("<%=vArrScript3%>" == "cart_equipment.asp" || "<%=vArrScript3%>" == "cart_equipment_register.asp") {
				$(".gnb-list .gnb-item").eq(2).find(".lnb-list .lnb-item").eq(3).find("a").addClass("active");
			}
			$(".gnb-list .gnb-item").eq(2).find("button").addClass("active");
		}		
	});
	
</script>