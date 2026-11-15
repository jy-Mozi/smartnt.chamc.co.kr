<%@CodePage=65001%>
<% 
'=============================================
' 작 성 자 : 노 광 옹
' 작 성 일 : 2025-03-21
' 파일내용 : 기기설비점검일지 출력
'=============================================
Option Explicit 
%>
<!-- #include virtual = "/common/library/common.asp"-->
<!doctype html>
<html lang="ko">
  <head>
	<!-- #include virtual = "/common/include/layout_header.asp"-->
	<style>
		@media print {
			.noprint {display:none;}
			.print {display:block;}
		}
		@page {
			size: landscape;
		}

		.section {
			border-radius: 0;
			box-shadow: none;
			h2 {
				font-weight: 800;
				font-size: 22px;
			}
			
			.table {
				overflow: hidden;
				display: flex;
				justify-content: end;

				& table {
					& th, & td {
						border: 1px solid var(--color-gray-01);
					}
					
					th, td {
						padding: 5px;
						text-wrap: initial;
					}
				}
			}
			
			.head_container {
				.table {
					table {
						width: 35%;
					}
				}
			}

			.body_container {
				.table {
					margin-top: 30px;
					table {
						border: 3px solid var(--color-gray-01);
					}
				}
			}
			    
			
			.date-picker {
				margin-top: 20px;
				display: flex;
				gap: 10px;
				.input {
					&.date {
						width: 180px;
					}
				}
			}
			
			.text_blue {
				color: var(--color-sub-02);
			}
			.text_red {
				color: var(--color-red);
			}

			.left {
				text-align: left;
			}
			.hide {
				display: none;
			}
			
		}

		.bx-calendar {
			position: absolute;
			right: 10px;
			pointer-events: none;
		}

	</style>
 <%
	Dim selDate , arrData , arrRecord , sql , rs , i , j , printerYn, selHoType
	selDate		= nullFillWith(Request("selDate"),Date(),False)
	selHoType	= nullFillWith(Request("selHoType"),"",False)
	printerYn = "Y"

	'sql = "EXEC ADM_MACHINE_CELSIUS_RECORD_LIST_PAGEING_V1 '{P1}','{P2}','{P3}','{P4}','{P5}','{P6}','{P7}';"
	'sql = Replace(sql , "{P1}", "1"		, vbTextCompare)
	'sql = Replace(sql , "{P2}", "1000"	, vbTextCompare)
	'sql = Replace(sql , "{P3}", "CD"		, vbTextCompare)
	'sql = Replace(sql , "{P4}", ""	, vbTextCompare)
	'sql = Replace(sql , "{P5}", selDate , vbTextCompare)
	'sql = Replace(sql , "{P6}", selDate , vbTextCompare)
	'sql = Replace(sql , "{P7}", selHoType	, vbTextCompare)


	'Set rs = getQueryDB(sql)
	'If Not(rs.Eof Or rs.Bof) Then 
		'arrData = rs.GetRows()
	'End If 
	'rsClose(rs)

	Dim AP_CHA_MANAGER_NM , AP_CHA_STAFF_NM , AP_HD_MANAGER_NM , AP_HD_STAFF_NM
	Dim AP_CHA_MANAGER_ID , AP_CHA_STAFF_ID , AP_HD_MANAGER_ID , AP_HD_STAFF_ID
	sql = "EXEC ADM_APPROVE_DETAIL_VIEW '{P1}','{P2}','{P3}','{P4}';"
	sql = Replace(sql , "{P1}" , selDate			, vbTextCompare)
	sql = Replace(sql , "{P2}" , "TEMPERATURE"		, vbTextCompare)
	sql = Replace(sql , "{P3}" , selHoType			, vbTextCompare)
	sql = Replace(sql , "{P4}" , ""					, vbTextCompare)


	Set rs = getQueryDB(sql)
	If Not(rs.Eof Or rs.Bof) Then 
		AP_CHA_MANAGER_NM	 = nullFillWith(rs("CHA_MANAGER_NAME"),"",False)
		AP_CHA_STAFF_NM	 = nullFillWith(rs("CHA_STAFF_NAME"),"",False)
		AP_HD_MANAGER_NM	 = nullFillWith(rs("HD_MANAGER_NAME"),"",False)
		AP_HD_STAFF_NM		 = nullFillWith(rs("HD_STAFF_NAME"),"",False)
		AP_CHA_MANAGER_ID	= nullFillWith(rs("CHA_MANAGER"),"",False)
		AP_CHA_STAFF_ID		= nullFillWith(rs("CHA_STAFF"),"",False)
		AP_HD_MANAGER_ID	= nullFillWith(rs("HD_MANAGER"),"",False)
		AP_HD_STAFF_ID		= nullFillWith(rs("HD_STAFF"),"",False)
	Else 
		AP_CHA_MANAGER_NM	= ""
		AP_CHA_STAFF_NM		= ""
		AP_HD_MANAGER_NM	= ""
		AP_HD_STAFF_NM		= ""
		AP_CHA_MANAGER_ID	= ""
		AP_CHA_STAFF_ID		= ""
		AP_HD_MANAGER_ID	= ""
		AP_HD_STAFF_ID		= ""
	End If 
	rsClose(rs)
 %>
  </head>
  <body>
    <form class="form" name="dataForm" id="dataForm">
	<input type="hidden" name="command" id="command" value="Approve" title="명령" style="background-color:lime;">
	<input type="hidden" name="FLAG" id="FLAG" value="TEMPERATURE" title="플래그값" style="background-color:lime;">
	<input type="hidden" name="approveGrade" id="approveGrade" value="" title="결재" style="background-color:lime;">
	<input type="hidden" name="hospitalType" id="hospitalType" title="병원구분" style="background-color:lime;">
	<section class="section">
		<div class="head_container">
			<div style="display: flex;justify-content: space-between;">
			<h2>온/습도 점검일지</h2>
				<%
				If AP_CHA_MANAGER_NM = "" Or AP_CHA_STAFF_NM = "" Or AP_HD_MANAGER_NM = "" Or AP_HD_STAFF_NM = "" Then 
					printerYn = "N"
				End If 
				printerYn = "Y"
				If printerYn = "Y" Then %>
				<button class="button noprint" id="btnPrint" onClick="javascript:reportPrint();void(0);" style="width:100px;padding: 0;height: 100%;">출력하기</button>
				<% End If %>
			</div>
			<div class="table">
				<table>
					<tbody>
						<tr>
							<td rowspan="3">결재</td>
							<td colspan="2">현대그린푸드</td>
							<td colspan="2">본원</td>
						</tr>	
						<tr>
							<td>담당</td>
							<td>점장</td>
							<td>담당</td>
							<td>팀장</td>
						</tr>
						<tr>
							<td style="height:30px;">
							<%
							If AP_HD_STAFF_NM = "" And vAdminCompany = "HD" And vAdminPosition = "STAFF" Then 
								%><small class="badge green noprint" onClick="javascript:btnApprove('1');void(0);" style="cursor:pointer;">승인</small><%
							End If 

							If AP_HD_STAFF_NM <> "" Then 
								rw(AP_HD_STAFF_NM)
								If Session("SNT_ADMIN_ID") = AP_HD_STAFF_ID Then 
									rw("<BR/><small class=""badge red noprint"" onClick=""javascript:btnApproveCancel('1');void(0);"" style=""cursor:pointer;"">취소</small>")
								End If
							End If 
							%>
							</td>
							<td>
							<%
							If AP_HD_MANAGER_NM = "" And vAdminCompany = "HD" And vAdminPosition = "MANAGER" Then 
								%><small class="badge orange noprint" onClick="javascript:btnApprove('2');void(0);" style="cursor:pointer;">승인</small><%
							End If 

							If AP_HD_MANAGER_NM <> "" Then 
								rw(AP_HD_MANAGER_NM)
								If Session("SNT_ADMIN_ID") = AP_HD_MANAGER_ID Then 
									rw("<BR/><small class=""badge red noprint"" onClick=""javascript:btnApproveCancel('2');void(0);"" style=""cursor:pointer;"">취소</small>")
								End If
							End If 
							%>
							</td>
							<td>
							<%
							If AP_CHA_STAFF_NM = "" And vAdminCompany = "CHA" And vAdminPosition = "STAFF"Then 
								%><small class="badge orange noprint" onClick="javascript:btnApprove('3');void(0);" style="cursor:pointer;">승인</small><%
							End If 
							If AP_CHA_STAFF_NM <> "" Then 
								rw(AP_CHA_STAFF_NM)
								If Session("SNT_ADMIN_ID") = AP_CHA_STAFF_ID Then 
									rw("<BR/><small class=""badge red noprint"" onClick=""javascript:btnApproveCancel('3');void(0);"" style=""cursor:pointer;"">취소</small>")
								End If
							End If 
							%>
							</td>
							<td>
							<%
							If AP_CHA_MANAGER_NM = "" And vAdminCompany = "CHA" And vAdminPosition = "MANAGER" Then 
							%><small class="badge orange noprint" onClick="javascript:btnApprove('4');void(0);"" style="cursor:pointer;">승인</small><% 
							End If 
							If AP_CHA_MANAGER_NM <> "" Then 
								rw(AP_CHA_MANAGER_NM)
								If Session("SNT_ADMIN_ID") = AP_CHA_MANAGER_ID Then 
									rw("<BR/><small class=""badge red noprint"" onClick=""javascript:btnApproveCancel('4');void(0);"" style=""cursor:pointer;"">취소</small>")
								End If
							End If 
							%>
							</td>
					</tbody>
				</table>
			</div>
	<div style="display:flex;gap:10px;">
			<div class="select noprint"  style="width:200px;height:100%;margin-top:20px;">
				<select id="selHoType" name="selHoType">
					<option value="">전체</option>
					<%=getCodeDB("select" , "P11" , selHoType)%>			
				</select>
			</div>
			<div class="date-picker">
				<div class="input date noprint">
					<input type="text" name="selDate" id="selDate" placeholder="YYYY-MM-DD" readonly value="<%=selDate%>"/>
					<i class="bx bx-calendar"></i>
				</div>
			</div>
			<span id="printtext" class="print"></span>
		</div>
	</div>
		<div class="body_container">
			<div class="table">
				<table>
					<colgroup>
						<col width="11%">
						<col width="11%">
						<col width="11%">
						<col width="11%">
						<col width="11%">
						<col width="11%">
						<col width="11%">
						<col width="11%">
						<col width="11%">
					</colgroup>
					<thead>
						<tr>
							<th rowspan="2">구역</th>
							<th rowspan="2">구분</th>
							<th colspan="2">조식</th>
							<th colspan="2">중식</th>
							<th colspan="2">석식</th>
							<th rowspan="2">특이사항</th>
						</tr>
						<tr>
							<th>온도</th>
							<th>점검자</th>
							<th>온도</th>
							<th>점검자</th>
							<th>온도</th>
							<th>점검자</th>
						</tr>
					</thead>
					<tbody>
					<%
					Dim CP_CHK_IDX , CP_CNT , CP_AREA_CODE , CP_AREA_NAME , CP_HOSPITAL_TYPE , CP_CODE_NM  
					Dim CP_STAFF1_NAME , CP_STAFF2_NAME , CP_STAFF3_NAME , CP_CURR_YMD
					Dim CP_AREA_BREAKFAST_C , CP_AREA_BREAKFAST_H , CP_AREA_LUNCH_C , CP_AREA_LUNCH_H , CP_AREA_DINNER_C, CP_AREA_DINNER_H
					Dim CP_ITEM_BREAKFAST_C , CP_ITEM_LUNCH_C , CP_ITEM_DINNER_C
					Dim CP_ETC1 , CP_ETC2 , CP_ETC3 , STR_TOT_ETC, CP_ITEM_CODE
					sql = "EXEC ADM_MACHINE_CELSIUS_RECORD_LIST_PRINT '{P1}','{P2}';"
					sql = Replace(sql , "{P1}" , selHoType , vbTextCompare)
					sql = Replace(sql , "{P2}" , selDate , vbTextCompare)

					Set rs = getQueryDB(sql)

					If Not(rs.Eof Or rs.Bof) Then 
						Do Until rs.Eof 
							CP_CHK_IDX		= rs("CHK_IDX")
							CP_CNT			= rs("CNT")
							CP_AREA_CODE	= rs("AREA_CODE")
							CP_AREA_NAME	= rs("AREA_NAME")
							CP_HOSPITAL_TYPE	= rs("HOSPITAL_TYPE")
							CP_CODE_NM		= rs("CODE_NM")
							CP_STAFF1_NAME	= rs("STAFF1_NAME")
							CP_STAFF2_NAME	= rs("STAFF2_NAME")
							CP_STAFF3_NAME	= rs("STAFF3_NAME")
							CP_CURR_YMD		= rs("CURR_YMD")
							CP_ITEM_CODE	= rs("ITEM_CODE")

							CP_AREA_BREAKFAST_C	= rs("AREA_BREAKFAST_C")
							CP_AREA_BREAKFAST_H	= rs("AREA_BREAKFAST_H")
							CP_AREA_LUNCH_C		= rs("AREA_LUNCH_C")
							CP_AREA_LUNCH_H		= rs("AREA_LUNCH_H")
							CP_AREA_DINNER_C	= rs("AREA_DINNER_C")	
							CP_AREA_DINNER_H	= rs("AREA_DINNER_H")

							CP_ITEM_BREAKFAST_C	= rs("ITEM_BREAKFAST_C")
							CP_ITEM_LUNCH_C		= rs("ITEM_LUNCH_C")
							CP_ITEM_DINNER_C	= rs("ITEM_DINNER_C")
							CP_ETC1		= rs("ETC1")
							CP_ETC2		= rs("ETC2")
							CP_ETC3		= rs("ETC3")


					%>
						<% 
						If CP_CHK_IDX  = "1" Then 
							STR_TOT_ETC = CP_ETC1 & CP_ETC2 & CP_ETC3							
							STR_TOT_ETC	= Replace(STR_TOT_ETC , "없음" , "")
							If CP_ITEM_CODE = "P04024" Then 
								CP_CNT = "0"
							End If 
						%>
					<tr>
						<td rowspan="<%=CInt(CP_CNT)+1%>"><%=CP_AREA_NAME%><BR/>(<%=CP_AREA_CODE%>)</td>
						<td>온도 / 습도</td>
						<td><%=CP_AREA_BREAKFAST_C%>℃  / <%=CP_AREA_BREAKFAST_H%>%</td>
						<td rowspan="<%=CInt(CP_CNT)+1%>"><%=CP_STAFF1_NAME%></td>
						<td><%=CP_AREA_LUNCH_C%>℃  / <%=CP_AREA_LUNCH_H%>%</td>
						<td rowspan="<%=CInt(CP_CNT)+1%>"><%=CP_STAFF2_NAME%></td>
						<td><%=CP_AREA_DINNER_C%>℃  / <%=CP_AREA_DINNER_H%>%</td>
						<td rowspan="<%=CInt(CP_CNT)+1%>"><%=CP_STAFF3_NAME%></td>
						<td rowspan="<%=CInt(CP_CNT)+1%>"><%=STR_TOT_ETC%></td>
					</tr>
					
					<tr <%If CP_ITEM_CODE = "P04024" Then rw(" style='display:none;'") End If %>>
						<td><%=CP_CODE_NM%></td>
						<td><%=CP_ITEM_BREAKFAST_C%>℃</td>
						<td><%=CP_ITEM_LUNCH_C%>℃</td>
						<td><%=CP_ITEM_DINNER_C%>℃</td>
					</tr>
						<% 
							STR_TOT_ETC = ""
						Else  
						%>
					<tr>
						<td><%=CP_CODE_NM%></td>
						<td><%=CP_ITEM_BREAKFAST_C%>℃</td>
						<td><%=CP_ITEM_LUNCH_C%>℃</td>
						<td><%=CP_ITEM_DINNER_C%>℃</td>
					</tr>
					<% End If %>
					<%
							rs.MoveNext
						Loop 
					End If 
					rsClose(rs)
					%>
					</tbody>
				</table>
			</div>
		</div>
	</section>
	</form>
  </body>
</html>
<script type="text/javascript">
<!--
$(document).ready(function(){
	const elem = document.getElementById('selDate');
	const rangepicker = new DateRangePicker(elem, {});

	elem.addEventListener('changeDate', function(e) {
	  //console.log(e.detail.date); // new Date() 찍어봤을때랑 똑같은 형태로 나옴
		var selDate = e.detail.date;
		var year = selDate.getFullYear();
		var month = selDate.getMonth() + 1;
		var day = selDate.getDate();
		//console.log(year + "-" + month + "-" + day);
		btnSubmit();
	});

	$("#selHoType").change(function(){
		btnSubmit();
	});
});


function btnSubmit() {
	var frm = document.dataForm;
	
	frm.action = "./cart_temperature_print.asp";
	frm.target = "_self";
	frm.method = "POST";
	frm.submit();		
}

function btnApprove(v) {
	var frm = document.dataForm;
	frm.approveGrade.value = v;
	if(!setInputAlert(frm.selDate, "날짜를")) return false;
	if(!setInputAlert(frm.selHoType, "병원구분을")) return false;
	frm.hospitalType.value = $("#selHoType").val();
	frm.command.value = "Approve";
	frm.action = "./approve_proc.asp";			
	frm.target = "procFrame";
	frm.method = "POST";
	frm.submit();	
}

function btnApproveCancel(v) {
	var frm = document.dataForm;
	frm.approveGrade.value = v;
	if(!setInputAlert(frm.selDate, "날짜를")) return false;
	if(!setInputAlert(frm.selHoType, "병원구분을")) return false;
	frm.hospitalType.value = $("#selHoType").val();
	frm.command.value = "ApproveCancel";
	frm.action = "./approve_proc.asp";			
	frm.target = "procFrame";
	frm.method = "POST";
	frm.submit();	
}

function reportPrint() {
	var text = "";
	var seldate = $("#selDate").val();
	seldate = seldate.split("-");

	
	text = "점검일시 : " + seldate[0] + "년 " + seldate[1] + "월 " + seldate[2] + "일";
	$("#printtext").text(text);
	window.print();
}

//-->
</script>
<iframe  name="procFrame" id="procFrame" style="width:0%; height:500px;display:none;" class="LEDB"></iframe>
