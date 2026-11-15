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
	Dim fmDate , toDate, selHoType
	fmDate		= nullFillWith(Request("fmDate"),Date(),False)
	toDate		= nullFillWith(Request("toDate"),Date(),False)
	selHoType	= nullFillWith(Request("selHoType"),"",False)
	Dim printerYn	:	printerYn = "Y"

	Dim AP_SEQ_IDX , AP_ITEM_CNT , AP_AREA_CODE , AP_AREA_NAME , AP_ITEM_LIST , AP_ITEM_NAME , AP_HOSPITAL_TYPE , arrItem
	Dim sql , rs , i , j , t 
	sql = "EXEC USP_MECHINE_FACILITY_LIST_PRINT '{P1}','{P2}';"
	sql = Replace(sql , "{P1}", selHoType		, vbTextCompare)
	sql = Replace(sql , "{P2}", fmDate		, vbTextCompare)
	'rw(sql)
	Set rs = getQueryDB(sql)
	If Not(rs.Eof Or rs.Bof) Then 
		arrItem = rs.GetRows()
	End If 
	rsClose(rs)




	Dim AP_CHA_MANAGER_NM , AP_CHA_STAFF_NM , AP_HD_MANAGER_NM , AP_HD_STAFF_NM
	Dim AP_CHA_MANAGER_ID , AP_CHA_STAFF_ID , AP_HD_MANAGER_ID , AP_HD_STAFF_ID
	sql = "EXEC ADM_APPROVE_DETAIL_VIEW '{P1}','{P2}','{P3}','{P4}';"
	sql = Replace(sql , "{P1}" , Replace(fmDate,"-","") , vbTextCompare)
	sql = Replace(sql , "{P2}" , "MACHINE"    , vbTextCompare)
	sql = Replace(sql , "{P3}" , selHoType    , vbTextCompare)
	sql = Replace(sql , "{P4}" , ""    , vbTextCompare)
	
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
	<input type="hidden" name="FLAG" id="FLAG" value="MACHINE" title="플래그값" style="background-color:lime;">
	<input type="hidden" name="selDate" id="selDate" value="<%=Replace(fmDate,"-","")%>" title="플래그값" style="background-color:lime;">
	<input type="hidden" name="approveGrade" id="approveGrade" value="" title="결재" style="background-color:lime;">
	<input type="hidden" name="hospitalType" id="hospitalType" title="병원구분" style="background-color:lime;">
	<section class="section">
		<div class="head_container">
			<div style="display: flex;justify-content: space-between;">
				<h2>기기설비점검일지</h2>
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
						</tr>
					</tbody>
				</table>
			</div>

		<div style="display:flex;gap:20px;border:0px solid lime;">
			<div class="select noprint"  style="width:200px;margin-top:20px;height:100%;">
				<select id="selHoType" name="selHoType">
					<option value="">전체</option>
					<%=getCodeDB("select" , "P11" , selHoType)%>			
				</select>
			</div>
			<div class="date-picker">
				<div class="input date noprint">
					<input type="text" name="fmDate" id="fmDate" placeholder="YYYY-MM-DD" readonly value="<%=fmDate%>"/>
					<i class="bx bx-calendar"></i>
				</div>
			</div>
				<div class="input noprint" style="width:200px;margin-top:20px;display:none;">
					<input type="text" name="toDate" id="toDate" placeholder="YYYY-MM-DD" readonly value="<%=toDate%>"/>
					<i class="bx bx-calendar"></i>
				</div>
			</div>
			<span id="printtext" class="print"></span>
		</div>
		<div class="body_container">
			<div class="table">
				<table>
					<colgroup>
						<col width="10%">
						<col width="10%">
						<col width="10%">
						<col width="10%">
						<col width="10%">
						<col width="10%">
						<col width="10%">
						<col width="10%">
						<col width="10%">
					</colgroup>
					<thead>
						<tr>
							<th>구역</th>
							<th>구분</th>
							<% 
							Dim currDate  : currDate = fmDate
							For i = 0 To 6 								
							%>
								<th><%=currDate%>(<%=strWeek(convertDate(currDate,"-"))%>)</th>
							<% 
								currDate = dateAdd("d" , 1 , currDate)
							Next %>
						</tr>

					</thead>
					<tbody>
						<%
						If isArray(arrItem) Then 
							For j = 0 To Ubound(arrItem , 2)
								AP_SEQ_IDX			=	arrItem(0,j)
								AP_ITEM_CNT			=	arrItem(1,j) 
								AP_AREA_CODE		=	arrItem(2,j)
								AP_AREA_NAME		=	arrItem(3,j)
								AP_ITEM_LIST		=	arrItem(4,j)
								AP_HOSPITAL_TYPE	=	arrItem(5,j)
								AP_ITEM_NAME		=	arrItem(6,j)																
						%>
						<tr>
							<% If AP_SEQ_IDX = "1" Then %>
							<td rowspan="<%=AP_ITEM_CNT%>"><%=AP_AREA_NAME%><BR/>(<%=AP_AREA_CODE%>)</td>
							<% End If %>

							<td><%=AP_ITEM_NAME%></td>
							<%
								sql = "EXEC USP_MECHINE_FACILITY_LIST_ITEM_PRINT '{P1}','{P2}','{P3}';"
								sql = Replace(sql , "{P1}" , AP_AREA_CODE , vbTextCompare)
								sql = Replace(sql , "{P2}" , AP_ITEM_LIST , vbTextCompare)
								sql = Replace(sql , "{P3}" , fmDate , vbTextCompare)				
'								rw(sql)
								Set rs = getQueryDB(sql)
								If Not(rs.Eof Or rs.Bof) Then 
									Do Until rs.Eof 
							%><td <%=txtClass(nullFillWith(rs(5),"",False))%>><%=getCode("name","CHECK_FLAG",nullFillWith(rs(5),"",False))%></td><%
										rs.MoveNext
									Loop 
								End If 
								rsClose(rs)
							%>
						</tr>
						<%
							Next
						End If 
						
						Dim STR_ETC
						STR_ETC = ""
						sql = "EXEC USP_MECHINE_FACILITY_LIST_ETC_PRINT '{P1}','{P2}';"
						sql = Replace(sql , "{P1}" , selHoType , vbTextCompare)
						sql = Replace(sql , "{P2}" , fmDate , vbTextCompare)
						'RW(SQL)
						Set rs  = getQueryDB(sql)
						If Not(rs.Eof Or rs.Bof) Then 
							Do Until rs.Eof
								STR_ETC = STR_ETC & rs(0) & rs(1)
								STR_ETC	= Replace(STR_ETC , "없음" , "")
								rs.MoveNext
							Loop 
						End If 
						rw("<tr><td colspan=""2"">조치사항</td><td colspan=""7"">" & STR_ETC & "</td></tr>")
						%>

					</tbody>
				</table>
			</div>
		</div>
	</section>

	</form>
  </body>
</html>

<%
Function strWeek(v) 
	Dim week , cWeek
	If nullFillWith(v,"",False) = "" Then 
		cWeek = v
	Else 
		week = weekDay(v)
		
		Select Case week
			Case "1"	: cWeek = "일"
			Case "2"	: cWeek = "월"
			Case "3"	: cWeek = "화"
			Case "4"	: cWeek = "수"
			Case "5"	: cWeek = "목"
			Case "6"	: cWeek = "금"
			Case "7"	: cWeek = "토"
		End Select 
	End If 


	strWeek = cWeek
End Function 

Function txtClass(t) 
	Dim strTxt
	Select Case t
		Case "1" 
			strTxt = "class=""text_blue"""
		Case "2"
			strTxt = ""
		Case "3" 
			strTxt = "class=""text_red"""
		Case Else 
			strTxt = ""
	End Select 
	txtClass = strTxt
End Function 
%>

<script type="text/javascript">

$(document).ready(function(){
	const elem = document.getElementById('fmDate');
	const rangepicker = new DateRangePicker(elem, {});

	elem.addEventListener('changeDate', function(e) {
	  //console.log(e.detail.date); // new Date() 찍어봤을때랑 똑같은 형태로 나옴
		var selDate = e.detail.date;
		var year = selDate.getFullYear();
		var month = selDate.getMonth() + 1;
		month = (month.toString().padStart(2, "0"));
		var day = selDate.getDate() ;
		//console.log(year + "-" + month + "-" + day);
		//alert(year + "-" + month + "-" + day);
		//alert(dateAdd(year + "-" + month + "-" + day , 6 , "d"));
		$("#toDate").val(dateAdd(year + "-" + month + "-" + day , 6 , "d"));
		btnSubmit();
	});

	$("#selHoType").change(function(){
		btnSubmit();
	});
});

function btnSubmit() {
	var frm = document.dataForm;
	
	frm.action = "./cart_equipment_print.asp";
	frm.target = "_self";
	frm.method = "POST";
	frm.submit();		
}



function reportPrint() {
	var text = "";
	var fmdate = $("#fmDate").val();
	var todate = $("#toDate").val();
	fmdate = fmdate.split("-");
	todate = todate.split("-");
	
	text = "점검일시 : " + fmdate[0] + "년 " + fmdate[1] + "월 " + fmdate[2] + "일 ~ " + todate[0] + "년 " + todate[1] + "월 " + todate[2] + "일";
	$("#printtext").text(text);
	window.print();
}

function btnApprove(v) {
	var frm = document.dataForm;
	frm.approveGrade.value = v;
	if(!setInputAlert(frm.fmDate, "날짜를")) return false;
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
	if(!setInputAlert(frm.fmDate, "날짜를")) return false;
	if(!setInputAlert(frm.selHoType, "병원구분을")) return false;
	frm.hospitalType.value = $("#selHoType").val();
	frm.command.value = "ApproveCancel";
	frm.action = "./approve_proc.asp";			
	frm.target = "procFrame";
	frm.method = "POST";
	frm.submit();	
}


window.onload = function() {		
	const datePickerElement = document.querySelector('.datepicker-input');
			
	// 기존 Datepicker 인스턴스가 있는지 확인하고 제거
	if (datePickerElement.datepicker) {
		datePickerElement.datepicker.destroy();
	}

	// datepicker 초기화
	const datepicker = new Datepicker(datePickerElement, {
		language: 'ko',
		autohide: true,
		orientation: 'left',
//			format: 'yyyy-mm-dd', // 입력 필드에 표시될 날짜 형식
		beforeShowDay: (date) => {
			const day = date.getDay();
			return {
				enabled: day === 1, // 월요일만 활성화
				classes: day === 1 ? 'monday-date' : 'disabled-date'
			};
		}
	});

	datePickerElement.datepicker = datepicker;
}

</script>

<iframe  name="procFrame" id="procFrame" style="width:0%; height:500px;display:none;" class="LEDB"></iframe>
