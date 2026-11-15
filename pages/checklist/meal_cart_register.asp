<%@CodePage=65001%>
<% 
'=============================================
' 작 성 자 : 노 광 옹
' 작 성 일 : 2024-10-17
' 파일내용 : 
'=============================================
Option Explicit 
%>
<!-- #include virtual = "/common/library/common.asp"-->
<!doctype html>
<html lang="ko">
  <head>
    <!-- #include virtual = "/common/include/layout_header.asp"-->
	<%
	Dim sCommand, sIdx, sCart, sFmDate, sPage, sToDate, sCartCode, sCurrDate
	sCommand = nullFillWith(Request("command"), "", False)
	sIdx = nullFillWith(Request("idx"), "", False)
	sCart = nullFillWith(Request("cart"), "", False)
	sFmDate = nullFillWith(Request("fmDate"), "", False)
	sPage = nullFillWith(Request("page"), "", False)
	sToDate = nullFillWith(Request("toDate"), "", False)
	sCurrDate = nullFillWith(Request("currDate"), Date() , False)
	sCartCode = nullFillWith(Request("cartCode"), "", False)

	Dim rs, sql , i , j , arrData, arrCart, arrComment,arrManager
	Dim arrComplaint(2) , arrEtc(2) , arrStaff(2) , arrRegDate(2)
	Dim CP_IDX, CP_CART_CODE, CP_CART_NAME
	sql = "EXEC USP_CART_LIST '';"
	Set rs = getQueryDB(sql)
	If Not(rs.Eof Or rs.Bof) Then 
		arrCart = rs.GetRows()
	End If 
	rsClose(rs)

	Dim VP_IDX, VP_CODE_NM,  VP_CART_VAL
	sql = "EXEC USP_CODE_COMMON_LIST 'P06';"
	Set rs = getQueryDB(sql)
	If Not(rs.Eof Or rs.Bof) Then 
		arrComment = rs.GetRows()
	End If 
	rsClose(rs)

	sql = "EXEC USP_ADMIN_LIST_PAGEING '{P1}','{P2}','{P3}','{P4}','{P5}','','';"
	sql = Replace(sql , "{P1}", "1" , vbTextCompare)
	sql = Replace(sql , "{P2}", "1000" , vbTextCompare)
	sql = Replace(sql , "{P3}", "" , vbTextCompare)
	sql = Replace(sql , "{P4}", "" , vbTextCompare)
	sql = Replace(sql , "{P5}", "" , vbTextCompare)
	Set rs = getQueryDB(sql)
	If Not(rs.Eof Or rs.Bof) Then 
		arrManager = rs.GetRows()
	End If 
	rsClose(rs)

	Dim OP_IDX, OP_CURR_YMD, OP_CART_CODE, OP_ITEM_CODE, OP_ITEM_NAME, OP_BREAKFAST, OP_LUNCH, OP_DINNER 
	Dim OP_STAFF1, OP_COMPLAINT1, OP_COMPLAINT_ETC1
	Dim OP_STAFF2, OP_COMPLAINT2, OP_COMPLAINT_ETC2 
	Dim OP_STAFF3, OP_COMPLAINT3, OP_COMPLAINT_ETC3 
	Dim OP_REG_DATE_IDX, OP_REG_DATE, OP_CONFIRM_FLAG


	sql = "EXEC USP_CART_RECORD_DETAIL_VIEW_JOIN  '{P1}','{P2}';"
	sql = Replace(sql , "{P1}" , sCurrDate , vbTextCompare)
	sql = Replace(sql , "{P2}" , sCartCode , vbTextCompare)
	Set rs = getQueryDB(sql)
	If Not(rs.Eof Or rs.Bof) Then 
		arrData = rs.GetRows()
	End If 
	rsClose(rs)

	sql = "SELECT * FROM ( "
	sql = sql & "SELECT MEAL_TYPE , REG_DATE "
	sql = sql & "FROM TBL_CART_ETC WHERE CURR_YMD = '{P1}' AND CART_CODE = '{P2}') AS TB "
	sql = sql & "PIVOT (MAX(REG_DATE) FOR MEAL_TYPE IN ([1],[2],[3])) AS PVT;"
	sql = Replace(sql , "{P1}" , sCurrDate , vbTextCompare)
	sql = Replace(sql , "{P2}" , sCartCode , vbTextCompare)
	
	Set rs = getQueryDB(sql)
	If Not(rs.Eof Or rs.Bof) Then 
		arrRegDate(0) = nullFillWith(rs(0) , "" , False)
		arrRegDate(1) = nullFillWith(rs(1) , "" , False)
		arrRegDate(2) = nullFillWith(rs(2) , "" , False)
	Else 
		arrRegDate(0) = ""
		arrRegDate(1) = ""
		arrRegDate(2) = ""
	End If 
	rsClose(rs)
	%>

	<style type="text/css">
		.bx-calendar {
			position: absolute;
			right: 10px;
			pointer-events: none;
		}
	</style>
  </head>
  <body>
    <div id="root">
      <div id="app">
        <header id="header">
			<!-- #include virtual = "/common/include/layout_left.asp"-->
        </header>
        <div id="wrap">
          <div class="scroll-area">
            <main id="main">
              <div class="main-container">
                <!-- #include virtual = "/common/include/layout_top.asp"-->
                <div class="breadcrumb">
                  <a href=""><i class="bx bx-home-alt-2"></i></a>
                  <i class="bx bx-chevron-right"></i>
                  <a href="">점검일지 등록</a>
                  <i class="bx bx-chevron-right"></i>
                  <a class="current" href="">배선카트 등록/수정</a>
                </div>
                <div class="main-content">
                  <section class="section">
                    <div class="section-container">
                      <div class="section-title">
                        <h2>배선카트 등록/수정</h2>
                      </div>
                      <div class="section-content">
                        <hr class="divider" />
                        <form class="form" name="dataForm" id="dataForm">
						<input type="hidden" name="command" id="command" value="Meal_Cart_Data" title="명령어" class="LEB">
						<input type="hidden" name="idx" id="idx" value="<%=sIdx%>" title="페이지번호"  class="LEB">
						<input type="hidden" name="page" id="page" value="<%=sPage%>" title="페이지번호" class="LEB">
						<input type="hidden" name="confirm" id="confirm" title="승인여부" class="LEB">
						  <fieldset class="input-group flex-box">
                            <div class="input-field item1-6">
                              <label class="label">날짜</label>
							<div class="date-picker">
							  <div class="input date">
								<input type="text" name="currDate" id="currDate" placeholder="YYYY-MM-DD" readonly value="<%=convertDate(sCurrDate,"-")%>"/>
								<i class="bx bx-calendar"></i>
							  </div>
							</div>
							</div>
                          </fieldset>
						  <hr class="divider" />
                          <fieldset class="input-group flex-box">
                            <div class="input-field item1-5">
                              <label class="label">카트번호</label>
                            <div class="select">
                              <select name="cartCode" id="cartCode">
                                <option value="">카트번호(전체)</option>
								<%								
								IF isArray(arrCart) Then
									For i = 0 To UBound(arrCart, 2)
										CP_IDX			= arrCart(0 , i)
										CP_CART_CODE	= arrCart(1 , i)
										CP_CART_NAME	= arrCart(2 , i)
								%><option value="<%=CP_CART_CODE%>" <%=selectValue(CP_CART_CODE,sCartCode,"select")%>><%=CP_CART_NAME%>&nbsp;(<%=CP_CART_CODE%>)</option><%	
									Next 
								End If 
								%>
                              </select>
                            </div>
                            </div>
                          </fieldset>
                          <hr class="divider" />						  
						  <fieldset class="input-group flex-box">
							<div class="input-field">
							  <label class="label">구분 리스트</label>
							  <div class="table">
								  <table>
									<thead>
									  <tr>
										<th class="fit-cell">구분</th>
										<th class="fit-cell">조식</th>
										<th class="fit-cell">중식</th>
										<th class="fit-cell">석식</th>
									  </tr>
									</thead>
									<tbody>
									<%
									IF isArray(arrData) Then
										For i = 0 To UBound(arrData, 2)
											OP_IDX				= arrData(0,i)
											OP_CURR_YMD			= arrData(1,i)
											OP_CART_CODE		= arrData(2,i)
											OP_ITEM_CODE		= arrData(3,i)
											OP_ITEM_NAME		= arrData(4,i)
											OP_BREAKFAST		= nullFillWith(arrData(5,i), "", False)
											OP_LUNCH			= nullFillWith(arrData(6,i), "", False)
											OP_DINNER			= nullFillWith(arrData(7,i), "", False)
											OP_COMPLAINT1		= nullFillWith(arrData(8,i), "", False)
											OP_COMPLAINT_ETC1	= nullFillWith(arrData(9,i), "", False)
											OP_STAFF1			= nullFillWith(arrData(10,i), "", False)
											OP_COMPLAINT2		= nullFillWith(arrData(11,i), "", False)
											OP_COMPLAINT_ETC2	= nullFillWith(arrData(12,i), "", False)
											OP_STAFF2			= nullFillWith(arrData(13,i), "", False)
											OP_COMPLAINT3		= nullFillWith(arrData(14,i), "", False)
											OP_COMPLAINT_ETC3	= nullFillWith(arrData(15,i), "", False)
											OP_STAFF3			= nullFillWith(arrData(16,i), "", False)
											OP_CONFIRM_FLAG		= arrData(17,i)
											OP_REG_DATE_IDX		= arrData(18,i)
											OP_REG_DATE			= arrData(19,i)
											If i = 0 Then 
												arrComplaint(0)	= OP_COMPLAINT1 : arrComplaint(1)	= OP_COMPLAINT2 : arrComplaint(2)	= OP_COMPLAINT3
												arrEtc(0) = OP_COMPLAINT_ETC1 : arrEtc(1) = OP_COMPLAINT_ETC2 : arrEtc(2) = OP_COMPLAINT_ETC3
												arrStaff(0) = OP_STAFF1 : arrStaff(1) = OP_STAFF2 : arrStaff(2) = OP_STAFF3
											End If 
									%>
									  <tr>										
										<td class="fit-cell"><%=OP_ITEM_NAME%>
										<input type="hidden" name="itemIdx" id="itemIdx" value="<%=OP_IDX%>" title="인덱스" style="background-color:lime;">
										<input type="hidden" name="itemCode" id="itemCode" value="<%=OP_ITEM_CODE%>" title="아이템코드" style="background-color:lime;">
										</td>
										<td class="fit-cell">
											<div class="select">
												<select id="breakfast" name="breakfast">
													<option value="">선택</option>
													<%=getCode("select","CHECK_FLAG",OP_BREAKFAST)%>
												</select>
											</div>
										</td>
										<td class="fit-cell">
											<div class="select">
											<div class="select">
												<select id="lunch" name="lunch">
													<option value="">선택</option>
													<%=getCode("select","CHECK_FLAG",OP_LUNCH)%>
												</select>
											</div>
											</div>
										</td>
										<td class="fit-cell">
											<div class="select">
											<div class="select">
												<select id="dinner" name="dinner">
													<option value="">선택</option>
													<%=getCode("select","CHECK_FLAG",OP_DINNER)%>
												</select>
											</div>
											</div>
										</td>
									  </tr>
									<%
										Next 
									Else 
										rw("<tr><td colspan=""4"">등록된 데이타가 없습니다</td></tr>")
									End If 
									%>
									</tbody>
								  </table>
							  </div>
							</div>
                          </fieldset>							
                          <hr class="divider" />
						  <fieldset class="input-group flex-box">
                            <div class="input-field">
                              <label class="label">지적사항</label>
								<div class="table">
									<table>
										<thead>
											<tr>
												<th class="fit-cell">지적사항</th>
												<th class="fit-cell">조식</th>
												<th class="fit-cell">중식</th>
												<th class="fit-cell">석식</th>
											</tr>
										</thead>
									<tbody>
										<tr>
											<td class="fit-cell">	
											지적사항
											</td>
											<% For j = 0 To 2 %>
											<td class="fit-cell">
												<div class="select">
												  <select name="complaint" id="complaint">
													<%												
													IF isArray(arrComment) Then
														For i = 0 To UBound(arrComment, 2)
															VP_IDX			= arrComment(0 , i)
															VP_CODE_NM	= arrComment(1 , i)
															VP_CART_VAL	= arrComment(2 , i)
													%><option value="<%=VP_CODE_NM%>" <%=selectValue(VP_CODE_NM,arrComplaint(j),"select")%>><%=VP_CODE_NM%></option><%	
														Next 
													End If 
													%>
												  </select>
												</div>											
											</td>
											<% Next %>
										</tr>
										<tr>
											<td class="fit-cell">기타</td>
											<% For j = 0 To 2 %>
											<td class="fit-cell">
											<div class="input"><input id="etc" name="etc" type="text" placeholder="기타" value="<%=arrEtc(j)%>" /></div>
											</td>
											<% Next %>
										</tr>
										<tr>
											<td class="fit-cell">점검자</td>
											<% For j = 0 To 2 %>
											<td class="fit-cell">                             
												<div class="select">
												  <select name="managerId" id="managerId">
												  <option value="">선택</option>
													<%												
													IF isArray(arrManager) Then
														For i = 0 To UBound(arrManager, 2)
													%><option value="<%=arrManager(3 , i)%>" <%=selectValue(arrManager(3 , i),arrStaff(j),"select")%>><%=arrManager(4 , i)%>&nbsp;(<%=arrManager(3 , i)%>)</option><%	
														Next 
													End If 
													%>
												  </select>
												</div>
											</td>
											<% Next %>							  
										</tr>
										<tr style="display:none;">
											<td class="fit-cell">등록일</td>
											<% For j = 0 To 2 %>
											<td class="fit-cell">
												<%=arrRegDate(j)%>
											</td>
											<% Next %>
										</tr>
									</tbody>
									</table>
								</div>

                            </div>
                          </fieldset>
							
                          <!--<hr class="divider" />-->
						  <%
							Dim AP_CHA_MANAGER , AP_CHA_STAFF , AP_HD_MANAGER , AP_HD_STAFF
							sql = "EXEC ADM_APPROVE_DETAIL_VIEW '{P1}','{P2}','','';"
							sql = Replace(sql , "{P1}" , sCurrDate , vbTextCompare)
							sql = Replace(sql , "{P2}" , "CART"    , vbTextCompare)

							Set rs = getQueryDB(sql)
							If Not(rs.Eof Or rs.Bof) Then 
								AP_CHA_MANAGER	 = nullFillWith(rs("CHA_MANAGER"),"",False)
								AP_CHA_STAFF	 = nullFillWith(rs("CHA_STAFF"),"",False)
								AP_HD_MANAGER	 = nullFillWith(rs("HD_MANAGER"),"",False)
								AP_HD_STAFF		 = nullFillWith(rs("HD_STAFF"),"",False)
							Else 
								AP_CHA_MANAGER	= ""
								AP_CHA_STAFF	= ""
								AP_HD_MANAGER	= ""
								AP_HD_STAFF		= ""
							End If 
							rsClose(rs)
						  %>
						  <fieldset class="input-group flex-box" style="margin-top:40px;display:none;">
                            <div class="input-field">
                              <label class="label">결재</label>
								<div class="table">
									<table>
										<thead>
											<tr>
												<th class="fit-cell" colspan="2">분당차병원</th>
												<th class="fit-cell" colspan="2" style="border-left:1px solid #cdd0d4;">현대그린푸드</th>
											</tr>
											<tr>
											<td class="fit-cell">팀장</td>
											<td class="fit-cell" style="border-left:1px solid #cdd0d4;">팀원</td>
											<td class="fit-cell" style="border-left:1px solid #cdd0d4;">팀장</td>
											<td class="fit-cell" style="border-left:1px solid #cdd0d4;">팀원</td>
											</tr>
										</thead>
									<tbody>
										<tr>
											<td class="fit-cell">
												<%
												If AP_CHA_MANAGER = "" And vAdminCompany = "CHA" And vAdminPosition = "MANAGER" Then 
												%>
													<button class="button orange" onClick="javascript:btnApprove('4');void(0);">승인하기</button>
												<% 
												End If 
												If AP_CHA_MANAGER <> "" Then 
													rw("<button class=""button dark"" >승인완료 : "&AP_CHA_MANAGER&"</button>")
												End If 
												%>
											</td>
											<td class="fit-cell" style="border-left:1px solid #cdd0d4;">
												<%
												If AP_CHA_STAFF = "" And vAdminCompany = "CHA" And vAdminPosition = "STAFF"Then 
													rw("<button class=""button orange"" onClick=""javascript:btnApprove('3');void(0);"">승인하기</button>")
												End If 

												If AP_CHA_STAFF <> "" Then 
													rw("<button class=""button dark"" >승인완료 : "&AP_CHA_STAFF&"</button>")
												End If 
												%>
											</td>
											<td class="fit-cell" style="border-left:1px solid #cdd0d4;">
												<%
												If AP_HD_MANAGER = "" And vAdminCompany = "HD" And vAdminPosition = "MANAGER" Then 
													rw("<button class=""button sub"" onClick=""javascript:btnApprove('2');void(0);"">승인하기</button>")
												End If 

												If AP_HD_MANAGER <> "" Then 
													rw("<button class=""button dark"" >승인완료 : "&AP_HD_MANAGER&"</button>")
												End If 
												%>
											</td>
											<td class="fit-cell" style="border-left:1px solid #cdd0d4;">
												<%
												If AP_HD_STAFF = "" And vAdminCompany = "HD" And vAdminPosition = "STAFF" Then 
													rw("<button class=""button sub"" onClick=""javascript:btnApprove('1');void(0);"">승인하기</button>")
												End If 

												If AP_HD_STAFF <> "" Then 
													rw("<button class=""button dark"" >승인완료 : "&AP_HD_STAFF&"</button>")
												End If 
												%>
											</td>
										</tr>
									</tbody>
									</table>
								</div>

                            </div>
                          </fieldset>
										
                          <hr class="divider" />							 
                          <div class="form-control">
                            <div class="right-box">
                              <a href="meal_cart.asp" class="button outline def">돌아가기</a>
                              <button type="button" class="button" id="btnSubmit" style="display:block;">등록하기</button>		
							  <% If sIdx <> "0" Then %>
                              <button type="button" class="button orange" id="btnRemove" style="display:block;">삭제하기</button>
							  <% End If %>
							  <%If OP_CONFIRM_FLAG = "2" Then %>
							  <!-- <button class="button red" onClick="javascript:chgConfirm('1');void(0);">승인하기</button> -->
							  <% Else %>
							  <!-- <button class="button dark" onClick="javascript:chgConfirm('2');void(0);">승인취소</button> -->
							  <% End If %>
                            </div>
                          </div>
                        </form>
                      </div>
                    </div>
                  </section>
                </div>
              </div>
            </main>
            <!-- #include virtual = "/common/include/layout_footer.asp"-->
          </div>
        </div>
      </div>
    </div>
  </body>
</html>
<script type="text/javascript">
<!--
$(document).ready(function(){

	const elem = document.getElementById('currDate');
	const rangepicker = new DateRangePicker(elem, {});

	elem.addEventListener('changeDate', function(e) {
	  //console.log(e.detail.date); // new Date() 찍어봤을때랑 똑같은 형태로 나옴
		var selDate = e.detail.date;
		var year = selDate.getFullYear();
		var month = selDate.getMonth() + 1;
		var day = selDate.getDate();
		//console.log(year + "-" + month + "-" + day);
		$("#cartCode").change();
	});

	$("#btnSubmit").click(function(){
		var frm = document.dataForm;
		if(!setInputAlert(frm.currDate, "날짜를")) return false;
		if(!setSelectAlert(frm.cartCode, "카트번호를")) return false;
	
		frm.page.value = "1";
		frm.action = "./meal_cart_proc.asp";			
		frm.target = "procFrame";
		frm.method = "POST";
		frm.submit();	
	});

	$("#cartCode").change(function(){
		var frm = document.dataForm;

		if(!setInputAlert(frm.currDate, "날짜를")) return false;
		if(!setSelectAlert(frm.cartCode, "카트번호를")) return false;
		frm.idx.value = "0";
		frm.page.value = "1";
		frm.action = "./meal_cart_register.asp";
		frm.target = "_self";
		frm.method = "POST";
		frm.submit();	
	});

	// 삭제하기
	$("#btnRemove").click(function(){
		Swal.fire({
		  title: "삭제하시겠습니까?",
		  text: "삭제된 데이타는 복구가 불가능합니다!",
		  icon: "question",
		  showCancelButton: true,
		  confirmButtonColor: "#3085d6",
		  cancelButtonColor: "#d33",
		  confirmButtonText: "삭제",
  		  cancelButtonText: "취소"
		}).then((result) => {
		  if (result.isConfirmed) {
			var frm = document.dataForm;
			if(!setSelectAlert(frm.cartCode, "카트번호를")) return false;
			frm.command.value = "Meal_Cart_Remove";
			frm.action = "./meal_cart_proc.asp";			
			frm.target = "procFrame";
			frm.method = "POST";
			frm.submit();	
		  }
		});	
	});
});		
// End Ready

function btnApprove(v) {
	var frm = document.dataForm;
	frm.confirm.value = v;
	frm.command.value = "Meal_Cart_Approve";
	frm.action = "./meal_cart_proc.asp";			
	frm.target = "procFrame";
	frm.method = "POST";
	frm.submit();	
}


//출력하기
function reportPrint() {
	 openPopup('meal_cart_print.asp','MEAL_CART',800,700,'scrollbars=yes,width=800,height=700');
}

function chgConfirm(v) {
	if ($("#idx").val() == "0") {
		sAlert("error","등록 후 승인처리 가능합니다.")
		return false;
	}

	var frm = document.dataForm;
	frm.confirm.value = v;
	frm.command.value = "Meal_Cart_Confirm";
	frm.action = "./meal_cart_proc.asp";			
	frm.target = "procFrame";
	frm.method = "POST";
	frm.submit();	
}

//-->
</script>
<iframe  name="procFrame" id="procFrame" style="width:900px; height:500px;display:none;" class="LEDB"></iframe>