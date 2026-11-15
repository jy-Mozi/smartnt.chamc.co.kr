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
<!-- include virtual = "/common/include/requestValue.asp"-->
<!doctype html>
<html lang="ko">
  <head>
    <!-- #include virtual = "/common/include/layout_header.asp"-->
	<%

	Dim sAreaCode , sFmDate , sToDate, page, selHoType
	sAreaCode	= nullFillWith(Request("areaCode"),"",False)
	sFmDate		= nullFillWith(request("fmDate"),Date(),False)
	sToDate		= nullFillWith(request("toDate"),Date(),False)
	page		= nullFillWith(Request("page"),"1",false)
	selHoType	= nullFillWith(Request("selHoType"),"",False)
	Dim listSize , pageSize , totCnt , totPage
	Dim vTemp , vNumber 
	listSize	= 25
	pageSize    = 10
	Session("Param") = "&fmDate=" & sFmDate & "&toDate=" & sToDate& "&areaCode=" & sAreaCode
	Session("rtnParam") = "?page=" & page & Session("Param")
	Dim sql , rs , i , j , arrData , arrArea

	Dim AP_IDX, AP_AREA_CODE, AP_AREA_NAME, AP_ITEM_LIST , AP_ETC
	sql = "EXEC ADM_MACHINE_CELSIUS_LIST '{P1}','{P2}','{P3}';"
	sql = Replace(sql , "{P1}", selHoType , vbTextCompare)
	sql = Replace(sql , "{P2}", "" , vbTextCompare)
	sql = Replace(sql , "{P3}", "" , vbTextCompare)

	Set rs = getQueryDB(sql)
	If Not(rs.Eof Or rs.Bof) Then 
		arrArea	= rs.GetRows()
	End If 
	rsClose(rs)

	Dim OP_CURR_YMD , OP_AREA_CODE , OP_AREA_NAME , arrMeal(2) , mealType
	arrMeal(0)	= "조식"
	arrMeal(1)	= "중식"
	arrMeal(2)	= "석식"

	sql = "EXEC ADM_MACHINE_CELSIUS_RECORD_LIST_PAGEING_V1 '{P1}','{P2}','{P3}','{P4}','{P5}','{P6}','{P7}';"
	sql = Replace(sql , "{P1}", page		, vbTextCompare)
	sql = Replace(sql , "{P2}", listSize	, vbTextCompare)
	sql = Replace(sql , "{P3}", "CD"		, vbTextCompare)
	sql = Replace(sql , "{P4}", sAreaCode	, vbTextCompare)
	sql = Replace(sql , "{P5}", sFmDate , vbTextCompare)
	sql = Replace(sql , "{P6}", sFmDate , vbTextCompare)
	sql = Replace(sql , "{P7}", selHoType , vbTextCompare)
	'rw(sql)
	Set rs = getQueryDB(sql)
	If Not(rs.Eof Or rs.Bof) Then 
		arrData = rs.GetRows()
		totCnt		= arrData(0,0)				'총 계시물
		totPage		= Ceil(totCnt , listSize)	'총 페이지
	Else 
		totCnt		= 0
		totPage		= 1
	End If 
	rsClose(rs)
	%>
	<style type="text/css">
		.table-search {
		 & .form {
			.datepicker-controls {
				.button {
					width: auto;
					padding-inline: 0px;
					}		
				}
			}
		}

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
                  <a class="current" href="">온/습도 점검일지</a>
                </div>
                <div class="main-content">
                  <section class="section">
                    <div class="section-container">
                      <div class="section-title">
                        <h2>온/습도 점검일지 서비스 일지</h2>
                      </div>
                      <div class="section-content">
                        <div class="table-search">
                          <form class="form" name="dataForm" id="dataForm">
							<input type="hidden" name="command" id="command" value="Cart_Celsius" title="명령어:" class="LEB">
							<input type="hidden" name="idx" id="idx" value="" title="인덱스" class="LEB">
							<input type="hidden" name="page" id="page" value="<%=page%>" title="페이지번호" class="LEB">
							<input type="hidden" name="currDate" id="currDate" title="선택 날짜" class="LEB">
							<input type="hidden" name="selAreaCode" id="selAreaCode" title="선택 구역번호" class="LEB">
                            <div class="select">
							<select id="selHoType" name="selHoType">
								<option value="">전체</option>
								<%=getCodeDB("select" , "P11" , selHoType)%>			
							</select>
                            </div>
                            <div class="select">
                              <select name="areaCode" id="areaCode">
                                <option value="">구역번호(전체)</option>
								<%
								IF isArray(arrArea) Then
									For i = 0 To UBound(arrArea, 2)
										AP_IDX			= arrArea(0 , i)
										AP_AREA_CODE	= arrArea(1 , i)
										AP_AREA_NAME	= arrArea(2 , i)
										AP_ITEM_LIST	= arrArea(3 , i)
										AP_ETC			= arrArea(4 , i)
								%><option value="<%=AP_AREA_CODE%>" <%=selectValue(AP_AREA_CODE,sAreaCode,"select")%>><%=AP_AREA_NAME%>&nbsp;(<%=AP_AREA_CODE%>)</option><%	
									Next 
								End If 
								%>
                              </select>
                            </div>
							<div class="date-picker">
							  <div class="input date">
								<input type="text" name="fmDate" id="fmDate" value="<%=sFmDate%>" placeholder="YYYY-MM-DD" readonly />
								<i class="bx bx-calendar"></i>
							  </div>
							</div>
							<div class="date-picker" style="display:none;">
							  <div class="input date">
								<input type="text" name="toDate" id="toDate" value="<%=sToDate%>" placeholder="YYYY-MM-DD" readonly />
								<i class="bx bx-calendar"></i>
							  </div>
							</div>
                            <button type="button" class="button" id="btnSearch">검색</button>
                            <button type="button" class="button outline def">엑셀 다운로드</button>
							<button class="button red" onClick="javascript:reportPrint();void(0);">승인•출력하기</button>
                          </form>
                        </div>
                        <div class="table">
                          <table>
                            <thead>
                              <tr>
                                <th class="fit-cell">번호</th>
								<th class="fit-cell">날짜</th>
                                <th class="fit-cell">구역번호</th>
                                <th class="fit-cell">구역명</th>
								<th class="fit-cell">구분</th>
                                <th class="fit-cell">점검자</th>
								<th class="fit-cell">점검일자</th>
                              </tr>
                            </thead>
                            <tbody>
							<%
							Dim MP_STAFF , MP_REG_DATE , MP_STAFF_NM
							IF isArray(arrData) Then
							vTemp = 1
							vNumber = totCnt - listSize * (page - 1)
								For i = 0 To UBound(arrData, 2)

									OP_CURR_YMD		= arrData(2,i)
									OP_AREA_CODE	= arrData(3,i)
									OP_AREA_NAME	= arrData(4,i)
									For j = 0 To 2
										mealType = j + 1
										sql = "EXEC USP_MACHINE_CELSIUS_ETC_LIST '{P1}','{P2}','{P3}';"
										sql = Replace(sql , "{P1}" , OP_CURR_YMD , vbTextCompare)
										sql = Replace(sql , "{P2}" , OP_AREA_CODE , vbTextCompare)
										sql = Replace(sql , "{P3}" , mealType , vbTextCompare)
										'rw(sql & "<br/>")
										Set rs = getQueryDB(sql)
										If Not(rs.Eof Or rs.Bof) Then 
											MP_STAFF		= rs("STAFF_NAME")
											'MP_REG_DATE		= rs("REG_DATE")
											'MP_STAFF_NM		= rs("ADMIN_NM")
										Else 								
											MP_STAFF		= ""
											MP_REG_DATE		= ""
											MP_STAFF_NM	= ""
										End If 

							%>
                              <tr onClick="btnDetail('<%=OP_CURR_YMD%>','<%=OP_AREA_CODE%>');" style="cursor:pointer;">
							  <% If j = 0 Then %>
                                <td class="fit-cell" rowspan="3"><%=vNumber%></td>
								<td class="fit-cell" rowspan="3"><%=convertDate(OP_CURR_YMD,"-")%></td>
                                <td class="fit-cell" rowspan="3"><%=OP_AREA_CODE%></td>
                                <td class="fit-cell" rowspan="3"><%=OP_AREA_NAME%></td>
   						      <% End If %>
								<td class="fit-cell"><%=arrMeal(j)%></td>
								<td class="fit-cell"><%=MP_STAFF%></td>
								<td class="fit-cell"><%=MP_REG_DATE%></td>
                              </tr>
							<%
									Next 
									vNumber = vNumber - 1
									vTemp = vTemp + 1
								Next 
							Else 
								rw("<tr><td colspan=""8"">등록된 데이타가 없습니다</td></tr>")
							End If 
							%>                              
                            </tbody>
                          </table>
                        </div>
                        <div class="table-control">
                          <div class="left-box">
							<%=getListPaging(page , Session("Param") , totCnt, totPage)%>
                          </div>
                          <div class="right-box">
                            <button type="button" class="button" id="btnSubmit">등록하기</button>
                          </div>
                        </div>
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
		$("#btnSearch").click(function(){
			var frm = document.dataForm;
			
			frm.action = "./cart_temperature_.asp";
			frm.target = "_self";
			frm.method = "POST";
			frm.submit();				
		});

		$("#btnSubmit").click(function(){
			var frm = document.dataForm;
			frm.page.value = "1";
			frm.idx.value = "0";
			frm.action = "./cart_temperature_register.asp";			
			frm.target = "_self";
			frm.method = "POST";
			frm.submit();	
		});

		$("#selHoType , #areaCode").change(function(){
			$("#btnSearch").trigger("click");
		});
	});		
	// End Ready

	function btnDetail(d,i) {
		var frm = document.dataForm;
		$("#currDate").val(d);
		$("#selAreaCode").val(i);
		frm.idx.value = "1";
		frm.page.value = "1";
		frm.action = "./cart_temperature_register.asp";			
		frm.target = "_self";
		frm.method = "POST";
		frm.submit();	
	}

	//출력하기
	function reportPrint() {
		 openPopup('cart_temperature_print.asp?selDate='+ $("#fmDate").val() +'&selHoType=' + $("#selHoType").val(),'CART_TEMPERATURE',1600,700,'scrollbars=yes,width=1600,height=700');
	}


//-->
</script>
<iframe  name="procFrame" id="procFrame" style="width:900px; height:500px;display:none;background-color:lime;"></iframe>