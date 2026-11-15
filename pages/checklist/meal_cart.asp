<%@CodePage=65001%>
<% 
'=============================================
' 작 성 자 : 노 광 옹
' 작 성 일 : 2024-10-17
' 파일내용 : 점검일지 등록>배선카트 리스트
'=============================================
Option Explicit 
%>
<!-- #include virtual = "/common/library/common.asp"-->
<!doctype html>
<html lang="ko">
  <head>
    <!-- #include virtual = "/common/include/layout_header.asp"-->
	<%
	Dim sFmDate , sToDate , sCart, page, selHoType
	sFmDate		= nullFillWith(request("fmDate"),Date(),False)
	sToDate		= nullFillWith(request("toDate"),Date(),False)
	sCart		= nullFillWith(request("cart"),"",False)
	page		= nullFillWith(Request("page"),"1",false)
	selHoType	= nullFillWith(Request("selHoType"),"",False)

	Dim listSize , pageSize , totCnt , totPage
	Dim vTemp , vNumber 
	listSize	= 25
	pageSize    = 10
	Session("Param") = "&fmDate=" & sFmDate & "&toDate=" & sToDate& "&cart=" & sCart


	Dim sql , rs , arrCart, arrData , i , j
	sql = "EXEC ADM_CART_LIST '{P1}','{P2}';"
	sql = Replace(sql , "{P1}" , "" , vbTextCompare)
	sql = Replace(sql , "{P2}" , selHoType , vbTextCompare)

	Set rs = getQueryDB(sql)
	If Not(rs.Eof Or rs.Bof) Then 
		arrCart = rs.GetRows()
	End If 
	rsClose(rs)

	sql = "EXEC ADM_CART_RECORD_LIST_PAGEING_V1 '{P1}','{P2}','{P3}','{P4}','{P5}','{P6}','{P7}';"
	sql = Replace(sql , "{P1}", page , vbTextCompare)
	sql = Replace(sql , "{P2}", listSize , vbTextCompare)
	sql = Replace(sql , "{P3}", "CD" , vbTextCompare)
	sql = Replace(sql , "{P4}", sCart , vbTextCompare)
	sql = Replace(sql , "{P5}", sFmDate , vbTextCompare)
	sql = Replace(sql , "{P6}", sFmDate , vbTextCompare)
	sql = Replace(sql , "{P7}" , selHoType , vbTextCompare)

	Set rs = getQueryDB(sql)
	If Not(rs.Eof Or rs.Bof) Then 
		arrData	= rs.GetRows()
		totCnt		= arrData(0,0)			'총 계시물
		totPage		= Ceil(totCnt , listSize)	'총 페이지
	Else 
		totCnt		= 0
		totPage		= 1
	End If 

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
                  <a class="current" href="">배선카트</a>
                </div>
                <div class="main-content">
                  <section class="section">
                    <div class="section-container">
                      <div class="section-title">
                        <h2>배선카트</h2>
                      </div>
                      <div class="section-content">
                        <div class="table-search">
                          <form class="form" name="dataForm" id="dataForm">
							<input type="hidden" name="command" id="command" value="UserAdd" title="명령어:기본 부모정보코드등록" style="background-color:lime;">
							<input type="hidden" name="idx" id="idx" value="" title="인덱스" style="background-color:lime;">
							<input type="hidden" name="page" id="page" value="<%=page%>" title="페이지번호" style="background-color:lime;">
							<input type="hidden" name="currDate" id="currDate" title="선택 날짜">
							<input type="hidden" name="cartCode" id="cartCode" title="카트 코드">
                            <div class="select">
							<select id="selHoType" name="selHoType">
								<option value="">전체</option>
								<%=getCodeDB("select" , "P11" , selHoType)%>			
							</select>
                            </div>
                            <div class="select">
                              <select name="cart" id="cart">
                                <option value="">카트번호</option>
								<%
								Dim CP_IDX, CP_CART_CODE, CP_CART_NAME
								IF isArray(arrCart) Then
									For i = 0 To UBound(arrCart, 2)
										CP_IDX			= arrCart(0 , i)
										CP_CART_CODE	= arrCart(1 , i)
										CP_CART_NAME	= arrCart(2 , i)
								%><option value="<%=CP_CART_CODE%>" <%=selectValue(CP_CART_CODE,sCart,"select")%>><%=CP_CART_NAME%>&nbsp;(<%=CP_CART_CODE%>)</option><%	
									Next 
								End If 
								%>
                              </select>
                            </div>
							<div class="date-picker">
							  <div class="input date">
								<input type="text" name="fmDate" id="fmDate" placeholder="YYYY-MM-DD" readonly value="<%=sFmDate%>"/>
								<i class="bx bx-calendar"></i>
							  </div>
							</div>
							<div class="date-picker" style="display:none;">
							  <div class="input date">
								<input type="text" name="toDate" id="toDate"placeholder="YYYY-MM-DD" readonly value="<%=sToDate%>"/>
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
                                <th class="fit-cell">배선카트 명</th>
                                <th class="fit-cell">구분</th>
                                <th class="fit-cell">점검자</th>
                              </tr>
                            </thead>
                            <tbody>
							<%
							Dim OP_CURR_YMD , OP_CART_CODE, OP_CART_NAME , arrMeal(2) , mealType
							Dim MP_STAFF , MP_REG_DATE , MP_SATAFF_NM
							arrMeal(0)	= "조식"
							arrMeal(1)	= "중식"
							arrMeal(2)	= "석식"
							IF isArray(arrData) Then
							vTemp = 1
							vNumber = totCnt - listSize * (page - 1)
								For i = 0 To UBound(arrData, 2)
									OP_CURR_YMD		= arrData(2 , i)
									OP_CART_CODE	= arrData(3 , i)
									OP_CART_NAME	= arrData(4 , i)

									For j = 0 To 2

									mealType = j + 1
									sql = "SELECT TA01.* , TA02.ADMIN_NM FROM TBL_CART_ETC AS TA01 LEFT OUTER JOIN TBL_ADMIN AS TA02 ON (TA01.STAFF = TA02.ADMIN_ID) "
									sql = sql & " WHERE TA01.CURR_YMD = '{P1}' AND TA01.CART_CODE = '{P2}' AND TA01.MEAL_TYPE = '{P3}';"
									sql = Replace(sql , "{P1}" , OP_CURR_YMD , vbTextCompare)
									sql = Replace(sql , "{P2}" , OP_CART_CODE , vbTextCompare)
									sql = Replace(sql , "{P3}" , mealType , vbTextCompare)
									Set rs = getQueryDB(sql)
									If Not(rs.Eof Or rs.Bof) Then 
										MP_STAFF		= rs("STAFF")
										MP_REG_DATE		= rs("REG_DATE")
										MP_SATAFF_NM	= rs("ADMIN_NM")
									Else 								
										MP_STAFF		= ""
										MP_REG_DATE		= ""
										MP_SATAFF_NM	= ""
									End If 
							%>
                              <tr onClick="btnDetail('<%=OP_CURR_YMD%>','<%=OP_CART_CODE%>');" style="cursor:pointer;">
								<% If j = 0 Then %>
                                <td class="fit-cell" rowspan="3"><%=vNumber%></td>
								<td class="fit-cell" rowspan="3"><%=convertDate(OP_CURR_YMD,"-")%></td>
                                <td class="fit-cell" rowspan="3"><%=OP_CART_NAME%>&nbsp;(<%=OP_CART_CODE%>)</td>
								<% End If %>
                                <td class="fit-cell"><%=arrMeal(j)%></td>
                                <td class="fit-cell"><%=MP_SATAFF_NM%></td>
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
                            <a href="#" class="button" id="btnSubmit">등록하기</a>
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
			
			frm.action = "./meal_cart.asp";
			frm.target = "_self";
			frm.method = "POST";
			frm.submit();				
		});

		$("#btnSubmit").click(function(){
			var frm = document.dataForm;
			frm.idx.value = "0";
			frm.page.value = "1";
			frm.action = "./meal_cart_register.asp";			
			frm.target = "_self";
			frm.method = "POST";
			frm.submit();	
		});

		$("#selHoType, #cart").change(function(){
			$("#btnSearch").trigger("click");
		});
	});		
	// End Ready

	//출력하기
	function reportPrint() {
		 openPopup('meal_cart_print.asp?selDate=' + $("#fmDate").val(),'MEAL_CART',800,700,'scrollbars=yes,width=800,height=700');
	}

	function btnDetail(d,c) {
		$("#currDate").val(d);
		$("#cartCode").val(c);
		
		var frm = document.dataForm;
		frm.idx.value = "1";
		frm.page.value = "1";
		frm.action = "./meal_cart_register.asp";			
		frm.target = "_self";
		frm.method = "POST";
		frm.submit();
	}
//-->
</script>
<iframe  name="procFrame" id="procFrame" style="width:900px;height:500px;display:none;" Class="LEB"></iframe>