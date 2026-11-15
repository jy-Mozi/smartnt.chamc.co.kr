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
<%
	Dim vIdx
		vIdx	= nullFillWith(Request("idx"), "" , False)
	
	Dim vSql, vRs, vI, vTxt, sql , rs
	' 구분항목 리스트
	Dim vItemData, vIn(0)
	vIn(0) = "P04"
	Set vRs = getSP("USP_CODE_COMMON_LIST",vIn)
	If Not(vRs.Eof Or vRs.Bof) Then
		vItemData = vRs.GetRows()
	End If 
	rsClose(vRs)

	
	' 수정일 경우 불러올 데이터
	Dim vAreaCode, vAreaName, vItemList, vEtc, arr_itemlist, arr_itemlistCode, vItemListNm, vHospitalType, vEtc2
	If vIdx <> "" Then 
		vTxt = "수정"

		vSql = "SELECT AREA_CODE, AREA_NAME, ITEM_LIST, HOSPITAL_TYPE , ETC, ETC2 , dbo.FUNC_STRING_COMMON_CODE(ITEM_LIST,',') AS ITEM_LIST_NM FROM TBL_MACHINE_CELSIUS WHERE IDX = " & vIdx
		Set vRs = getQueryDB(vSql)
		If Not(vRs.Eof Or vRs.Bof) Then
			vAreaCode		= vRs("AREA_CODE")
			vAreaName		= vRs("AREA_NAME")
			vItemList		= vRs("ITEM_LIST")
			vEtc			= vRs("ETC")
			vEtc2			= vRs("ETC2")
			vItemListNm		= vRs("ITEM_LIST_NM")
			vHospitalType	= vRs("HOSPITAL_TYPE")
		End If 
		rsClose(vRs)
	Else
		vTxt = "등록"
	End If 
	vHospitalType	= nullFillWith(vHospitalType , "", False)
%>
<!doctype html>
<html lang="ko">
  <head>
    <!-- #include virtual = "/common/include/layout_header.asp"-->
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
                  <a href="">점검항목 등록</a>
                  <i class="bx bx-chevron-right"></i>
                  <a class="current" href="">온/습도 점검일지</a>
                </div>
                <div class="main-content">
                  <section class="section">
                    <div class="section-container">
                      <div class="section-title">
                        <h2>온/습도 점검일지 등록/수정</h2>
                      </div>
                      <div class="section-content">
                        <hr class="divider" />
                        <form class="form">
						  <input type="hidden" id="idx" name="idx" value="<%=vIdx%>">
						  <fieldset class="input-group flex-box">
                            <div class="input-field item1-5">
                              <label class="label">병원구분</label>
								<div class="select w-350">
								  <select id="hospitalType" name="hospitalType">
								  <option value="">선택</option>
								  <%=getCodeDB("select" , "P11" , vHospitalType)%>
								  </select>
								</div>
                            </div>
                          </fieldset>
                          <hr class="divider" />
                          <fieldset class="input-group flex-box">
                            <div class="input-field item1-5">
                              <label class="label">구역번호</label>
                              <div class="input">
                                <input id="areacode" name="areacode" type="text" placeholder="구역번호" value="<%=vAreaCode%>"/>
                              </div>
                            </div>
                          </fieldset>

                          <hr class="divider" />
						  <fieldset class="input-group flex-box">
                            <div class="input-field item1-5">
                              <label class="label">구역 명</label>
                              <div class="input">
                                <input id="areaname" name="areaname" type="text" placeholder="구역 명" value="<%=vAreaName%>"/>
                              </div>
                            </div>
                          </fieldset>
                          <hr class="divider" />
						  <fieldset class="input-group flex-box">
                            <div class="input-field item1-3">
                              <label class="label">적정온도습도</label>
                              <div class="input">
                                <input id="etc2" name="etc2" type="text" placeholder="적정온도습도 입력해주세요" value="<%=vEtc2%>" />
                              </div>
                            </div>
                          </fieldset>
						  <hr class="divider" />
						  <fieldset class="input-group flex-box">
							<div class="input-field">
							  <label class="label">구분항목</label>
							  <div class="flex-box x-small">
								<div class="select w-350">
								  <select id="item" name="item">
								  <option value="">선택</option>
								  <% 
									If isArray(vItemData) Then 
										For vI = 0 To UBound(vItemData,2)
											rw("<option value='" & vItemData(2,vI) & "'>" & vItemData(1,vI) & "</option>")
										Next
									End If 
								  %>
								  </select>
								</div>
								<div class="mw-100">
									<button type="button" class="button" onclick="itemAdd()">추가</button>
								</div>								
							  </div>
							  <div id="itemList" class="w-500 mt-30">
							  <%
							   vI = 0
								sql = "EXEC USP_CODE_TABLE_LIST '{P1}','{P2}';"
								sql = Replace(sql , "{P1}" , vItemList , vbTextCompare)
								sql = Replace(sql , "{P2}" , ","			, vbTextCompare)
								Set rs = getQueryDB(sql)
								If Not(rs.Eof Or rs.Bof) Then 
									Do Until rs.Eof 
									arr_itemlist = rs(4)
									arr_itemlistCode = rs(3)

							
							  %>
								<div class="area">
									<div class="flex-box between center my-10 px-20">
										<p><%=arr_itemlist%></p>
										<button type="button" class="button outline def w-auto" style="margin-left: 30px;" onclick="del(this,'<%=arr_itemlistCode%>')">삭제하기</button>
									</div>
									<hr class="divider my-10" alt="<%=vI%>"/>
								</div>
							  <%
										vI = vI + 1
										rs.MoveNext
									Loop 
								End If 
								rsClose(rs)
							    %>
							  </div>
							</div>
                          </fieldset>

                          <hr class="divider" />
						  <fieldset class="input-group flex-box">
                            <div class="input-field item1-3">
                              <label class="label">비고</label>
                              <div class="input">
                                <input id="etc" name="etc" type="text" placeholder="비고" value="<%=vEtc%>" />
                              </div>
                            </div>
                          </fieldset>

                          <hr class="divider" />
                          <div class="form-control">
                            <div class="right-box">
							  <a href="cart_temperature.asp" class="button outline def">돌아가기</a>
                              <button class="button" onclick="register()"><%=vTxt%>하기</button>
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

  <script>
	
	var arr = new Array();
	$(document).ready(function() {
		// 아이템리스트 배열로 넣기
		var itemlist = "<%=vItemList%>";
		if (itemlist != "") {			
			var split_list = itemlist.split(",");
			for (var i=0;i<split_list.length;i++ ) {
				arr.push(split_list[i]);
			}
		}

		$(document).on("click", ".newdel", function() {
			if (arr.indexOf($(this).attr("alt")) >= 0) {
				arr.splice(arr.indexOf($(this).attr("alt")), 1);
				$(this).closest(".area").remove();
			}
		});
	});

	
	function itemAdd() {
		var itemVal = $("#item").val();
		var itemNm = $("#item option:checked").text();

		if (itemVal == "") {
			sAlert("warning","항목을 선택해주세요");
		} else {
			if (arr.indexOf(itemVal) >= 0) {
				sAlert("warning","이미 등록된 항목입니다");
			} else {
				var html = "";
				html += "<div class=\"area\">";
				html += "	<div class=\"flex-box between center my-10 px-20\">";
				html += "	<p>" + itemNm + "</p>";
				html += "	<button type=\"button\" class=\"button outline def w-auto newdel\" style=\"margin-left: 30px;\" alt=\"" + itemVal + "\">삭제하기</button>";
				html += "	</div>";
				html += "	<hr class=\"divider my-10\" />";
				html += "</div>";
									
				arr.push(itemVal);
				$("#itemList").append(html);
			}
		}		
	}

	function register() {
		if ($("#areacode").val() == "") {
			sAlert("warning","구역번호를 입력해주세요");
		} else if ($("#areaname").val() == "") {
			sAlert("warning","구역명을 입력해주세요");
		} else if (arr.length == 0) {
			sAlert("warning","구분항목을 추가해주세요");
		} else if ($("#hospitalType").val() == "") {
			sAlert("warning","병원구분을 선택해주세요");
		} else {
			$.ajax({
				type : "GET",
				url : "proc.asp",
				data	: {
					command		:	"cart_temperatureRegister",
					idx			:	$("#idx").val(),
					areacode    :	$("#areacode").val(),
					areaname    :	$("#areaname").val(),
					itemlist	:	arr.join(),
					etc			:	$("#etc").val(),
					etc2		:	$("#etc2").val()
				},
				success : function(json){
					var jsonData	= eval(json);
					var resultCd	= jsonData[0].resultCode;
					
					if (resultCd == "S") {
						sAlertMove("success","요청 성공!","m","cart_temperature.asp<%=Session("rtnParam")%>");
					}
				},
				error : function(request, status, error){
					sAlert("error","처리도중 에러가 발생했습니다");
					return false;
				}
			});		//END-AJAX
		}
	}
	
	// 항목삭제
	function del(e,code) {
		if (arr.indexOf(code) >= 0) {
			arr.splice(arr.indexOf(code), 1);
			$(e).closest(".area").remove();
		}
	}

  </script>
  </body>
</html>
