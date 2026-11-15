<%@CodePage=65001%>
<% 
'=============================================
' 작 성 자 : 노 광 옹
' 작 성 일 : 2024-10-17
' 파일내용 : 설정-코드등록
'=============================================
Option Explicit 
Dim sParCode , currParName , currParUseYn
sParCode = nullFillWith(Request("parCode"),"",False)	
%>
<!-- #include virtual = "/common/library/common.asp"-->
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
        <div class="header-backdrop"></div>
        <div id="wrap">
          <div class="scroll-area">
            <main id="main">
              <div class="main-container">
			    <!-- #include virtual = "/common/include/layout_top.asp"-->
                <div class="breadcrumb">
                  <a href=""><i class="bx bx-home-alt-2"></i></a>
                  <i class="bx bx-chevron-right"></i>
                  <a href="">설정</a>
                  <i class="bx bx-chevron-right"></i>
                  <a class="current" href="">코드 등록</a>
                </div>
                <div class="main-content">
                  <section class="section">
                    <div class="section-container">
                      <div class="section-title">
                        <h2>코드 등록</h2>
                      </div>
					  <form name="dataForm" id="dataForm">
						<input type="hidden" name="command" id="command" value="Par_Add" title="명령어:기본 부모정보코드등록" style="background-color:lime;">
						<input type="hidden" name="currParCode" id="currParCode" value="<%=sParCode%>" style="background-color:lime;">
                      <div class="tab">
                        <div data-id="tab3" class="tab-content current">
                          <div class="section-content">
                            <div class="flex-box">
                              <div class="field">
                                <div class="card mb-20">
                                  <div class="card-title">
                                    <div class="flex-box center between">
                                      <h2>부모 정보 코드</h2>
                                      <button class="button def w-fit" id="addParCode">
                                        부모코드 신규 등록
                                        <i class="bx bx-chevron-right"></i>
                                      </button>
                                    </div>
                                  </div>
                                  <div class="table" style="height:500px;">
                                    <table>
                                      <thead>
                                        <tr>
                                          <th class="fit-cell">번호</th>
                                          <th>코드</th>
                                          <th>선택</th>
                                          <th class="fit-cell">사용여부</th>
                                        </tr>
                                      </thead>
                                      <tbody>
									  <%
										Dim sql , rs
										Dim OP_IDX, OP_PAR_CODE_NM, OP_PAR_CODE_VAL,OP_PAR_USE_YN , OP_PAR_ETC1 , OP_PAR_ETC2 , OP_PAR_ETC3
										sql = "EXEC USP_CODE_PAR_LIST '';"
										Set rs = getQueryDB(sql)
										If Not(rs.Eof Or rs.Bof) Then 
											Do Until rs.Eof
											OP_IDX				= rs("IDX")
											OP_PAR_CODE_NM		= rs("PAR_CODE_NM")
											OP_PAR_CODE_VAL		= rs("PAR_CODE_VAL")
											OP_PAR_USE_YN		= rs("USE_YN")
											OP_PAR_ETC1			= rs("SUB_CODE_NM1")
											OP_PAR_ETC2			= rs("SUB_CODE_NM2")
											OP_PAR_ETC3			= rs("SUB_CODE_NM3")

											If sParCode = OP_PAR_CODE_VAL Then 
												currParName		= OP_PAR_CODE_NM
												currParUseYn	= OP_PAR_USE_YN
											End If 
									  %>
                                        <tr onClick="btnComCodeList('<%=OP_PAR_CODE_VAL%>','<%=OP_PAR_CODE_NM%>','<%=OP_PAR_USE_YN%>','<%=OP_PAR_ETC1%>','<%=OP_PAR_ETC2%>','<%=OP_PAR_ETC3%>');void(0);" style="cursor:pointer;">
                                          <td class="fit-cell"><%=OP_IDX%></td>
                                          <td><%=OP_PAR_CODE_VAL%></td>
                                          <td><%=OP_PAR_CODE_NM%></td>
                                          <td class="fit-cell"><%=OP_PAR_USE_YN%></td>
                                        </tr>
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
                                <div class="card">
                                  <div class="card-title">선택된 부모 코드 수정</div>
                                  <div class="card-content">
                                    <div class="flex-box center">
                                      <label class="input-field item1-3">코드값</label>
                                      <div class="input readonly">
                                        <input type="text" name="parCode" id="parCode" value="부모 코드 자동생성" readonly />
                                      </div>
                                    </div>
                                    <hr class="divider my-10" />
                                    <div class="flex-box center">
                                      <label class="input-field item1-3">코드명</label>
                                      <div class="input">
                                        <input type="text" name="parCodeName" id="parCodeName" />
                                      </div>
                                    </div>
                                    <hr class="divider my-10" />
                                    <div class="flex-box center">
                                      <label class="input-field item1-3">사용여부</label>
										<div class="select" style="width:300px;float:left;">
										  <select name="parCodeUse" id="parCodeUse">
											<option value="">선택</option>
											<option value="Y">사용</option>
											<option value="N">사용안함</option>
										  </select>
										  </div>
                                    </div>
                                    <hr class="divider my-10" />
                                    <div class="flex-box center">
                                      <label class="input-field item1-3">보조 코드명1</label>
                                      <div class="input">
                                        <input type="text" name="parCodeEtc1" id="parCodeEtc1" />
                                      </div>
                                    </div>
                                    <hr class="divider my-10" />
                                    <div class="flex-box center">
                                      <label class="input-field item1-3">보조 코드명2</label>
                                      <div class="input">
                                        <input type="text" name="parCodeEtc2" id="parCodeEtc2" />
                                      </div>
                                    </div>
                                    <hr class="divider my-10" />
                                    <div class="flex-box center">
                                      <label class="input-field item1-3">보조 코드명3</label>
                                      <div class="input">
                                        <input type="text" name="parCodeEtc3" id="parCodeEtc3" />
                                      </div>
                                    </div>
                                  </div>
                                  <div class="card-footer right">
                                    <button type="submit" class="button px-50" id="btnParCode">부모코드 등록하기</button>
                                  </div>
                                </div>
                              </div>
                              <div class="field">
                                <div class="card mb-20">
                                  <div class="card-title">
                                    <div class="flex-box center between">
                                      <h2>자식 정보 코드</h2>
                                      <button class="button def w-fit" id="addComCode">
                                        자식코드 신규 등록
                                        <i class="bx bx-chevron-right"></i>
                                      </button>
                                    </div>
                                  </div>
                                  <div class="table" style="height:500px;">
                                    <table>
                                      <thead>
                                        <tr>
                                          <th class="fit-cell">번호</th>
                                          <th>코드</th>
                                          <th>선택</th>
                                          <th class="fit-cell">사용여부</th>
                                        </tr>
                                      </thead>
                                      <tbody id="comCodeListView">
										<!-- 자식정보 코드 -->
                                      </tbody>
                                    </table>
                                  </div>
                                </div>
                                <div class="card">
                                  <div class="card-title">선택된 자식 코드 수정</div>
                                  <div class="card-content">
                                    <div class="flex-box center">
                                      <label class="input-field item1-3">코드값</label>
                                      <div class="input readonly">
                                        <input type="text" name="comCode" id="comCode" value="" readonly />
                                      </div>
                                    </div>
                                    <hr class="divider my-10" />
                                    <div class="flex-box center">
                                      <label class="input-field item1-3">코드명</label>
                                      <div class="input">
                                        <input type="text" name="comCodeName" id="comCodeName" />
                                      </div>
                                    </div>
                                    <hr class="divider my-10" />
                                    <div class="flex-box center">
                                      <label class="input-field item1-3">사용여부</label>
										<div class="select">
										  <select name="comCodeUse" id="comCodeUse"  style="width:300px;float:left;">
											<option value="">선택</option>
											<option value="Y">사용</option>
											<option value="N">사용안함</option>
										  </select>
										  </div>
                                    </div>
                                    <hr class="divider my-10" />
                                    <div class="flex-box center">
                                      <label class="input-field item1-3">보조 코드명1</label>
                                      <div class="input">
                                        <input type="text" name="comCodeEtc1" id="comCodeEtc1" />
                                      </div>
                                    </div>
                                    <hr class="divider my-10" />
                                    <div class="flex-box center">
                                      <label class="input-field item1-3">보조 코드명2</label>
                                      <div class="input">
                                        <input type="text" name="comCodeEtc2" id="comCodeEtc2" />
                                      </div>
                                    </div>
                                    <hr class="divider my-10" />
                                    <div class="flex-box center">
                                      <label class="input-field item1-3">보조 코드명3</label>
                                      <div class="input">
                                        <input type="text" name="comCodeEtc3" id="comCodeEtc3" />
                                      </div>
                                    </div>
                                  </div>
                                  <div class="card-footer right">
                                    <button type="submit" class="button px-50" id="btnComCode">자식코드 등록하기</button>
                                  </div>
                                </div>
                              </div>
                            </div>
                          </div>
                        </div>
                        <div data-id="tab4" class="tab-content">TAB 4 CONTENT</div>
                        <div data-id="tab5" class="tab-content">TAB 5 CONTENT</div>
                      </div>
					  </form>
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

		if ($("#currParCode").val() != "")
		{
			btnComCodeList($("#currParCode").val() , "<%=currParName%>", "<%=currParUseYn%>", "", "", "");
		}

		//--------------------------------
		// 부모코드 등록 초기화
		//--------------------------------
		$("#addParCode").click(function(){
			//alert("부모 코드 등록");
			$("#btnParCode").text("부모코드 저장하기");
			$("#command").val("Par_Add");		
			$("#parCode").val("부모정보코드 자동생성");				
			$("#parCodeName").val("");
			$("#parCodeUse").val("Y");		
			$("#parCodeEtc1").val("");
			$("#parCodeEtc2").val("");
			$("#parCodeEtc3").val("");
		});


		//--------------------------------
		// 부모코드 수정 및 등록
		//--------------------------------
		$("#btnParCode").click(function(){
			var frm = document.dataForm;
			if(!setInputAlert(frm.parCodeName, "부모정보 코드명을")) return false;
			if(!setSelectAlert(frm.parCodeUse, "사용여부를")) return false;
			//$("#command").val("Par_Add");		
			frm.action = "./proc_code.asp";
			frm.target = "procFrame";
			frm.method = "POST";
			frm.submit();
		});

		//--------------------------------
		// 자식코드 등록 초기화
		//--------------------------------
		$("#addComCode").click(function() {		
			if ($("#parCodeName").val() == "")
			{
				alert("부모코드를 먼저 선택하여 주세요!");
				return false;
			}
			$("#btnComCode").text("자식코드 등록하기");
			$("#command").val("SON_ADD");	
			$("#comCode").val("자동생성");
			$("#comCodeUse").val("Y");	
			$("#comCodeName").val("");
		});

		//--------------------------------
		// 자식코드 수정 및 등록
		//--------------------------------
		$("#btnComCode").click(function(){
			if ($("#command").val() == "SON_ADD" || $("#command").val() == "Com_Mod" )
			{
				var frm = document.dataForm;
				if(!setInputAlert(frm.comCodeName, "자식정보 코드명을")) return false;
				if(!setSelectAlert(frm.comCodeUse, "사용여부를")) return false;
				
				frm.action = "./proc_code.asp";
				frm.target = "procFrame";
				frm.method = "POST";
				frm.submit();		
			}
			else {
				alert("부모정보코드를 선택 후 , 상단의 [자식코드 신규 등록]버튼을 먼저 클릭해 주세요!");
				return false;
			}

		});

	});		//End Ready


	//선택된 자식 정보 코드를 노출함
	function btnComCodeList(v , vn , vu , ve1 , ve2 , ve3) 
	{
		$.ajax({
			type : "GET",
			url : "proc_code.asp",
			data	: {
				command			:	"ComCodeList",
				parCode     	:	v
			},
			success : function(json){
				//뷰 삭제
				$("#comCodeListView").empty();

				var jsonData	= eval(json);
				var resultCd	= jsonData[0].header.resultCode;
				var listView	= "";

				$("#parCode").val(v);
				$("#currParCode").val(v);
				$("#parCodeName").val(vn);
				$("#parCodeUse").val(vu);
				$("#parCodeEtc1").val(ve1);
				$("#parCodeEtc2").val(ve2);
				$("#parCodeEtc3").val(ve3);

				$("#command").val("Par_Mod");						
				$("#btnParCode").text("부모코드 수정하기");

				if (resultCd == "S") {
					var resultData	= jsonData[0].result;


					//보여줄 데이터 draw
					for(var i = 0;i < resultData.length;i++){
						var idx		= resultData[i].idx;
						var codeNm	= resultData[i].codeNm;
						var codeVal	= resultData[i].codeVal;
						var useYn		= resultData[i].useYn;
						var codeEtc1	= resultData[i].codeEtc1;
						var codeEtc2	= resultData[i].codeEtc2;
						var codeEtc3	= resultData[i].codeEtc3;

						var bgcolor;
						if (useYn == "N") {
							bgcolor = "style='background-color: yellow;'";
						} else {
							bgcolor = "style=''";
						}

						listView		+= "<tr onClick=\"viewComDetailView('" + codeVal + "');void(0);\"  style=\"cursor:pointer;\">";
						listView		+= "	<td class=\"fit-cell\">"+ (i+1) +"</td>";
						listView		+= "	<td>"+codeVal+"</td>";
						listView		+= "	<td>"+codeNm+"</td>";
						listView		+= "	<td class=\"fit-cell\" " + bgcolor + ">"+useYn+"</td>";
						listView		+= "</tr>";

					}					
					$("#comCodeListView").append(listView);

					$("#comCode").val('');
					$("#comCodeName").val('');
					$("#comCodeEtc1").val('');
					$("#comCodeEtc2").val('');
					$("#comCodeEtc3").val('');
					$("#comCodeUse").val('Y');
				} else {
					//뷰 노출
					listView	+= "<td class=\"fit-cell\" colspan=\"4\">데이타가 없습니다</td>";
					$("#comCodeListView").append(listView);
				}
			},
			complete : function() {
				window.scrollTo({ top: 150, behavior: 'smooth' });
			},
			error : function(request, status, error){
				alert("처리도중 에러가 발생했습니다");
				return false;
			}
		});		//END-AJAX
	}

	//선택된 자식 정보 코드 상세보기
	function viewComDetailView(v) {
		$.ajax({
			type : "GET",
			url : "proc_code.asp",
			data	: {
				command			:	"ComCodeDetailView",
				comCode     	:	v
			},
			success : function(json){
				
				var jsonData	= eval(json);
				var resultCd	= jsonData[0].header.resultCode;
				var listView	= "";

				if (resultCd == "S") {
					var resultData	= jsonData[0].result;

					//보여줄 데이터 draw

						var idx		= resultData[0].idx;
						var codeNm	= resultData[0].codeNm;
						var codeVal	= resultData[0].codeVal;
						var useYn		= resultData[0].useYn;
						var codeEtc1	= resultData[0].codeEtc1;
						var codeEtc2	= resultData[0].codeEtc2;
						var codeEtc3	= resultData[0].codeEtc3;

						$("#comCode").val(codeVal);
						$("#comCodeName").val(codeNm);
						$("#comCodeEtc1").val(codeEtc1);
						$("#comCodeEtc2").val(codeEtc2);
						$("#comCodeEtc3").val(codeEtc3);					
						$("#comCodeUse").val(useYn);
						$("#command").val("Com_Mod");
						$("#btnComCode").text("자식코드 수정하기");
				}
			},
			error : function(request, status, error){
				alert("처리도중 에러가 발생했습니다");
				return false;
			}
		});		//END-AJAX
	}
//-->
</script>
<iframe  name="procFrame" id="procFrame" style="width:900px; height:500px;display:none;background-color:lime;"></iframe>