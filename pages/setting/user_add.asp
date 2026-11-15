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
Dim sCommand, sIdx, sFilter, sKeyword, page
sCommand = nullFillWith(Request("command"), "", False)
sIdx = nullFillWith(Request("idx"), "0", False)
sFilter = nullFillWith(Request("filter"), "", False)
sKeyword = nullFillWith(Request("keyword"), "", False)
page	= nullFillWith(Request("page"),"1",false)

Dim sql , rs, dataFlag, dataFlagStr
Dim OP_IDX, OP_ADMIN_ID , OP_EMPLOYEE_NUM, OP_ADMIN_NM , OP_ADMIN_LEVEL_NAME, OP_DEPT_NAME 
Dim OP_ADMIN_LEVEL, OP_ADMIN_PW, OP_ADMIN_COMPANY , OP_ADMIN_POSITION, OP_HOSPITAL_TYPE
If CInt(sIdx) > 0 Then 
	sql = "EXEC USP_ADMIN_DETAIL_VIEW '{P1}';"
	sql = Replace(sql , "{P1}" , sIdx , vbTextCompare)

	Set rs = getQueryDB(sql)
	If Not(rs.Eof Or rs.Bof) Then 
		OP_IDX				= rs("IDX")
		OP_ADMIN_ID			= rs("ADMIN_ID")
		OP_EMPLOYEE_NUM		= rs("EMPLOYEE_NUM")
		OP_ADMIN_NM			= rs("ADMIN_NM")
		OP_ADMIN_LEVEL_NAME	= rs("ADMIN_LEVEL_NAME")
		OP_DEPT_NAME		= rs("DEPT_NAME")
		OP_ADMIN_LEVEL		= rs("ADMIN_LEVEL")
		OP_ADMIN_PW			= rs("ADMIN_PW")
		OP_ADMIN_COMPANY	= nullFillWith(rs("ADMIN_COMPANY"),"" , False)
		OP_ADMIN_POSITION	= nullFillWith(rs("ADMIN_POSITION"),"", False)
		OP_HOSPITAL_TYPE	= nullFillWith(rs("HOSPITAL_TYPE"),"", False)
		dataFlag	= "MOD"
		dataFlagStr	= "수정하기"
	End If 
	rsClose(rs)
Else 
	dataFlag	= "ADD"
	dataFlagStr	= "등록하기"
End If 
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
                  <a href="">설정</a>
                  <i class="bx bx-chevron-right"></i>
                  <a class="current" href="">사용자 등록</a>
                </div>
                <div class="main-content">
                  <section class="section">
                    <div class="section-container">
                      <div class="section-title">
                        <h2>사용자 등록</h2>
                      </div>
                      <div class="section-content">
                        <hr class="divider" />
                        <form class="form" name="dataForm" id="dataForm">
						<input type="hidden" name="command" id="command" title="명령어" value="USER_DATA" style="background-color:lime;">
						<input type="hidden" name="idx" id="idx" title="인덱스" value="<%=sIdx%>" style="background-color:lime;">
						<input type="hidden" name="page" id="page" value="<%=page%>" title="인덱스" style="background-color:lime;">
                          <fieldset class="input-group flex-box">
                            <div class="input-field item1-3">
                              <label class="label">권한</label>
                              <div class="select">
                                <select name="level" id="level">
								<%=getCodeDB("select","P01",OP_ADMIN_LEVEL)%>
                                </select>
                              </div>
                            </div>
                          </fieldset>
                          <hr class="divider" />
                          <fieldset class="input-group flex-box">
                            <div class="input-field item1-3">
                              <label class="label">소속</label>
                              <div class="select">
                                <select name="company" id="company">
								<option value="">선택</option>
								<%=getCodeDB("select","P09",OP_ADMIN_COMPANY)%>
                                </select>
                              </div>
                            </div>
                          </fieldset>
                          <hr class="divider" />
						  <fieldset class="input-group flex-box">
                            <div class="input-field item1-3">
                              <label class="label">병원구분</label>
								<div class="select">
								  <select id="hospitalType" name="hospitalType">
								  <option value="">선택</option>
								  <%=getCodeDB("select" , "P11" , OP_HOSPITAL_TYPE)%>
								  </select>
								</div>
                            </div>
                          </fieldset>
                          <hr class="divider" />
                          <fieldset class="input-group flex-box">
                            <div class="input-field item1-3">
                              <label class="label">직급</label>
                              <div class="select">
                                <select name="position" id="position">
								<option value="">선택</option>
								<option value="P10001" <%=selectValue(OP_ADMIN_POSITION,"P10001","select")%>>팀장</option>
								<option value="P10002" <%=selectValue(OP_ADMIN_POSITION,"P10002","select")%>>담당</option>
                                </select>
                              </div>
                            </div>
                          </fieldset>
                          <hr class="divider" />
                          <fieldset class="input-group flex-box">
                            <div class="input-field item1-3">
                              <label class="label">아이디</label>
                              <div class="input">
                                <input name="id" id="id" type="text" value="<%=OP_ADMIN_ID%>" placeholder="아이디" maxlength="30"/>
                              </div>
                            </div>
                          </fieldset>
                          <hr class="divider" />
                          <fieldset class="input-group flex-box">
                            <div class="input-field item1-3">
                              <label class="label">이름</label>
                              <div class="input">
                                <input name="nm" id="nm" type="text" value="<%=OP_ADMIN_NM%>" placeholder="이름"  maxlength="30"/>
                              </div>
                            </div>
                          </fieldset>
                          <hr class="divider" />
						  <fieldset class="input-group flex-box">
                            <div class="input-field item1-3">
                              <label class="label">비밀번호</label>
                              <div class="input">
                                <input name="pw" id="pw" type="text" value="" placeholder="비밀번호" maxlength="300" />
                              </div>
                            </div>
                          </fieldset>
                          <hr class="divider" />
						  <fieldset class="input-group flex-box">
                            <div class="input-field item1-3">
                              <label class="label">사번</label>
                              <div class="input">
                                <input name="num" id="num" type="text"  value="<%=OP_EMPLOYEE_NUM%>" placeholder="사번" maxlength="30" />
                              </div>
                            </div>
                          </fieldset>
                          <hr class="divider" />
						  <fieldset class="input-group flex-box">
                            <div class="input-field item1-3">
                              <label class="label">부서</label>
                              <div class="input">
                                <input name="dept" id="dept" type="text"  value="<%=OP_DEPT_NAME%>" placeholder="부서" maxlength="30" />
                              </div>
                            </div>
                          </fieldset>
                          <hr class="divider" />
                          <div class="form-control">
                            <div class="right-box">
                              <a href="user_list.asp?1=1<%=Session("Param")%>" class="button outline def">돌아가기</a>
                              <button type="button" class="button" id="btnSubmit"><%=dataFlagStr%></button>
							  <%If dataFlag	= "MOD" Then %>
                              <button type="button" class="button" id="btnRemove">삭제하기</button>
							  <%End If %>
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
		if ($("#company").val() == "P09002")
		{
			document.getElementById("position").options[1].text = "점장";
		}
		//소속이 현대그린푸드인 경우 점장, 분당차병원인 경우 팀장으로 변경
		$("#company").change(function(){
			var sel = this.value;
			if (sel == "P09001")
				document.getElementById("position").options[1].text = "팀장";
			else if (sel == "P09002")
				document.getElementById("position").options[1].text = "점장";
		});

		$("#btnSubmit").click(function(){

			var frm = document.dataForm;				
			if(!setSelectAlert(frm.level, "권한을")) return false;
			if(!setSelectAlert(frm.company, "소속을")) return false;
			if(!setSelectAlert(frm.hospitalType, "병원구분을")) return false;
			if(!setSelectAlert(frm.position, "직급을")) return false;
			if(!setInputAlert(frm.id, "아이디를")) return false;
			if(!setInputAlert(frm.nm, "이름을")) return false;
			if(!setInputAlert(frm.pw, "비밀번호를")) return false;			
			frm.action = "./proc_user.asp";
			frm.target = "procFrame";
			frm.method = "POST";
			frm.submit();
		});

		$("#btnRemove").click(function(){
			Swal.fire({
				title: "선택된 데이타를 삭제하시겠습니까?",
				text: "",
				icon: "warning",
				showCancelButton: true,
				confirmButtonColor: "#3085d6",
				cancelButtonColor: "#d33",
				cancelButtonText: "아니오!",
				confirmButtonText: "예"
			}).then((result) => {
				if (result.isConfirmed) {
					var frm = document.dataForm;	
					frm.command.value = "USER_REMOVE";
					frm.page.value = "1";
					frm.action = "./proc_user.asp";
					frm.target = "procFrame";
					frm.method = "POST";
					frm.submit();
				}
			});

		});
	});		//End Ready
//-->
</script>
<iframe  name="procFrame" id="procFrame" style="width:900px; height:500px;display:block;background-color:#lime;"></iframe>