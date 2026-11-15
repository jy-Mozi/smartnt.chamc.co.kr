<%@CodePage=65001%>
<% 
'=============================================
' 작 성 자 : 노 광 옹
' 작 성 일 : 2024-10-17
' 파일내용 : 사용자 리스트
'=============================================
Option Explicit 
%>
<!-- #include virtual = "/common/library/common.asp"-->
<%
Dim sCommand, sIdx, sFilter, sKeyword, page , sLevel , sCompany , sPosition
sCommand = nullFillWith(Request("command"), "", False)
sIdx = nullFillWith(Request("idx"), "", False)
sFilter = nullFillWith(Request("filter"), "", False)
sKeyword = nullFillWith(Request("keyword"), "", False)
page	= nullFillWith(Request("page"),"1",false)
sLevel		= nullFillWith(Request("level"), "", False)
sCompany	= nullFillWith(Request("company"), "", False)
sPosition	= nullFillWith(Request("position"), "", False)

Dim sql , rs, arrData , i 

Dim listSize , pageSize , totCnt , totPage
Dim vTemp , vNumber 
listSize	= 10
pageSize    = 10
Session("Param") = "&filter=" & sFilter & "&keyword=" & sKeyword & "&level="&sLevel&"&position="&sPosition&"&company=" & sCompany

sql = "EXEC USP_ADMIN_LIST_PAGEING '{P1}','{P2}','{P3}','{P4}','{P5}','{P6}','{P7}';"
sql = Replace(sql , "{P1}", page , vbTextCompare)
sql = Replace(sql , "{P2}", listSize , vbTextCompare)
sql = Replace(sql , "{P3}", sLevel , vbTextCompare)
sql = Replace(sql , "{P4}", sCompany , vbTextCompare)
sql = Replace(sql , "{P5}", sPosition , vbTextCompare)
sql = Replace(sql , "{P6}", sFilter , vbTextCompare)
sql = Replace(sql , "{P7}", sKeyword , vbTextCompare)


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
                        <div class="table-search">
                          <form class="form" name="dataForm" id="dataForm">
						<input type="hidden" name="command" id="command" value="UserAdd" title="명령어:기본 부모정보코드등록"style="background-color:lime;">
						<input type="hidden" name="idx" id="idx" value="" title="인덱스" style="background-color:lime;">
						<input type="hidden" name="page" id="page" value="<%=page%>" title="인덱스" style="background-color:lime;">
                            <div class="select">
                              <select name="level" id="level">
							    <option value="">권한</option>
								<%=getCodeDB("select","P01",sLevel)%>
                              </select>
                            </div>
                            <div class="select">
                              <select name="company" id="company">
							    <option value="">소속</option>
								<%=getCodeDB("select","P09",sCompany)%>
                              </select>
                            </div>
                            <div class="select">
                              <select name="position" id="position">
							    <option value="">직급</option>
								<%=getCodeDB("select","P10",sPosition)%>
                              </select>
                            </div>
                            <div class="select">
                              <select name="filter" id="filter">
							    <option value="">선택</option>
                                <option value="nm" <%=selectValue("nm" , sFilter,"select" )%> >이름</option>
                                <option value="id" <%=selectValue("id" , sFilter,"select" )%> >아이디</option>
                              </select>
                            </div>
                            <div class="input">
                              <input type="search" name="keyword" id="keyword" value="<%=sKeyword%>" placeholder="검색어를 입력하세요." />
                            </div>
                            <button type="button" class="button" id="btnSearch">검색</button>
                            <!--<button type="button" class="button outline"><i class="bx bx-check"></i>전체</button>-->
                          </form>
                        </div>
                        <div class="table">
                          <table>
                            <thead>
                              <tr>
                                <th class="fit-cell">번호</th>
                                <th class="fit-cell">아이디</th>
                                <th class="fit-cell">이름</th>
                                <th class="fit-cell">소속</th>
                                <th class="fit-cell">부서</th>
                                <th class="fit-cell">병원구분</th>
                                <th class="fit-cell">직급</th>
                              </tr>
                            </thead>
                            <tbody>
							<%
							Dim OP_IDX, OP_ADMIN_ID , OP_EMPLOYEE_NUM, OP_ADMIN_NM , OP_ADMIN_LEVEL_NAME, OP_DEPT_NAME
							Dim OP_COMPANY, OP_POSITION, OP_HOSPITAL_TYPE, OP_POSITION_STR
							IF isArray(arrData) Then
							vTemp = 1
							vNumber = totCnt - listSize * (page - 1)
								For i = 0 To UBound(arrData, 2)
									OP_IDX				= arrData(2 , i)
									OP_ADMIN_ID			= arrData(3 , i)
									OP_EMPLOYEE_NUM		= arrData(6 , i)
									OP_ADMIN_NM			= arrData(4 , i)
									OP_ADMIN_LEVEL_NAME	= arrData(12 , i)
									OP_DEPT_NAME		= arrData(7 , i)	
									OP_COMPANY			= arrData(9 , i)
									OP_POSITION			= arrData(10 , i)
									OP_HOSPITAL_TYPE	= arrData(11 , i)
									OP_POSITION_STR		= arrData(16 , i)
							%>
                              <tr onClick="btnDetail('<%=OP_IDX%>');" style="cursor:pointer;">
                                <td class="fit-cell"><%=vNumber%></td>
                                <td class="fit-cell"><%=OP_ADMIN_ID%></td>
                                <td class="fit-cell"><%=OP_ADMIN_NM%></td>
                                <td class="fit-cell"><%=getCodeDB("name" , "P09" , OP_COMPANY)%></td>
                                <td class="fit-cell"><%=OP_DEPT_NAME%></td>
                                <td class="fit-cell"><%=getCodeDB("name" , "P11" , OP_HOSPITAL_TYPE)%></td>
                                <td class="fit-cell"><%=OP_POSITION_STR%></td>
                              </tr>							
							<%
									vNumber = vNumber - 1
									vTemp = vTemp + 1
								Next 
							Else 
								rw("<tr><td colspan=""6"">등록된 데이타가 없습니다</td></tr>")
							End If 
							%>

                            </tbody>
                          </table>
                        </div>
                        <div class="table-control" >
                          <div class="left-box">
							<%=getListPaging(page , Session("Param") , totCnt, totPage)%>
                          </div>
                          <div class="right-box">
                            <button type="button" class="button" id="btnSubmit">등록하기</a>
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
			//if(!setInputAlert(frm.keyword, "검색어를")) return false;
			//if(!setSelectAlert(frm.filter, "검색항목을")) return false;
			
			frm.action = "./user_list.asp";
			frm.target = "_self";
			frm.method = "POST";
			frm.submit();				
		});

		$("#btnSubmit").click(function(){
			var frm = document.dataForm;
			frm.page.value = "1";
			frm.action = "./user_add.asp";			
			frm.target = "_self";
			frm.method = "POST";
			frm.submit();	
		});
	});		//End Ready

	function btnDetail(v) {
		$("#idx").val(v);
		$("#btnSubmit").trigger("click");
	}
//-->
</script>
<iframe  name="procFrame" id="procFrame" style="width:900px; height:500px;display:none;background-color:lime;"></iframe>