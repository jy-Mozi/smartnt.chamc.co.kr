<%@CodePage=65001%>
<% 
'=============================================
' 작 성 자 : 노 광 옹
' 작 성 일 : 2024-10-17
' 파일내용 : 로그인
'=============================================
Option Explicit 
%>
<!-- #include virtual = "/common/library/common.asp"-->
<!doctype html>
<html lang="ko">
  <head>
    <!-- #include virtual = "/common/include/layout_header.asp"-->
  </head>
  <body>
    <div id="root" data-page="login">
      <div id="app">
        <main id="main">
          <section class="login-section">
            <div class="section-container">
              <div class="section-title">
                <h1 class="logo">
                  <img class="logo-image" src="/assets/images/logo_bk.svg"/>
                </h1>
				<p class="description">분당차병원 스마트 영양관리</p>
              </div>
              <div class="section-content">
                <form class="form" name="dataForm" id="dataForm">
				<input type="hidden" name="command" id="command" title="명령어" value="LoginChk" style="background-color:lime;">
                  <fieldset class="input-group">
                    <div class="input">
                      <input type="text" name="id" placeholder="ID"/>
                    </div>
                    <div class="input">
                      <input type="password" name="password" placeholder="PW" autocomplete="false"/>
                    </div>
                  </fieldset>
                  <hr class="divider" />
                  <button class="button login-button" id="btnLogin" >로그인</button>
                </form>
              </div>
            </div>
          </section>
        </main>
      </div>
    </div>
  </body>
</html>
<script type="text/javascript">
<!--
	$(document).ready(function(){

		$("#btnLogin").click(function(){

			var frm = document.dataForm;				
			if(!setInputAlert(frm.id, "아이디를")) return false;			
			if(!setInputAlert(frm.password, "비밀번호를")) return false;			
			frm.action = "./proc.asp";
			frm.target = "procFrame";
			frm.method = "POST";
			frm.submit();
		});
	});		//End Ready
//-->
</script>

<iframe  name="procFrame" id="procFrame" style="width:900px; height:500px;display:block;background-color:#lime;"></iframe>