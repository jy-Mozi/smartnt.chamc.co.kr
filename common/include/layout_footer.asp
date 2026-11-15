<footer id="footer">
	<div class="footer-container">
		<div class="footer-content">
			<p class="copyright">Copyright Beplus Healthcare Inc.. All rights reserved.</p>
		</div>
	</div>
</footer>
<%
If Session("SNT_ADMIN_COMPANY") = "" Then 
	Call gotoAlert("error","세션이 종료되었습니다!\n로그인페이지로 이동합니다.","m","/")
End If 
%>