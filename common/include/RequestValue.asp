<%

'																
'	넘어온 값 체크												
																
Dim vItem
 For Each vItem In Request.Form
	Response.Write "<font color='#006600'>[Form]</font> " & vItem & " = " & Request.Form(vItem) & "<BR>"
 Next
 For Each vItem In Request.QueryString
	  Response.Write "<font color='#CC6633'>[QueryString]</font> " & vItem & " = " & Request.QueryString(vItem) & "<BR>"
 Next
 '
 '
 '
%>
