<%
'==============================
' 2026-11-17 접근 차단
' 로그인 체크 없이 그룹웨어 직원 목록을 엑셀로 내려받을 수 있어 막음.
' 다시 사용하려면 아래 3줄을 지우고 인증 체크를 추가할 것.
'==============================
Response.Status = "404 Not Found"
Response.End
%><!--#include virtual="/common/library/secret_config.asp"--><%
Response.Buffer = True
Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment;filename=ECM_"&left(Now(),7)&".xls"
'Response.Charset = "utf-8"

' DB 연결
Dim conn, rs, sql
Set conn = Server.CreateObject("ADODB.Connection")

' ODBC DSN 방식
conn.Open gloGwdbConnStr

' SQL
sql = "sELECT    u.login_id,    d.name AS dept,    u.name,    c.name AS company,    u.status   FROM go_users u     JOIN go_user_profiles up ON u.user_profile_id = up.id     JOIN go_companies c ON c.id = u.company_id and (c.id =42 or c.id =42 or c.id =21 or c.id =27 or c.id =10 or c.id =43  or c.id =25  or  c.id =44  )     LEFT JOIN go_dept_members dm ON dm.user_id = u.id     LEFT JOIN go_departments d ON d.id = dm.department_id     LEFT JOIN go_domain_codes c2 ON c2.id = up.position_id     LEFT JOIN go_domain_codes c1 ON c1.id = dm.duty_id     where u.status<>'DELETE' order by c.id , u.status"

Set rs = conn.Execute(sql)
%>
<html>
<head>

</head>
<body>

<table border="1">
    <tr style="background-color:#d9d9d9;">
        <th>login_id</th>
        <th>dept</th>
        <th>name</th>
		<th>company</th>
		<th>status</th>
    </tr>

<%
Do Until rs.EOF
%>
    <tr>
        <td><%=rs("login_id")%></td>
        <td><%=rs("dept")%></td>
        <td><%=rs("name")%></td>
		<td><%=rs("company")%></td>
		<td><%=rs("status")%></td>		

    </tr>
<%
    rs.MoveNext
Loop
%>

</table>

</body>
</html>

<%
rs.Close
Set rs = Nothing
conn.Close
Set conn = Nothing
%>