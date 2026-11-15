'<%
''response.write request.serverVariables("REMOTE_HOST")
'dim DebugIP : DebugIP = "115.140.110.58"		'(제대혈 샘자리) 119.194.200.148  ' 119.194.200.147		' 115.140.110.58
'dim DebugID : DebugID = "test_kjh"	'"test_kjh"	'"kjh"	'testerJK	'CHANtest		smh81		' CHAN248		'stjf282
'dim isDebug : isDebug = False
''isDebug = True
'
'
'Dim arr_NoDebug_PathInfo(2)
'arr_NoDebug_PathInfo(0) = "/default.asp"
'arr_NoDebug_PathInfo(1) = "/permission/request_object.asp"
'arr_NoDebug_PathInfo(2) = "/pp_ax_hub.asp"			'//	sms결제
'
'
'Function IS_NoDebug_Path()
'	Dim rtnVal : rtnVal = False
'	Dim iRow
'	
'	For iRow=0 To UBound(arr_NoDebug_PathInfo) - 1
'		If InStr(Request.ServerVariables("PATH_INFO"), arr_NoDebug_PathInfo(iRow)) >= 1 Then
'			rtnVal = True
'			Exit For
'		End If
'	Next
'	
'	IS_NoDebug_Path = rtnVal
'End Function
'
'
'Function Is_Debuging()
'	Dim rtnVal : rtnVal = False
'	
'	If isDebug = True And Is_Debuger() = True And IS_NoDebug_Path() = False Then
'		rtnVal = True
'	End If
'
'	Is_Debuging = rtnVal
'End Function
'
'
'Function Is_Debuger()
'	Dim rtnVal : rtnVal = False
'	
'	'www.cymedi.com/new_ver2/admin
''	If InStr(Request.ServerVariables("PATH_INFO"), "/new_ver2/admin/") = 0 Then 
''		'//	Front
''		If request.serverVariables("REMOTE_HOST") = DebugIP Then
''			rtnVal = True
''		End If
''	Else
''		'//	BackOffice
''		If request.serverVariables("REMOTE_HOST") = DebugIP And DebugID = session("adminid") Then
''			rtnVal = True
''		End If
''	End If
'	
'	
''	If request.serverVariables("REMOTE_HOST") = DebugIP And _
''		 (DebugID = session("adminid") Or DebugID = session("id") Or InStr(session("id"), "test") >= 1) Then	
'	If request.serverVariables("REMOTE_HOST") = DebugIP And _
'		 DebugID = session("UserID") Then
'		
'		rtnVal = True
'	End If
'	
'	Is_Debuger = rtnVal
'End Function
'
'
'
'
'
'if Is_Debuging() = True Then
''	response.write "<font color='#330099' size='2'>"
''	response.write ""
''	response.write "</font>"
''  response.write "<br>"
'	
''	response.write right(DateString(now()),12) & "-" & RndMakerNumberA(4)
'	'response.write session.sessionID & "-" & RndMakerNumberA(4)
''	response.write "<br>"
'	
'	response.write "<font color='#CC0000'>"& Request.Servervariables("LOCAL_ADDR") & Request.ServerVariables("PATH_INFO") &"</font>"
'	response.write "<br>"
'	response.write "<font size='2'>ref_url : " & Request.Servervariables("HTTP_REFERER") &"</font>"
'	response.write "<br>"
''	response.write "<font color='#CC0000'>server.mapPath : "& server.mapPath(".") &"</font>"
''	response.write "<br>"
'	
'	response.write "<font color='#9933FF' size='2'>"
'	response.write "UserID : " & session("User_ID")
'	response.write ", UserName : " & session("User_Name")
'	response.write ", insurance_code : " & session("insurance_code")
'	response.write ", branch_code : " & session("branch_code")
'	response.write ", mLevel : " & session("mLevel")
'	response.write ", h_phone : " & session("h_phone")
'	response.write ", Date() : " & Date()
'	'response.write ", time : " & time()
'	response.write "</font>"
'	response.write "<br>"
'	
'	
'	call DisplayRequests() 
'	'Call Display_ServerVariable()
'end if
'
'
'
'Function Dbg_strSql(ByVal argSql)
'	if Is_Debuging() = True And Trim(argSql) <> "" Then
'		argSql = Replace(argSql, "select ", "<font color='blue'>select</font> ")
'		argSql = Replace(argSql, "Select ", "<font color='blue'>Select</font> ")
'		argSql = Replace(argSql, "SELECT ", "<font color='blue'>SELECT</font> ")
'		
'		argSql = Replace(argSql, " from ", " <font color='blue'>from</font> ")
'		argSql = Replace(argSql, " From ", " <font color='blue'>From</font> ")
'		argSql = Replace(argSql, " FROM ", " <font color='blue'>FROM</font> ")
'		
'		argSql = Replace(argSql, " where ", " <font color='blue'>where</font> ")
'		argSql = Replace(argSql, " Where ", " <font color='blue'>Where</font> ")
'		argSql = Replace(argSql, " WHERE ", " <font color='blue'>WHERE</font> ")
'
'		argSql = Replace(argSql, " order by ", " <font color='blue'>order by</font> ")
'		argSql = Replace(argSql, " Order by ", " <font color='blue'>Order by</font> ")
'		argSql = Replace(argSql, " ORDER BY ", " <font color='blue'>ORDER BY</font> ")
'
'		argSql = Replace(argSql, " group by ", " <font color='blue'>group by</font> ")
'		argSql = Replace(argSql, " Group by ", " <font color='blue'>Group by</font> ")
'		argSql = Replace(argSql, " GROUP BY ", " <font color='blue'>GROUP BY</font> ")
'		
'		argSql = Replace(argSql, " inner join ", " <font color='blue'>inner join</font> ")
'		argSql = Replace(argSql, " INNER JOIN ", " <font color='blue'>INNER JOIN</font> ")
'		
'		argSql = Replace(argSql, " right outer join ", " <font color='blue'>right outer join</font> ")
'		argSql = Replace(argSql, " RIGHT OUTER JOIN ", " <font color='blue'>RIGHT OUTER JOIN</font> ")
'		
'		argSql = Replace(argSql, " left outer join ", " <font color='blue'>left outer join</font> ")
'		argSql = Replace(argSql, " LEFT OUTER JOIN ", " <font color='blue'>LEFT OUTER JOIN</font> ")
'
'		argSql = Replace(argSql, " left join ", " <font color='blue'>left join</font> ")
'		argSql = Replace(argSql, " LEFT JOIN ", " <font color='blue'>LEFT JOIN</font> ")
'		
'		argSql = Replace(argSql, " on ", " <font color='blue'>on</font> ")
'		argSql = Replace(argSql, " ON ", " <font color='blue'>ON</font> ")
'		
'		argSql = Replace(argSql, " as ", " <font color='blue'>as</font> ")
'		argSql = Replace(argSql, " AS ", " <font color='blue'>AS</font> ")
'		
'		argSql = Replace(argSql, " distinct ", " <font color='blue'>distinct</font> ")
'		argSql = Replace(argSql, " Distinct ", " <font color='blue'>Distinct</font> ")
'		argSql = Replace(argSql, " DISTINCT ", " <font color='blue'>DISTINCT</font> ")
'
'		argSql = Replace(argSql, " distinct(", " <font color='blue'>distinct</font>(")
'		argSql = Replace(argSql, " Distinct(", " <font color='blue'>Distinct</font>(")
'		argSql = Replace(argSql, " DISTINCT(", " <font color='blue'>DISTINCT</font>(")
'
'		argSql = Replace(argSql, " max(", " <font color='blue'>max</font>(")
'		argSql = Replace(argSql, " Max( ", " <font color='blue'>Max</font>(")
'		argSql = Replace(argSql, " MAX( ", " <font color='blue'>MAX</font>(")
'
'		argSql = Replace(argSql, " case when ", " <font color='blue'>case when</font> ")
'		argSql = Replace(argSql, " Case when ", " <font color='blue'>Case when</font> ")
'		argSql = Replace(argSql, " Case When ", " <font color='blue'>Case When</font> ")
'		argSql = Replace(argSql, " CASE WHEN ", " <font color='blue'>CASE WHEN</font> ")
'
'		argSql = Replace(argSql, " then ", " <font color='blue'>then</font> ")
'		argSql = Replace(argSql, " Then ", " <font color='blue'>Then</font> ")
'		argSql = Replace(argSql, " THEN ", " <font color='blue'>THEN</font> ")
'
'		argSql = Replace(argSql, " end ", " <font color='blue'>end</font> ")
'		argSql = Replace(argSql, " End ", " <font color='blue'>End</font> ")
'		argSql = Replace(argSql, " END ", " <font color='blue'>END</font> ")
'
'		argSql = Replace(argSql, " and ", " <font color='blue'>and</font> ")
'		argSql = Replace(argSql, " And ", " <font color='blue'>And</font> ")
'		argSql = Replace(argSql, " AND ", " <font color='blue'>AND</font> ")
'		
'		
'		argSql = Replace(argSql, "insert ", "<font color='blue'>insert</font> ")
'		argSql = Replace(argSql, "Insert ", "<font color='blue'>Insert</font> ")
'		argSql = Replace(argSql, "INSERT ", "<font color='blue'>INSERT</font> ")
'
'		argSql = Replace(argSql, "into ", "<font color='blue'>into</font> ")
'		argSql = Replace(argSql, "Into ", "<font color='blue'>Into</font> ")
'		argSql = Replace(argSql, "INTO ", "<font color='blue'>INTO</font> ")
'		
'		argSql = Replace(argSql, "values ", "<font color='blue'>values</font> ")
'		argSql = Replace(argSql, "Values ", "<font color='blue'>Values</font> ")
'		argSql = Replace(argSql, "VALUES ", "<font color='blue'>VALUES</font> ")
'
'		argSql = Replace(argSql, "values(", "<font color='blue'>values</font>(")
'		argSql = Replace(argSql, "Values(", "<font color='blue'>Values</font>(")
'		argSql = Replace(argSql, "VALUES(", "<font color='blue'>VALUES</font>(")
'		
'		
'		argSql = Replace(argSql, "update ", "<font color='blue'>update</font> ")
'		argSql = Replace(argSql, "Update ", "<font color='blue'>Update</font> ")
'		argSql = Replace(argSql, "UPDATE ", "<font color='blue'>UPDATE</font> ")
'
'		argSql = Replace(argSql, "set ", "<font color='blue'>set</font> ")
'		argSql = Replace(argSql, "Set ", "<font color='blue'>Set</font> ")
'		argSql = Replace(argSql, "SET ", "<font color='blue'>SET</font> ")
'
'
'		argSql = Replace(argSql, "delete ", "<font color='blue'>delete</font> ")
'		argSql = Replace(argSql, "Delete ", "<font color='blue'>Delete</font> ")
'		argSql = Replace(argSql, "DELETE ", "<font color='blue'>DELETE</font> ")
'		
'		argSql = Replace(argSql, Chr(13) & Chr(10), "<br> ")
'		
'		response.write "<br>"
'		response.write argSql
'		response.write "<br>"
'	End if
'End Function
'
'
'function DbgCmdItem(argCmd)
'  if Is_Debuging() = True then
'      with argCmd
'        response.write "<br>"
'        response.write .CommandText & "<br>"
'        dim iDbg
'        for iDbg=0 to .Parameters.Count - 1
'          response.write .Parameters(iDbg).Name & " = " & .Parameters(iDbg).Value
'          response.write "<br>"
'        next
'        response.write "<br>"
'      end with
'  end if
'end function
'
'function DbgCmd(argCmd, isViewType, isBr)
'  if Is_Debuging() = True then
'    dim str : str = ""
'    dim value 
'    dim tType
'    dim iDbg
'		Dim commandText
'    
'    with argCmd
'      str =  str &  "<br>"
'      
'			commandText = Replace(.CommandText, "{", "")
'			commandText = Replace(commandText, "}", "")
'			commandText = Replace(commandText, "call", "")
'			'commandText = Replace(commandText, "(", "")
'			'commandText = Replace(commandText, ")", "")
'			'commandText = Replace(commandText, ",", "")
'			commandText = Replace(commandText, "?", "")
'			str =  str &  commandText & " &nbsp;&nbsp;&nbsp; "
'			
'      if isBr = true then str =  str &  "<br>"
'      
'      for iDbg=0 to .Parameters.Count - 1
'        value = .Parameters(iDbg).Value 
'				
'        tType = .Parameters(iDbg).Type
'        if tType=129 or tType=0 or tType=201 or tType=203 or tType=200 or tType=12 or tType=202 or tType=130 then
'          value = "'" & value & "'"
'        end if
'        
'        if isViewType = true then
'          str = str & "(" & getCmdTypeEnum(tType) & ") "
'        end if
'        
'        
'				'If Trim(value) <> "" And value <> "''" And Trim(.Parameters(iDbg).Name) <> "" Then
'				If Trim(value) <> "" And value <> "''" And Trim(.Parameters(iDbg).Name) <> "" _
'					 Or .Parameters(iDbg).Direction = 4 Or .Parameters(iDbg).Direction = 2 Then
'							str =  str & .Parameters(iDbg).Name &"="& value
'							
'							If .Parameters(iDbg).Direction = 2 Then
'								str =  str & "{out}"
'							End If
'				End If
'
'				
'				if isBr = true then
'          str =  str &  "<br>"
'        else
'          If iDbg <> .Parameters.Count - 1 then
'						If Trim(value) <> "" And value <> "''" And Trim(.Parameters(iDbg).Name) <> "" _
'							 Or .Parameters(iDbg).Direction = 4 Or .Parameters(iDbg).Direction = 2 Then
'								str =  str & ", "
'						End If
'					End if
'        end if
'      next
'      
'      response.write str & "<br>"
'    end with
'  end if
'end function
'
'
'
'
'
'
'
'
'function getCmdTypeEnum(argType)
'  dim rtnVal
'  
'  if argType = null or isNull(argType) then
'    rtnVal = "Null"
'  end if
'  
'  select case argType
'    case 0x2000 : rtnVal = "AdArray"
'    case 20 : rtnVal = "adBigInt"
'    case 128 : rtnVal = "adBinary"
'    case 11 : rtnVal = "adBoolean"
'    case 8 : rtnVal = "adBSTR"
'    case 136 : rtnVal = "adChapter"
'    case 129 : rtnVal = "adChar"
'    case 6 : rtnVal = "adCurrency"
'    case 7 : rtnVal = "adDate"
'    case 133 : rtnVal = "adDBDate"
'    case 134 : rtnVal = "adDBTime"
'    case 135 : rtnVal = "adDBTimeStamp"
'    case 14 : rtnVal = "adDecimal"
'    case 5 : rtnVal = "adDouble"
'    case 0 : rtnVal = "adEmpty"
'    case 10 : rtnVal = "adError"
'    case 64 : rtnVal = "adFileTime"
'    case 72 : rtnVal = "adGUID"
'    case 9 : rtnVal = "adIDispatch"
'    case 3 : rtnVal = "adInteger"
'    case 13 : rtnVal = "adIUnknown"
'    case 205 : rtnVal = "adLongVarBinary"
'    case 201 : rtnVal = "adLongVarChar"
'    case 205 : rtnVal = "adLongVarBinary"
'    case 203 : rtnVal = "adLongVarWChar"
'    case 131 : rtnVal = "adNumeric"
'    case 138 : rtnVal = "adPropVariant"
'    case 4 : rtnVal = "adSingle"
'    case 2 : rtnVal = "adSmallInt"
'    case 16 : rtnVal = "adTinyInt"
'    case 21 : rtnVal = "adUnsignedBigInt"
'    case 19 : rtnVal = "adUnsignedInt"
'    case 18 : rtnVal = "adUnsignedSmallInt"
'    case 17 : rtnVal = "adUnsignedTinyInt"
'    case 132 : rtnVal = "adUserDefined"
'    case 204 : rtnVal = "adVarBinary"
'    case 200 : rtnVal = "adVarChar"
'    case 12 : rtnVal = "adVariant"
'    case 139 : rtnVal = "adVarNumeric"
'    case 202 : rtnVal = "adVarWChar"
'    case 130 : rtnVal = "adWChar"
'    case else : response.write "열거값에없음" : response.end
'  end select
'  
'  getCmdTypeEnum = rtnVal
'end function
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'sub DisplayRequests()
'  if Is_Debuging() = True then
'    dim strForm, strQuery, item
'    
'    strForm = "<table width='100%' border='1' bordercolor='#CCCCCC' cellpadding='0' cellspacing='0'>"
'    strForm = strForm & "<tr><td>"
'    strForm = strForm & "<font color='#993333' size='2'>::: Request.Form :::</font> <br>"
'    strForm = strForm & "<font size='2'>"
'    for each item in Request.Form
'      strForm = strForm &   item & " = " & request.Form(item) & "<br>"
'    next
'    strForm = strForm & "</font>"
'    strForm = strForm & "<br></td></tr>"
'    strForm = strForm & "</table>"
'    
'    
'    strQuery = "<table width='100%' border='1' bordercolor='#CCCCCC'  cellpadding='0' cellspacing='0'>"
'    strQuery = strQuery & "<tr><td>"
'    strQuery = strQuery & "<font color='#993333' size='2'>::: Request.QueryString :::</font> <br>"
'    strForm = strForm & "<font size='2'>"
'    for each item in Request.QueryString
'      strQuery = strQuery & item & " = " & request.QueryString(item) & "<br>"
'    next
'    strQuery = strQuery & "</font>"
'    strQuery = strQuery & "<br></td></tr>"
'    strQuery = strQuery & "</table>"
'    
'    response.write strForm & strQuery
'  end if
'end sub
'
'
'' 현재페이지명만 얻기
'function getPagename()
'    dim rtnPageName
'    
'    Path=Request.ServerVariables("URL")
'    Path_s = StrReverse(Path)              
'    
'    if instr(Path_s,"/") > 0 then       
'        Path_sLine = instr(Path_s,"/")
'    else
'        Path_sLine = 0
'    end if
'    
'    Path_left = Left(Path_s, Path_sLine-1)
'    rtnPageName = StrReverse(Path_left)    
'    
'    getPagename = rtnPageName
'end function
'
'
'
'
'
'
'
'Sub Display_ServerVariable()
'	Dim strKey
'	
'	response.write "<TABLE BORDER='1'> "
'	response.write "	<TR><TD><B>Server Variable</B></TD><TD><B>Value</B></TD></TR> "
'	
'	For Each strKey In Request.ServerVariables
'		response.write "<TR> "
'		response.write "	<TD>" & strKey & "</TD> "
'		response.write "	<TD>" & Request.ServerVariables(strKey) & "</TD> "
'		response.write "</TR> "
'	Next
'	
'	response.write "</TABLE> "
'end sub
'
'
'
'
'
'
'
'
'
'
'
'%>
'
'
