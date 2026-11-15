<%
'-----------------------------------
' split 
'-----------------------------------
function split_cut (str,strchk,num)
    dim arr_str 
    dim i ,cCnt
    

    if trim(str) = "" then 
        split_cut = ""
        exit function			
    end if 

    if isNull(str) Then 
        split_cut = ""
        exit function			
    end if
    
    if len(str) = 0 Then 
        split_cut = ""
        exit function			
    end if

    cCnt = 0

    for i = 1 to len(str)
        if mid(str,i,1) = strchk then 
            cCnt = cCnt + 1
        end if 
    next 
    
    if cCnt < cInt(num) Then 
        'rw(cCnt & "-" & num)
        split_cut = ""
        exit function 
    end if 
    


    'split할 구분자가 있는지 체크 함 
    if instr(str,strchk) > 1  then 

        arr_str = split(str,strchk)
        split_cut = arr_str(num)
    else
        split_cut = ""
    end if 
end function

function selectvalue(str1,str2,ftype)

  if trim(str1) = trim(str2) then
   
   if ftype = "checkbox" or ftype = "radio" then
    selectvalue=  " checked"
   elseif ftype = "select" then
    selectvalue= " selected"
   ElseIf ftype = "current" Then 
    selectvalue= " current "
   ElseIf ftype = "data-selected" Then 
    selectvalue= " data-selected "
   end if
  end if
end Function


Function time24(strValue)
  time24 = Right("0"&strValue,2)
End function


'//==============================================
'// 기 능 : 날짜 int 형을 날짜 
'//         20121020  => 2012-10-20 으로 변경해 줌
'//==============================================
Function convertDate(regDt , strSplit)
    If Len(Trim(regDt)) = 8 Then   
		If strSplit = "str" Then 
	        convertDate = Left(regDt,4) & "년" & Mid(regDt,5,2) & "월" & Mid(regDt,7) & "일"
		Else 
		    convertDate = Left(regDt,4) & strSplit & Mid(regDt,5,2) & strSplit & Mid(regDt,7)  
		End If 

    Else 
        convertDate = regDt
    End If 
End Function 

Function blankDelete(str)
    Dim i , txt 
    txt = ""
    For i = 1 To Len(str)


        If Mid(Str , i , 1) = " " Then 
            txt = txt & ""
        Else 
            txt = txt & Mid(Str , i , 1) 
        End If 
    Next 

    blankDelete = txt
End Function 

'년월일 시분초
Function strDateTime ()	
	strDateTime	= Replace(Date() , "-" , "") & Replace(FormatDateTime(now(),4) & Right(now(),2) , ":" , "")
End Function 
%>

