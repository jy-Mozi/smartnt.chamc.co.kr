
//Keyup or Keypress 시 앞자리에 자동으로 0 채우기 
function CharFormat(num,cha) {
	var pow_num = Math.pow(10,cha);
	var rt_value = Number(pow_num) + Number(num);

	num = String(rt_value).substring(1,rt_value.length);
	
	return num;
}

//리턴값을 넘겨주는 식 
//ex) alert(CharNum("433",8)
function CharNum(num,cha) {
	var pow_num = Math.pow(10,cha);
	var rt_value = Number(pow_num) + Number(num);

	CharNum = String(rt_value).substring(1,rt_value.length);
	
	return CharNum;
}

//숫자 체크
function checkNumber(obj)
{    
    alert(event.keyCode);
	if((event.keyCode <= 47 || event.keyCode > 58) && (event.keyCode < 96 || event.keyCode > 105)  && (event.keyCode < 37 || event.keyCode > 40)  && event.keyCode != 8 && event.keyCode != 9 && event.keyCode != 46)
	{
		alert("숫자만 입력하실 수 있습니다");
		event.returnValue = false;
		obj.focus();
	}
}

//숫자 체크 (-50000)도 가능토록
function checkNumberInt(obj)
{    
//    alert(event.keyCode);
	if((event.keyCode <= 47 || event.keyCode > 58) && (event.keyCode < 96 || event.keyCode > 105)  && (event.keyCode < 37 || event.keyCode > 40)  && event.keyCode != 8 && event.keyCode != 9 && event.keyCode != 46 && event.keyCode != 45)
	{
		alert("숫자만 입력하실 수 있습니다");
		event.returnValue = false;
		obj.focus();
	}
}

//숫자 체크
function checkNum(obj)
{

	if((event.keyCode <= 47 || event.keyCode > 58) && event.keyCode != 13)
	{
		alert("숫자만 입력하실 수 있습니다");
		event.returnValue = false;
		obj.focus();
	}
}

//숫자 및 소수점(.) 까지만 허용 
function checkNumSpot(obj)
{
	 if (event.keyCode < 46 || event.keyCode > 57 || event.keyCode == 47) {
		alert("숫자 및 소수점(.)만 입력하실 수 있습니다.");
		event.returnValue = false;
		obj.focus();
	 }
}

//숫자 체크
function checkNumPhone(obj)
{

	if(((event.keyCode <= 47 || event.keyCode > 58 )  && event.keyCode != 13 && event.keyCode != 45))

	{
		alert("숫자또는 '-' 만 입력이 가능합니다.");
		event.returnValue = false;
		obj.focus();
	}
}



//숫자 검사
function formatCheckNum(num) {

  num = num.toString().replace(/\\|\,/g,'');

  if(isNaN(num)) {
	alert("숫자만 입력하실수 있습니다.");
	num = num.substring(0,num.length-1);
	return (num);
  } 
    return num;
}


//기능 : 입력 포커스 자동 이동(객체, length)
function autoFocus(obj1, obj2, len)
{
	if(obj1.value.length == len)
	{
		obj2.focus();
	}
}


//기능 : 입력 포커스 선택 이동 (이동텝 객체 ) 
function moveFocus(obj) {
    obj.focus();
}


/* 
* 기능설명 : 팝업 
* 파라미터 : URL, 팝업이름, 왼쪽위치, 상단위치, 팝업가로사이즈, 팝업세로사이즈, 툴바, 메뉴바, 상태바, 스크롤바, 리사이징, 로케이션바 
*/


function popup_window_common(url, name, left, top, width, height, toolbar, menubar, statusbar, scrollbar, resizable, location) { 
    
	toolbar_str = toolbar ? 'yes' : 'no';     
	menubar_str = menubar ? 'yes' : 'no';     
	statusbar_str = statusbar ? 'yes' : 'no';     
	scrollbar_str = scrollbar ? 'yes' : 'no';     
	resizable_str = resizable ? 'yes' : 'no';     
	location_str = location ? 'yes' : 'no'; 
    
	var mdWindow = null; 
    
	mdWindow = window.open(url, name, 'left='+left+',top='+top+',width='+width+',height='+height+',toolbar='+toolbar_str 
								+',menubar='+menubar_str+',status='+statusbar_str+',scrollbars='+scrollbar_str 
								+',resizable='+resizable_str+',location='+location_str); 

    
	if(!mdWindow){         
		alert("팝업차단을 해제해주세요.");         
		return;     
	} 
    
	mdWindow.focus(); 
    
	return mdWindow; 
}



//팝업 폼값 겨거 가기 
function popup_window_common_submit (url, name, frmname, left, top, width, height, toolbar, menubar, statusbar, scrollbar, resizable, location) { 
    
	toolbar_str = toolbar ? 'yes' : 'no';     
	menubar_str = menubar ? 'yes' : 'no';     
	statusbar_str = statusbar ? 'yes' : 'no';     
	scrollbar_str = scrollbar ? 'yes' : 'no';     
	resizable_str = resizable ? 'yes' : 'no';     
	location_str = location ? 'yes' : 'no'; 
    
	var mdWindow = null; 
    
	mdWindow = window.open(url, name, 'left='+left+',top='+top+',width='+width+',height='+height+',toolbar='+toolbar_str 
								+',menubar='+menubar_str+',status='+statusbar_str+',scrollbars='+scrollbar_str 
								+',resizable='+resizable_str+',location='+location_str); 
	
	var obj = eval(frmname);

	obj.target=name;
    obj.action=url;
	obj.submit();
}


function chk_datediff(sDay,eDay) {
		sDay = sDay.split("-");
		eDay = eDay.split("-");
		var date1 = new Date(sDay[0],sDay[1],sDay[2]).valueOf();
		var date2 = new Date(eDay[0],eDay[1],eDay[2]).valueOf();

		return date1 - date2;
}


//현재 날짜/시간 YYYY-MM-DD 포멧 함수 
function getDayStamp() {
  var d = new Date();
  var s = 
	leadingZeros(d.getFullYear(), 4) + '-' + 
	leadingZeros(d.getMonth() + 1, 2) + '-' + 
	leadingZeros(d.getDate(), 2); 
  return s;
}


//현재 시간 HH:MM 포멧 함수 
function getTimeStamp() {
  var d = new Date();
  var s = 
	leadingZeros(d.getHours(), 2) + ':' +
	leadingZeros(d.getMinutes(), 2) + ':' + 
	leadingZeros(d.getSeconds(), 2);
	
  return s;
}


//현재 날짜/시간 YYYY-MM-DD HH:MM:SS 포멧 함수 
function getDan_TimeStamp() {
  var d = new Date();

  var s =
	leadingZeros(d.getFullYear(), 4) + '-' +
	leadingZeros(d.getMonth() + 1, 2) + '-' +
	leadingZeros(d.getDate(), 2) + ' ' +
	leadingZeros(d.getHours(), 2) + ':' +
	leadingZeros(d.getMinutes(), 2)	 + ':' +
	leadingZeros(d.getSeconds(), 2);
  return s;
}

//자리수 체크 하여 모자란 부분에 앞에 '0'을 추가 함 
function leadingZeros(n, digits) {
  var zero = '';
  n = n.toString();

  if (n.length < digits) {
	for (i = 0; i < digits - n.length; i++)
	  zero += '0';
  }
  return zero + n;
}


//머니형 10,000
function formatCurrency(num) {
  num = num.toString().replace(/\\|\,/g,'');
  if(isNaN(num))
  num = "0";
  sign = (num == (num = Math.abs(num)));
  num = Math.floor(num*100+0.50000000001);

  num = Math.floor(num/100).toString();

  for (var i = 0; i < Math.floor((num.length-(1+i))/3); i++)
  num = num.substring(0,num.length-(4*i+3))+','+
  num.substring(num.length-(4*i+3));
//  return (((sign)?'':'-') + '\\' + num);
  return (num);
 }


//머니형 10,000
function formatCurrencyWithNumChk(num) {

  num = num.toString().replace(/\\|\,/g,'');

  if(isNaN(num)) {
	alert("숫자만 입력하실수 있습니다.");
	num = num.substring(0,num.length-1);
  }
  
  if(isNaN(num))
  num = "0";
  sign = (num == (num = Math.abs(num)));
  num = Math.floor(num*100+0.50000000001);

  num = Math.floor(num/100).toString();

  for (var i = 0; i < Math.floor((num.length-(1+i))/3); i++)
  num = num.substring(0,num.length-(4*i+3))+','+
  num.substring(num.length-(4*i+3));
//  return (((sign)?'':'-') + '\\' + num);
  return (num);
 }


  //NotformatCurrency -통화단위","를 제거
  function NotformatCurrency(num) {
  num = num.toString().replace(',','');
  return (num);
}

function EscapeURL(s) {
	s = escape(s).split("%20").join("+");

	var next = 0;
	var ind = 0;
	var c;
	
	while ((ind = s.indexOf("%u",next)) >= 0) {
		c = unescape(s.substring(ind,ind+6));
		s = s.substring(0,ind) + c + s.substring(ind+6,s.length);
		
		next = ind + c.length;
		if (next >= s.length) break;
	}
	
	return s;
}

function CheckType(s,spc) {
	var i;
	for(i=0; i<s.length; i++) {
		if (spc.indexOf( s.substring(i, i+1)) < 0) {
			return false;
		}
	}
		
	return true;
}


function StringSize (s) {
	var i;
	var len = 0;
	for ( i=0 ; i<s.length; i++) {
		if ( s.charCodeAt(i) > 255 ) {
			len += 2;
		} else {
			len ++;
		}
	}
	return len;
}

function CheckBAN (s,ban) {
	var i;
	for(i=0; i<s.length; i++) {
		if (ban.indexOf( s.substring(i, i+1)) >= 0) {
			return false;
		}
	}

	return true;
}


//바이트 수 체크 제한
function textAreaLengthCheck(thisObj, maxLengthByte) { 			
	var tempByteLength = 0, cutByteLength = 0;  
  
	for(var i = 0; i < thisObj.value.length; i++) {  			  
		if(escape(thisObj.value.charAt(i)).length > 4) {              
			tempByteLength += 2;         
		} else {            
			tempByteLength++;    
		}		 

		if(tempByteLength < maxLengthByte) {  
			cutByteLength++; 
		}	 
	}  
		
	//$("#nowbyte").text (tempByteLength);

	if(tempByteLength > maxLengthByte) {  
		alert(maxLengthByte + "이상 입력할수 없습니다."); 
		thisObj.value = thisObj.value.substring(0, (cutByteLength%2==1)?cutByteLength:cutByteLength+1);  
	} 
}


//바이트 수 체크 제한
function textAreaLengthCheckOnly(thisObj, maxLengthByte) { 			
	var tempByteLength = 0, cutByteLength = 0;  
  
	for(var i = 0; i < thisObj.value.length; i++) {  			  
		if(escape(thisObj.value.charAt(i)).length > 4) {              
			tempByteLength += 2;         
		} else {            
			tempByteLength++;    
		}		 

		if(tempByteLength < maxLengthByte) {  
			cutByteLength++; 
		}	 
	}  

	if(tempByteLength > maxLengthByte) {  
		alert(maxLengthByte + "자 이상 입력할수 없습니다."); 
		thisObj.value = thisObj.value.substring(0, (cutByteLength%2==1)?cutByteLength:cutByteLength+1);  
	} 
}



//생년월일로 만나이 구하기
function computeAge(birthday) 
{ 

	var bday=parseInt(birthday.substring(6,8)); 
	var bmo=(parseInt(birthday.substring(4,6))-1); 
	var byr=parseInt(birthday.substring(0,4)); 
	//alert(byr + "-" + bmo + "-" + bday); 
	
	var byr; 
	var age; 
	var now = new Date(); 

	tday=now.getDate(); 
	tmo=(now.getMonth()); 
	tyr=(now.getFullYear()); 


	if((tmo > bmo)||(tmo==bmo & tday>=bday)) { 
		age=byr 
	} else{ 
		age=byr+1; 
	} 

	return tyr-age; 
} 


		
 /**
  * Time 스트링을 자바스크립트 Date 객체로 변환
  * parameter time: Time 형식의 String
  */
 function toTimeObject(time) { //parseTime(time)
     var year  = time.substr(0,4);
     var month = time.substr(4,2) - 1; // 1월=0,12월=11
     var day   = time.substr(6,2);
     return new Date(year,month,day);
 }

/*
  공백제거 
*/
//function Trim(str) {
    //str = input.replace(/(^\s*)|(\s*$)/, "");
    //return str;
//}



		//addDate("m",3,"2012-03-05","-")
		//addDate("d",10,"20120305","")
		//addDate("yyyy",1,"20120305","")

		//addDate("m",-3,"2012-03-05","-")
		//addDate("d",-10,"20120305","")
		//addDate("yyyy",-1,"20120305","")
		function addDate(pInterval, pAddVal, pYyyymmdd, pDelimiter)
		{
		 var yyyy;
		 var mm;
		 var dd;
		 var cDate;
		 var oDate;
		 var cYear, cMonth, cDay;
		 
		 if (pDelimiter != "") {
		  pYyyymmdd = pYyyymmdd.replace(eval("/\\" + pDelimiter + "/g"), "");
		 }
		 
		 yyyy = pYyyymmdd.substr(0, 4);
		 mm  = pYyyymmdd.substr(4, 2);
		 dd  = pYyyymmdd.substr(6, 2);
		 
		 if (pInterval == "yyyy") {
		  yyyy = (yyyy * 1) + (pAddVal * 1); 
		 } else if (pInterval == "m") {
		  mm  = (mm * 1) + (pAddVal * 1);
		 } else if (pInterval == "d") {
		  dd  = (dd * 1) + (pAddVal * 1);
		 }
		 
		 cDate = new Date(yyyy, mm - 1, dd) // 12월, 31일을 초과하는 입력값에 대해 자동으로 계산된 날짜가 만들어짐.
		 cYear = cDate.getFullYear();
		 cMonth = cDate.getMonth() + 1;
		 cDay = cDate.getDate();
		 
		 cMonth = cMonth < 10 ? "0" + cMonth : cMonth;
		 cDay = cDay < 10 ? "0" + cDay : cDay;
		 
		 if (pDelimiter != "") {
		  return cYear + pDelimiter + cMonth + pDelimiter + cDay;
		 } else {
		  return cYear + cMonth + cDay;
		 }
		 
		}



//check_fgnno('외국인번호13자리');
//check_juminno('주민번호13자리');
//check_busino('사업자번호10자리');

// 재외국인 번호 체크
function check_fgnno(fgnno) {
        var sum=0;
        var odd=0;
        buf = new Array(13);
        for(i=0; i<13; i++) { buf[i]=parseInt(fgnno.charAt(i)); }
        odd = buf[7]*10 + buf[8];
        if(odd%2 != 0) { return false; }
        if( (buf[11]!=6) && (buf[11]!=7) && (buf[11]!=8) && (buf[11]!=9) ) {
                return false;
        }
        multipliers = [2,3,4,5,6,7,8,9,2,3,4,5];
        for(i=0, sum=0; i<12; i++) { sum += (buf[i] *= multipliers[i]); }
        sum = 11 - (sum%11);
        if(sum >= 10) { sum -= 10; }
        sum += 2;
        if(sum >= 10) { sum -= 10; }
        if(sum != buf[12]) { return false }
        return true;
}
// 주민번호 체크
function check_juminno(juminno) {
        if(juminno=="" || juminno==null || juminno.length!=13) {
                alert("주민등록번호를 적어주세요.");
                return false;
        }
        var jumin1 = juminno.substr(0,6);
        var jumin2 = juminno.substr(6,7);
        var yy         = jumin1.substr(0,2);        // 년도
        var mm = jumin1.substr(2,2);        // 월
        var dd = jumin1.substr(4,2);        // 일
        var genda = jumin2.substr(0,1);        // 성별
        var msg, ss, cc;
        // 숫자가 아닌 것을 입력한 경우
        if (!isNumeric(jumin1)) {
                alert("주민등록번호 앞자리를 숫자로 입력하세요.");
                return false;
        }
        // 길이가 6이 아닌 경우
        if (jumin1.length != 6) {
                alert("주민등록번호 앞자리를 다시 입력하세요.");
                return false;
        }
        // 첫번째 자료에서 연월일(YYMMDD) 형식 중 기본 구성 검사
        if (yy < "00" || yy > "99" ||
                mm < "01" || mm > "12" ||
                dd < "01" || dd > "31") {
                alert("주민등록번호 앞자리를 다시 입력하세요.");
                return false;
        }
        // 숫자가 아닌 것을 입력한 경우
        if (!isNumeric(jumin2)) {
                alert("주민등록번호 뒷자리를 숫자로 입력하세요.");
                return false;
        }
        // 길이가 7이 아닌 경우
        if (jumin2.length != 7) {
                alert("주민등록번호 뒷자리를 다시 입력하세요.");
                return false;
        }
        // 성별부분이 1 ~ 4 가 아닌 경우
        if (genda < "1" || genda > "4") {
                alert("주민등록번호 뒷자리를 다시 입력하세요.");
                return false;
        }
        // 연도 계산 - 1 또는 2: 1900년대, 3 또는 4: 2000년대
        cc = (genda == "1" || genda == "2") ? "19" : "20";
        // 첫번째 자료에서 연월일(YYMMDD) 형식 중 날짜 형식 검사
        if (isYYYYMMDD(parseInt(cc+yy), parseInt(mm), parseInt(dd)) == false) {
                alert("주민등록번호 앞자리를 다시 입력하세요.");
                return false;
        }
        // Check Digit 검사
        if (!isSSN(jumin1, jumin2)) {
                alert("입력한 주민등록번호를 검토한 후, 다시 입력하세요.");
                return false;
        }
        return true;
}
// 사업자등록번호 체크
function check_busino(vencod) {
        var sum = 0;
        var getlist =new Array(10);
        var chkvalue =new Array("1","3","7","1","3","7","1","3","5");
        for(var i=0; i<10; i++) { getlist[i] = vencod.substring(i, i+1); }
        for(var i=0; i<9; i++) { sum += getlist[i]*chkvalue[i]; }
        sum = sum + parseInt((getlist[8]*5)/10);
        sidliy = sum % 10;
        sidchk = 0;
        if(sidliy != 0) { sidchk = 10 - sidliy; }
        else { sidchk = 0; }
        if(sidchk != getlist[9]) { return false; }
        return true;
}

function isYYYYMMDD(y, m, d) {
        switch (m) {
        case 2:        // 2월의 경우
                if (d > 29) return false;
                if (d == 29) {
                        // 2월 29의 경우 당해가 윤년인지를 확인
                        if ((y % 4 != 0) || (y % 100 == 0) && (y % 400 != 0))
                                return false;
                }
                break;
        case 4:        // 작은 달의 경우
        case 6:
        case 9:
        case 11:
                if (d == 31) return false;
        }
        // 큰 달의 경우
        return true;
}
function isNumeric(s) {
        for (i=0; i<s.length; i++) {
                c = s.substr(i, 1);
                if (c < "0" || c > "9") return false;
        }
        return true;
}


function isLeapYear(y) {
        if (y < 100)
        y = y + 1900;
        if ( (y % 4 == 0) && (y % 100 != 0) || (y % 400 == 0) ) {
                return true;
        } else {
                return false;
        }
}
function getNumberOfDate(yy, mm) {
        month = new Array(29,31,28,31,30,31,30,31,31,30,31,30,31);
        if (mm == 2 && isLeapYear(yy)) mm = 0;
        return month[mm];
}
function isSSN(s1, s2) {
        n = 2;
        sum = 0;
        for (i=0; i<s1.length; i++)
                sum += parseInt(s1.substr(i, 1)) * n++;
        for (i=0; i<s2.length-1; i++) {
                sum += parseInt(s2.substr(i, 1)) * n++;
                if (n == 10) n = 2;
        }
        c = 11 - sum % 11;
        if (c == 11) c = 1;
        if (c == 10) c = 0;
        if (c != parseInt(s2.substr(6, 1))) return false;
        else return true;
}


//이미지 팝업에 맞게 띄우기

function imgOpen(imgPath){
 poto = new Image();
 poto.src = (imgPath);
 Controlla(imgPath);
}
 
 
function Controlla(imgPath){
 if((poto.width!=0) && (poto.height!=0)) {
  winopen(imgPath, poto.width, poto.height);
 } else{
  funzione = "Controlla('" + imgPath + "')";
  intervallo = setTimeout(funzione, 20);
 }
}
 
 
function winopen(imgPath, Width, Height) {
 var winHandle = window.open("" ,"windowName","toolbar=no,scrollbars=no,resizable=no, top=200, left=200 ,width=" + Width + ",height=" + Height)
  if(winHandle != null) {
  var htmlString = "<html><head><title>이미지</title></head>" 
   htmlString += "<body leftmargin=0 topmargin=0 marginwidth=0 marginheight=0>"
   htmlString += "<a href=javascript:window.close()><img src=" + imgPath + " border=0 alt=이미지클릭:화면닫기></a>"
   htmlString += "</body></html>"
   winHandle.document.open()
   winHandle.document.write(htmlString)
   winHandle.document.close()
  } 
  if(winHandle != null) winHandle.focus()
   return winHandle
}

//==================================================
//	기능 폼 리셋함 
//==================================================
function formReset(frmname)	{

	var frm = eval(frmname);
	var em = frm.elements;

	for(var i=0; i<em.length; i++) {
		if(em[i].type == 'text') em[i].value = '';
		if(em[i].type == 'checkbox') em[i].checked = false;
		if(em[i].type == 'radio') em[i].checked = false;
		if(em[i].type == 'select-one') em[i].options[0].selected = true;
		if(em[i].type == 'textarea') em[i].value = '';
	}
}


//================================================================
//  기능 : 영문과 숫자만 입력 가능 return false ; return true ; 
//================================================================
function onlyNumEng(obj){

  var str = obj.value;

  for(var i=0;i<str.length;i++){	  
	   if(!((str.charCodeAt(i) >= 48 && str.charCodeAt(i) <= 57) || 
		   (str.charCodeAt(i) >= 65 && str.charCodeAt(i) <= 90) ||
			(str.charCodeAt(i) >= 97 && str.charCodeAt(i) <= 122))){							
			return false;
	   }//if
  }	 
	return true;
}


//===============================================================
// 기능 : 레이어 팝업 띄우기 
//===============================================================
/* 스크롤 위치 */
var getNowScroll = function(){
	var de = document.documentElement;
	var b = document.body;
	var now = {};
	now.X = document.all ? (!de.scrollLeft ? b.scrollLeft : de.scrollLeft) : (window.pageXOffset ? window.pageXOffset : window.scrollX);
	now.Y = document.all ? (!de.scrollTop ? b.scrollTop : de.scrollTop) : (window.pageYOffset ? window.pageYOffset : window.scrollY);
	return now;
}

/* 공통 레이어 띄우기	pUrl : 레이어 파일명, w : 레이어 width, h : 레이어 height */
function _openLayerWithLoading(pUrl, w,  h) {

	nowScroll = getNowScroll();
	_closeLayer(); // 레이어 초기화

	var divTag		= '';
	var backTag		= '';
	backTag		+="<div id='backTag' style='display:none;z-index:1000;background-color:#000;filter:alpha(opacity=50);opacity:0.5;-moz-opacity:0.5;'>";
    backTag		+="</div>";
	divTag		+="<div id='feedLayer' style='display:none;z-index:1003;position:absolute;'><iframe src='' frameborder='0' allowtransparency='true'>";
	divTag		+="</iframe></div>";

	$("body").append(backTag);
	$("#backTag").show();


	
    var layer1 = $('#backTag');
	layer1.css({
		position: 'fixed'
		, top   : '0'
		, left  : '0'
		, width : '100%'
		, height: '100%'
	});
    
    

	$("#imgloading").css({
          position      :   'absolute'
		, top           :   $(window).scrollTop() + ($(window).height() - $("#imgloading").height()) /2
		, left          :   '50%'
	});

	var iHeight = nowScroll.Y + 30; // 스크롤 위치에서 200px 위치


    var layerPopTopValue = $(window).scrollTop() + ($(window).height() - parseInt(h)) /2;
     
    if (parseInt(layerPopTopValue) < 0 ) {
        layerPopTopValue = 50
    }


	$("body").append(divTag);
	$("#feedLayer").fadeIn(300);
	var layer2 = $('#feedLayer');
	var layer2frame = $('#feedLayer iframe');
	layer2frame.attr("src", pUrl);
	layer2.css({
		top   : layerPopTopValue
		, left: '50%'
		, marginLeft : -w/2
	});
	layer2frame.css({
		width : w
		, height: h
	})
}


function _openLayer(pUrl, w,  h) {

            
	nowScroll = getNowScroll();
	_closeLayer(); // 레이어 초기화

	var divTag		= '';
	var backTag		= '';
	backTag		+="<div id='backTag' style='display:none;z-index:1000;background-color:#000;filter:alpha(opacity=80);opacity:0.8;-moz-opacity:0.8;'>";
    backTag		+="</div>";
    //backTag     +="<div id='imgloading' style=width:100px;heigth:100px;top:0px;left:0px;'><img src='/img/ajax-loader1.gif'></div>"; 
	divTag		+="<div id='feedLayer' style='display:none;z-index:1003;position:absolute;'><iframe name='lf' id='lf' src='' frameborder='0' allowtransparency='true'>";
	divTag		+="</iframe></div>";

	$("body").append(backTag);
	$("#backTag").show();

    var top = $(window).scrollTop() + ($(window).height() - $("#imgloading").height()) /2 ;
    var left = $(window).scrollLeft() + ($(window).width() - $("#imgloading").width()) / 2;
	$("#imgloading").css('position', 'absolute');
    $("#imgloading").css('top',top);
    $("#imgloading").css('left',left);

	
    var layer1 = $('#backTag');
	layer1.css({
		position: 'fixed'
		, top   : '0'
		, left  : '0'
		, width : '100%'
		, height: '100%'
	});
    
    

	var iHeight = nowScroll.Y + 30; // 스크롤 위치에서 200px 위치
    
    var layerPopTopValue = $(window).scrollTop() + ($(window).height() - parseInt(h)) /2;
     
    if (parseInt(layerPopTopValue) < 0 ) {
        layerPopTopValue = 50
    }


	$("body").append(divTag);
	$("#feedLayer").show();//fadeIn(200);
    
	var layer2 = $('#feedLayer');
	var layer2frame = $('#feedLayer iframe');
	layer2frame.attr("src", pUrl);
	layer2.css({
		top   : layerPopTopValue
		, left: '50%'
		, marginLeft : -w/2
	});
	layer2frame.css({
		width : w
		, height: h
	})
}




/* 공통 레이어 띄우기	pUrl : 레이어 파일명, w : 레이어 width, h : 레이어 height */
function _openLayerOnlyLoading() {
	nowScroll = getNowScroll();
	_closeLayer(); // 레이어 초기화

	var divTag		= '';
	var backTag		= '';
	backTag		+="<div id='backTag' style='display:none;z-index:1000;background-color:#000;filter:alpha(opacity=50);opacity:0.5;-moz-opacity:0.5;'>";
	backTag     +=" <div id='imgloading' style=width:100px;heigth:100px;top:0px;left:0px;'><img src='/img/ajax-loader1.gif'></div>";    
    backTag		+="</div>";
	divTag		+="<div id='feedLayer' style='display:none;z-index:1003;position:absolute;'><iframe src='' frameborder='0' allowtransparency='true'>";
	divTag		+="</iframe></div>";

	$("body").append(backTag);
	$("#backTag").show(300);


	
    var layer1 = $('#backTag');
	layer1.css({
		position: 'fixed'
		, top   : '0'
		, left  : '0'
		, width : '100%'
		, height: '100%'
	});
    
    

	$("#imgloading").css({
          position      :   'absolute'
		, top           :   $(window).scrollTop() + ($(window).height() - $("#imgloading").height()) /2
		, left          :   '50%'
	});

}


function _closeLayer(){
    $("#feedLayer, #backTag ,#imgloading").remove();
    //$("#imgloading").css("display","none");

}

function _closeLayerReload(gUrl){
	$("#feedLayer, #backTag ").remove();
    $("#imgloading").css("display","none");
	window.location.href = gUrl;
}




function _closeLayerNewPage(pUrl, w, h){
	$("#feedLayer, #backTag ,#imgloading").remove();
    //$("#imgloading").css("display","none");
	_openLayer(pUrl, w, h);
}

function _layerSizeTo(w , h) {    
//    var left = $(window).scrollLeft() + ($(window).width() - parseInt(w)) / 2;
////    alert(w + "_______" + h);
//	$("#feedLayer iframe").css({
//		  top   :  left
//        , width :  w
//		, height: h
//	});

	var layer2 = $('#feedLayer');
	var layer2frame = $('#feedLayer iframe');
//	layer2frame.attr("src", pUrl);
	layer2.css({
		top   : "200"
		, left: '50%'
		, marginLeft : -w/2
	});
	layer2frame.css({
		width : w
		, height: h
	})
}




//전화번호 체크
function phoneNumberChcek(objtel1, objtel2, objtel3, numName, alflag)
{
    var Digit = "0123456789";

    if ((objtel1.value.length < 2 )||(objtel2.value.length < 3)||(objtel3.value.length < 4)) {
        alert(numName + "를 입력해 주세요");
        if(objtel1.value.length < 2 )
            objtel1.focus();
        else
            if(objtel2.value.length < 3)
                objtel2.focus();
            else
                if(objtel3.value.length < 4)
                    objtel3.focus();
        
        
        return false;
    }

    if (Tcheck(objtel1, numName, Digit, 2, 3, true, alflag)) {
        return false ;
    }
    
    if (Tcheck(objtel2, numName, Digit, 3, 4, true, alflag)) {
        return false ;
    }
    
    if (Tcheck(objtel3, numName, Digit, 3, 4, true, alflag)) {
        return false ;
    }

    return true;
}



function Left(str, n){
if (n <= 0)
    return "";
else if (n > String(str).length)
    return str;
else
    return String(str).substring(0,n);
}

function Right(str, n){
    if (n <= 0)
       return "";
    else if (n > String(str).length)
       return str;
    else {
       var iLen = String(str).length;
       return String(str).substring(iLen, iLen - n);
    }
}



//특수문자 입력여부//
function Tcheck(target, cmt, astr, lmin, lmax, focusflag, alertflag) {
    var i
    var t = target.value

    if (astr.length > 1) {
        for (i=0; i<t.length; i++)
            if(astr.indexOf(t.substring(i,i+1))<0) {
                if(alertflag)
                    alert(cmt + '에 허용할 수 없는 문자가 입력되었습니다');
                else
                {
                    if(document.all.alertspan != null)
                        document.all.alertspan.innerHTML = "<font color=blue>" + cmt + "에 허용할 수 없는 문자가 입력되었습니다.</font>";
                }
                
                target.value = "";
                if(focusflag) target.focus();					
                return true;
            }
    }
    return false
}


function mailFormChk() {
    
    var fm = document.dataForm;
    var mailchar="0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ_@.";

    if(fm.email1.value.length == 0){
        alert("메일을 입력해 주세요");
        fm.email1.focus();
        return false;
    }	
    else
    {
        if (Tcheck(fm.email1, '메일', mailchar, fm.email1.value.length, fm.email1.value.length, true, true)) {
            return false ;
        }
        
        
        if(fm.email2.value == "input")
        {
            if (Tcheck(fm.email3, '메일', mailchar, fm.email1.value.length, fm.email1.value.length, true, true)) {
                return false ;
            }				
            fm.email.value = fm.email1.value + "@" + fm.email3.value;
        }
        else
            fm.email.value = fm.email1.value + "@" + fm.email2.value;
        
        if(!isEmail(fm.email.value))
        {
            alert("메일 형식이 맞지 않습니다.");
            fm.email1.focus();
            return false;
        }
        return true;
    }
    
}

	//메일 체크  
	function isEmail(str)
	{
		var supported = 0 ;
			
		if (window.RegExp) {
			var tempStr = "a";
			var tempReg = new RegExp(tempStr);
			if(tempReg.test(tempStr)) 
				supported = 1;
		}
		
		if(!supported)
			return (str.indexof(".") > 2) && (str.indexof("@") > 0);
			
		var r1 = new RegExp("(@.*@)|(\\.\\.)|(^\\.)");
		var r2 = new RegExp("^.+\\@(\\[?)[a-zA-Z0-9\\-\\.]+\\.([a-zA-Z]{2,3}|[0-9]{1,3})(\\]?)$");
		
		return (!r1.test(str) && r2.test(str));
	}



//생년월일로 만나이 구하기
function calculateManAge(birthday) 
{ 

	var bday=parseInt(birthday.substring(6,8)); 
	var bmo=(parseInt(birthday.substring(4,6))-1); 
	var byr=parseInt(birthday.substring(0,4)); 
	//alert(byr + "-" + bmo + "-" + bday); 
	
	var byr; 
	var age; 
	var now = new Date(); 

	tday=now.getDate(); 
	tmo=(now.getMonth()); 
	tyr=(now.getFullYear()); 


	if((tmo > bmo)||(tmo==bmo & tday>=bday)) { 
		age=byr 
	} else{ 
		age=byr+1; 
	} 
	
	return tyr-age; 
} 



function insur_open_window(name, url, left, top, width, height, toolbar, menubar, statusbar, scrollbar, resizable)
{
  toolbar_str = toolbar ? 'yes' : 'no';
  menubar_str = menubar ? 'yes' : 'no';
  statusbar_str = statusbar ? 'yes' : 'no';
  scrollbar_str = scrollbar ? 'yes' : 'no';
  resizable_str = resizable ? 'yes' : 'no';
  window.open(url, name, 'left='+left+',top='+top+',width='+width+',height='+height+',toolbar='+toolbar_str+',menubar='+menubar_str+',status='+statusbar_str+',scrollbars='+scrollbar_str+',resizable='+resizable_str);
}