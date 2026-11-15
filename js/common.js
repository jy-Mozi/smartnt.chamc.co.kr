//var FRAME_CTI	= window.parent.frames["frame_cti"];
//var FRAME_MAIN	= window.parent.frames["frame_main"];
//document.onkeydown = LockF5; // 새로고침 방지

/*
문자열의 공백제거
사용법 : Trim(obj.userName.value);
*/
function Trim( str ) {
	var resultStr = "";

	resultStr = TrimLeft(str);
	resultStr = TrimRight(resultStr);

	return resultStr;
}

function TrimLeft( str ) {
	var resultStr = "";
	var i = len = 0;
	if (str+"" == "undefined" || str == null) return null;
	str += "";

	if (str.length == 0) resultStr = "";
	else {
		len = str.length;
  		while ((i <= len) && (str.charAt(i) == " "))i++;
  		resultStr = str.substring(i, len);
  	}

  	return resultStr;
}

function TrimRight( str ) {
	var resultStr = "";
	var i = 0;

	if (str+"" == "undefined" || str == null) return null;
	str += "";

	if (str.length == 0) resultStr = "";
	else {
  		i = str.length - 1;
  		while ((i >= 0) && (str.charAt(i) == " "))
 			i--;
	  		resultStr = str.substring(0, i + 1);
	  	}

	  	return resultStr;
}

/*
기능 : 필드에 숫자만 입력할수 있는 함수
사용법 : onkeyup="onlyNumber(this);"
*/
function onlyNumber(el) {
	var oElement = (arguments[0] != null) ? arguments[0] : this;
	var charChk;

	for(var i=0; i<oElement.value.length; i++){
		charChk = oElement.value.charCodeAt(i);

		if(charChk > 57 || charChk < 48){
			alert("공백없이 숫자로만 입력해주세요.");
			oElement.value = oElement.value.substring(0, i);
			oElement.focus();
			return;
		}
	}
}
/*
기능 : 공통 Input Box Blank Check함수
사용법1 : setInputAlert(form1.field, fieldName);
*/
function setInputAlert(field, fieldName) {
	if(Trim(field.value) == "") {
		//alert(fieldName +" 입력해주세요.");
		sAlert("warning",fieldName +" 입력해주세요.");
		field.focus();
		return false;
	}else{
		return true;
	}
}

/*
기능 : 공통 Select Box Blank Check함수
사용법1 : setSelectAlert(form1.field, fieldName);
*/
function setSelectAlert(field, fieldName) {
	//alert(field.value);
	if(Trim(field.value) == "") {
		//alert(fieldName +" 선택해주세요.");
		sAlert("warning",fieldName +" 선택해주세요.");
		field.focus();
		return false;
	}else{
		return true;
	}
}

/*
기능 : 공통 Check Box Blank Check함수
사용법1 : setCheckAlert(form1.field, fieldName);
*/
function setCheckAlert(field, fieldName) {
	if(field == null)
		return false;
	var temp = 0;
	if(field.checked){
		temp++;
	} else {
		for(i = 0; i < field.length; i++){
			if(field[i].checked == true){
				temp = temp + 1;
			}
		}
	}
	if(temp == 0){
		//alert(fieldName+" 선택해주세요.");
		sAlert("warning",fieldName +" 선택해주세요.");
		return false;
	}else{
		return true;
	}
}

/*
설명: 원하는 문자 들 설정하고 문자열을 검사한다.
chars = "ABC"
containsCharsOnly("string",chars)
*/
function containsCharsOnly(input,chars) {
	for(var inx = 0; inx < input.length; inx++) {
       	if (chars.indexOf(input.charAt(inx)) == -1){
           return false;
		}
    }
    return true;
}

/*
설명: 자동으로 탭키효과Auto Tab
사용법 : onKeyUp="return autotab(this, 6, event);"
*/
function autotab(input,len, e) {
	var isNN = (navigator.appName.indexOf("Netscape")!=-1);
	var keycd = (isNN) ? e.which : e.keyCode;
	var filter = (isNN) ? [0,8,9] : [0,8,9,16,17,18,37,38,39,40,46];
	if(input.value.length >= len && !iscontain(filter, keycd))
	{
	  input.value = input.value.slice(0, len);
	  input.form[(getindex(input)+1) % input.form.length].focus();
	}
}

//주민등록번호 검사 입력란 두개짜리
function checkSocialID(no1, no2){
    var pattern = /^([0-9]{6})-?([0-9]{7})$/;
	var num = no1.value +""+ no2.value;
    if (!pattern.test(num)){
		alert("주민등록번호를 확인해주세요.");
		no1.focus();
		return ;
	}

    num = RegExp.$1 + RegExp.$2;

	var sum = 0;
	var last = num.charCodeAt(12) - 0x30;
	var bases = "234567892345";

	for (var i=0; i<12; i++){
		if (isNaN(num.substring(i,i+1))){
			alert("주민등록번호를 확인해주세요.");
			no1.focus();
			return false;
		}

		sum += (num.charCodeAt(i) - 0x30) * (bases.charCodeAt(i) - 0x30);
	}

	var mod = sum % 11;
	if ((11 - mod) % 10 != last){
		alert("주민등록번호를 확인해주세요.");
		no1.focus();
		return false;
	}

	return true;
}

function stringLength(str){
	char_cnt = 0;

	for (var i = 0; i < str.length; i++)	{
		var chr = str.substr(i,1);
		chr = escape(chr);
		key_eg = chr.charAt(1);

		switch(key_eg){
		case "u":
			key_num = chr.substr(2,(chr.length - 1));
			char_cnt += 2;
			break;
		case "B":
			char_cnt += 2;
			break;
		default:
			char_cnt += 1;
		}
	}

	return char_cnt;
}

//이메일 검사
function checkEmail(obj)
{
	var exp = new RegExp ("^[A-Za-z0-9-_\\.]{2,}@[A-Za-z0-9-_\\.]{2,}\\.[A-Za-z0-9-_]{2,}$");

	// 한글이 포함돼 있거나 이메일 형식이 아닌 경우
    alert(obj);
	if (stringLength(obj) == obj.length && exp.test(obj)){
		return true;
	}else{
		alert("이메일 주소를 확인해주세요.");
		return false;
	}
}

//선택된 Value가져오기
function checkValue(checkbox) {
	var countChecked = 0;
	if(checkbox.checked) {
		countChecked = checkbox[i].value;
	} else {
		for( i=0 ; i<checkbox.length ; i++ ) {
			if( checkbox[i].checked ) {
				countChecked = checkbox[i].value;
			}
		}
	}
	return countChecked;
}

//선택된 value값을 구분자로 가져오기
function checkValues(checkbox, mark) {
	var arr = new Array();
	if(checkbox.checked) {
		arr.push(checkbox.value);
	} else {
		for( i=0 ; i<checkbox.length ; i++ ) {
			if( checkbox[i].checked ) {
				arr.push(checkbox[i].value);
			}
		}
	}
	return arr.join(mark);
}

//선택된 객체 수 가져오기
function checkCount(checkbox) {
	var countChecked = 0;
	if(checkbox.checked){
		countChecked++;
	} else {
		for( i=0 ; i<checkbox.length ; i++ ) {
			if( checkbox[i].checked ) {
				countChecked++;
			}
		}
	}
	return countChecked;
}

// Flash 전용 호출함수 id_name : ID또는Name 값, swf_name : 쇼크웨이브파일 경로(이름,GET값포함), wdith: 넓이값, height: 높이값
function swf_func(id_name,swf_name,width,height){
	document.writeln("<object classid='clsid:d27cdb6e-ae6d-11cf-96b8-444553540000' codebase='/asset/flash/swflash.cab#version=8,0,0,0' width='"+width+"' height='"+height+"' id='"+id_name+"' align='middle'>");
	document.writeln("<param name='allowScriptAccess' value='always'>");
	document.writeln("<param name='movie' value='"+swf_name+"'>");
	document.writeln("<param name='quality' value='high'>");
	document.writeln("<param name='bgcolor' value='#ffffff'>");
	document.writeln("<param name='wmode' Value='Transparent'>");
	document.writeln("</object>");
}

//동영상 전용 호출함수
function mov_func(id,name,width,height,showcontrols){
	document.writeln("<OBJECT classid='clsid:22D6F312-B0F6-11D0-94AB-0080C74C7E95' width='"+width+"' height='"+height+"' id='"+id+"'>");
	if (name != "")
		{document.writeln("<PARAM NAME='Filename' VALUE='" +name+ "'>");
			}else{
				document.writeln("<PARAM NAME='TransparentAtStart' VALUE='1'>");
				document.writeln("<PARAM NAME='AnimationAtStart' VALUE='0'>");
	}
	document.writeln("<PARAM NAME='Showcontrols' VALUE='"+showcontrols+"'>");
	document.writeln("<PARAM NAME='Autostart' VALUE='true'>");
	document.writeln("</OBJECT>");
}

//기능 : 팝업창을 화면 중앙에 띄우는 함수
//사용법 : openPopup('auth_write.asp','auth',667,500,'scrollbars=yes,width=667,height=500')
function openPopup( url, name, width, height, features) {
    var left= (screen.width - width) / 2;
    var top = (screen.height - height) / 2;
	var popup = window.open(url, name, features + 'width='+ width +', height='+ height +', left=' + left + ', top=' + top);
	popup.focus();
}


//한글인지 검사
function onlyKorean(str){
	for ( var i=0; i < str.length; i++ ) {
		if ( str.charCodeAt(i) < 0xAC00 || str.charCodeAt(i) > 0xD7A3){
			if (( str.charCodeAt(i) < 12593 || str.charCodeAt(i) > 12643 ) && ( str.charCodeAt(i) != 32)) {
				return true;
			}
		}
	}
	return false;
}

//기 능 : select박스의 option을 제거
function setNull(obj) {
	while (0 < obj.options.length){
		obj.options[0] = null;
	}
}

//기 능 : select박스의 option을 추가한다. JSON이용
function createOptions(sel,_options){
	if(_options == null || _options.length == 0){
		return;
	}

	for(var i = 0;i<_options.length;i++){
		opt = document.createElement("option");
		opt.value = _options[i].rtnVal;
		var op = document.createTextNode(_options[i].rtnName)
		if(_options[i].rtnStyle == '1'){
			opt.style.color = "#F9483A";
		}
		opt.appendChild(op);
		sel.appendChild(opt);
	}
}

//전화번호 형식 체크(숫자-숫자-숫자)
function isValidPhone(input) {
	var format = /^(\d+)-(\d+)-(\d+)$/;
	return isValidFormat(input,format);
}

//패턴과 입력한 형식이 같은지 검사
function isValidFormat(input,format) {
	if (input.value.search(format) != -1) {
		return true; //올바른 포맷 형식
	}
	return false;
}

// 3자리마다 콤마 찍기
function commify(n) {
  var reg = /(^[+-]?\d+)(\d{3})/;   // 정규식
  n += '';                          // 숫자를 문자열로 변환

  while (reg.test(n))
    n = n.replace(reg, '$1' + ',' + '$2');

  return n;
}

// shadowbox
function openShadowbox(content, player, title, width, height){
	Shadowbox.init(
		{overlayOpacity: 0.7, overlayColor: "#000",	players: player, modal: true}
	);
	Shadowbox.open({content : content, player : player, title : title, width : width, height : height});
}

// fancybox
function openFancybox(url, x, y, t){
	$.fancybox.open({
		modal : true
		, type : "iframe"
		, href : url
		, width : x
		, height : y
		, title : t
		, scrolling : "yes"
		, openEffect : "fade"
	});
}

// fancybox close
function closeFancybox(){
	$.fancybox.close();
}

// 새로고침
function LockF5(){
	if (event.keyCode == 116) {
		event.keyCode = 0;
		return false;
	}
}

// 쿠키 셋팅
function setCookie(name, value, expiredays)
{
	//alert(value);
	var todayDate = new Date();
	todayDate.setDate( todayDate.getDate() + expiredays );
	document.cookie = name + "=" + escape( value ) + "; path=/; expires=" + todayDate.toGMTString() + ";"
}

// 쿠키 가져오기
function getCookie(name)
{
    var nameOfCookie = name + "=";
    var x = 0;
    while ( x <= document.cookie.length ){
          var y = (x+nameOfCookie.length);
          if ( document.cookie.substring( x, y ) == nameOfCookie ){
              if ( (endOfCookie=document.cookie.indexOf( ";", y )) == -1 )
                          endOfCookie = document.cookie.length;
              return unescape( document.cookie.substring( y, endOfCookie ) );
          }
          x = document.cookie.indexOf( " ", x ) + 1;
          if ( x == 0 )
            break;
    }
    return "";
}

// Description : 시작일과 종료일 정상여부와 시작일이 작은지 검사
function stratEndCheck(start,end){
	v1 = start.split("-");
	v2 = end.split("-");

	a1 = new Date(v1[0], v1[1], v1[2]).getTime();
	a2 = new Date(v2[0], v2[1], v2[2]).getTime();

	if(a1 < a2 || a2 == ''){
		return true;
	}else{
		return false;
	}
}

// 폰번호 체크
function checkPhoneNumber(obj){
	var isPhoneNumber = /^[0-9*#]{4,12}$/;
	if(!setInputAlert(obj, "전화번호를")) return false;
	if(!isPhoneNumber.test(obj.value)){
		alert("전화번호를 확인해주세요.");
		return false;
	}
	return true;
}

//채크박스 토글 파라미터 : 체크박스 이름 
function checkBoxToggle(obj, name){

	$("input[name="+name+"]").each(function(chk){

		if($(obj).attr('checked')=='checked'){
			this.checked = true;
		}else{
			this.checked = false;
		}

	});

}

function numberPad(n, width) {
    n = n + '';
    return n.length >= width ? n : new Array(width - n.length + 1).join('0') + n;
}



//특수문자 제거 `\'", 제거
function chkChar(obj){
	var RegExp = /['\",]/gi;	//정규식 구문
	if (RegExp.test(obj.value)) {
		// 특수문자 모두 제거
		obj.value = obj.value.replace(RegExp , '');
	}
}


function sAlert(f,m) {
	var tp = f;
	Swal.fire({
		title: m,
		icon: tp,
		imageWidth: 300,
		imageHeight: 200
	});
}


function sAlertMove(f,m,a,p) {
	var tp = f;
	Swal.fire({
		title: m,
		icon: tp,
		imageWidth: 300,
		imageHeight: 200
	}).then((ok)=> {
		if (a == "r") {
			top.location.reload();
		} else if (a == "m") {
			top.location.href = p;
		}
		
	});
}



// dateAdd (dateadd("2025-03-25" , -3 , "d"));
function dateAdd(sDate, v, t) {
	var yy = parseInt(sDate.substr(0, 4), 10);
	var mm = parseInt(sDate.substr(5, 2), 10);
	var dd = parseInt(sDate.substr(8), 10);

	if(t == "d"){
		d = new Date(yy, mm - 1, dd + v);
	}else if(t == "m"){
		d = new Date(yy, mm - 1 + v, dd);
	}else if(t == "y"){
		d = new Date(yy + v, mm - 1, dd);
	}else{
		d = new Date(yy, mm - 1, dd + v);
	}
	yy = d.getFullYear();
	mm = d.getMonth() + 1; mm = (mm < 10) ? '0' + mm : mm;
	dd = d.getDate(); dd = (dd < 10) ? '0' + dd : dd;

	return '' + yy + '-' +  mm  + '-' + dd;
}