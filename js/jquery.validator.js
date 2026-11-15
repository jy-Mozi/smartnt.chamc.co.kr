/*
*작 성 자 : @전윤기
*작 성 일 : @2012-05-22
*파일내용 : @리터럴 형식을 이용한 form 유효성 체크
    1) 리턴 값 true|false true : 전송, false : 전송하지 않음
    2) 매개변수 객체 리터럴
    3) required이 false라도 빈값이 아니라면 type형식대로 체크하는걸 기본으로 함
var validator = {
    유효성체크할 엘리먼트 아이디 | checkbox, radio는 임의 기임: {
        type: "text|number|email|checkbox|radio|ssn|compare|regexp", -- 필수: 체크할 유효성검사 타입
        -- text 속성       = 일반적인 문자열
        -- number 속성     = 숫자만 사용
        -- email 속성      = 이메일
        -- checkbox 속성   = checkbox
        -- radio 속성      = radio
        -- ssn 속성        = 주민등록번호
        -- compare 속성    = 값 비교
        -- regexp 속성     = 정규식표현으로 문자열 체크
        -- function 속성   = 별도의 함수 호출

        , required: true|false             -- 필수: true = 필수 입력, false = 비필수이나 값이 있을때는 type속성대로 유효성검사
        , message: "내용을 입력해주세요." -- 필수: 검사시 통과하지 못했을 때 사용자에 보여질 alert 메시지

        -- type: text의 옵션, 옵션전체가 선택 사항이며, 반드시 options리터럴 형식을 가져야 함.
        , options: {
            minlength: 3,   -- 선택사항 : 최소 글자수 Int
            maxlength: 5,   -- 선택사항 : 최대 글자수 Int
            lang: "kor"     -- 선택사항 : 언어체크 현재 kor만 존재함(얼마든 추가할 수 있음)
        }

        -- type: number의 옵션, 옵션전체가 선택 사항이며, 반드시 options리터럴 형식을 가져야 함.
        , options: {
            from: 20,       -- 선택사항 : 시작 값 Int
            to: 100,        -- 선택사항 : 끝 값 Int
            minlength: 3,   -- 선택사항 : 최소 글자수 Int
            maxlength: 5,   -- 선택사항 : 최대 글자수 Int
        }

        -- type: email의 옵션, 하나의 input으로 입력받는다면 options리터럴을 사용하면 안된다.
        -- 만약 이메일을 2개의 input으로 받는다면 options리터럴은 필수가 되어야 함.
        , options: {
            id: "email",    -- 이메일 아이디 부분 엘리먼트 아이디 (@는 프로그램내에서 자동으로 붙는다.)
            domain: "host"  -- 이메일 도메일 부분 엘리먼트 아이디
        }

        -- type: checkbox의 옵션
        , options: {
            group: "habi",  -- 필수 : checkbox의 name
            count: 2        -- 선택사항 : 몇개이상 선택할때 사용
        }

        -- type : radio의 옵션
        , options: {
            group: "agree"  -- 필수 : radio의 name
        }

        -- type : ssn의 옵션
        , options: {
            ssn1: "ssn1",   -- 필수 : 주민번호 앞자리를 입력할 엘리먼트 아이디
            ssn2: "ssn2"    -- 필수 : 주민번호 뒷자리를 입력할 엘리먼트 아이디
        }

        -- type : compare의 옵션
        options: {
            target: "pwd1",      -- 필수 : 비교할 엘리먼트 아이디
            error: "에러입니다" -- 필수 : 값이 다를 경우 보여질 alert메시지
        }

        -- type : regexp의 옵션
        , options: {
            express: {/정규식표현/}, -- 필수 : 정규식표현
            error: "에러입니다."     -- 필수 : 정규식표현에 맞지 않을때 보여질 alert메시지
        }

        -- type : function의 옵션
        , options: {
            fnc: "callfunction"     --필수 : 호출할 함수 이름( 위 함수는 반드시 bool형을 반환해야 한다.)
        }
    }
}

정규식표현 사용하는 예시
    /^[0-9]+$/;                                         // 숫자
    /^[a-zA-Z]+$/;                                      // 알파벳(대문자,소문자)
    /^[A-Z]+$/;                                         // 알파벳(대문자)
    /^[가-힣]*$/;                                       // 한글
    /^[0-9]{2,3}-[0-9]{3,4}-[0-9]{4}$/;                 // 집전화
    /^[0-9]{3}-[0-9]{3,4}-[0-9]{4}$/;                   // 휴대폰
    /^[_a-zA-Z0-9-]+@[._a-zA-Z0-9-]+\.[a-zA-Z]+$/;      // 이메일
    /^[.a-zA-Z0-9-]+.[a-zA-Z]+$/;                       // 도메인
    /^(\d{1,3})\.(\d{1,3})\.(\d{1,3})\.(\d{1,3})$/;     // IP
*/

(function ($) {
    // method 호출
    $.fn.validator = function (literal) {
        if ($.type(literal) === "object") {
            return validateLiteral(literal, $(this));
        } else {
            alert("rule을 정의해주세요.");
            return false;
        }
    };

    // 리터럴 노드 검사
    var validateLiteral = function (literal, frm) {
        for (var id in literal) {
            if (literal.hasOwnProperty(id)) {
                if ($.type(literal[id]) === "object") {
                    if (!validation(id, literal[id], frm)) {
                        return false;
                    }
                } else {
                    alert(id + "의 속성값을 지정해주세요.");
                    return false;
                }
            }
        }
        return true;
    };

    // 분기문
    var validation = function (id, obj, frm) {
        var type = obj.type             // 체크할 속성
        , required = obj.required       // 필수여부
        , message = obj.message         // 메세지
        , options = obj.options         // 옵션
        , element = $("#" + id, frm);   // 유효성 검사할 엘리먼트 id

        if (type === undefined || required === undefined || message === undefined) {
            alert("아이디 " + id + "의 필수 속성이 누락되었습니다.\n필수 : type, required, message");
            return false;
        }

        // type 별로 체크
        switch (type) {
            case "text":
                if (!validateText(frm, element, required, message, options)) return false;
                break;
            case "number":
                if (!validateNumber(frm, element, required, message, options)) return false;
                break;
            case "email":
                if (!validateEmail(frm, element, required, message, options)) return false;
                break;
            case "ssn":
                if (!validateSSN(frm, element, required, message, options)) return false;
                break;
            case "compare":
                if (!validateCompare(frm, element, required, message, options)) return false;
                break;
            case "checkbox":
                if (!validateCheckbox(frm, element, required, message, options)) return false;
                break;
            case "radio":
                if (!validateRadio(frm, element, required, message, options)) return false;
                break;
            case "regexp":
                if (!validateRegular(frm, element, required, message, options)) return false;
                break;
            case "function":
                if (!validateFunction(frm, element, required, options)) return false;
                break;
            default:
                alert("허용하지 되지 않은 type입니다.");
                return false;
                break;
        }
        return true;
    };

    // 자바스트립트 함수 호출
    var validateFunction = function (frm, element, required, options) {
        if (!options || !options.fnc) {
            alert("function : 필수속성이 누락되었습니다.\n필수속성 : fnc");
            return false;
        }
        return eval(options.fnc);
    };

    // 주민등록번호 체크
    var validateSSN = function (frm, element, required, message, options) {
        if (!options || !options.ssn1 || !options.ssn2) {
            alert("필수속성이 누락되었습니다.\n필수속성 : ssn1, ssn2");
            return false;
        }

        var ssn1 = $("#" + options.ssn1, frm)
        , ssn2 = $("#" + options.ssn2, frm)
        , pattern = /^([0-9]{6})-?([0-9]{7})$/
        , ssn = ssn1.val() + "" + ssn2.val();

        if ($.trim(ssn1.val()) === "" && required) {
            return validateText(frm, ssn1, required, message, options);
        }

        if ($.trim(ssn2.val()) === "" && required) {
            return validateText(frm, ssn2, required, message, options);
        }

        if (required) {
            if (!validateRegExp(pattern, ssn)) {
                alert("주민등록번호를 확인해주세요.");
                ssn1.focus();
                return false;
            }

            ssn = RegExp.$1 + RegExp.$2;
            var sum = 0;
            var last = ssn.charCodeAt(12) - 0x30;
            var bases = "234567892345";

            for (var i = 0; i < 12; i++) {
                if (isNaN(ssn.substring(i, i + 1))) {
                    alert("주민등록번호를 확인해주세요.");
                    ssn1.focus();
                    return false;
                }
                sum += (ssn.charCodeAt(i) - 0x30) * (bases.charCodeAt(i) - 0x30);
            }

            var mod = sum % 11;
            if ((11 - mod) % 10 != last) {
                alert("주민등록번호를 확인해주세요.");
                ssn1.focus();
                return false;
            }
        } else {
            // 빈값이 아니라면 다시 유효성 검사
            if (ssn1.val() !== "" || ssn2.val() !== "") {
                return validateSSN(frm, element, true, message, options);
            }
        }
        return true;
    };

    // 다른 엘리먼트와 값 비교
    var validateCompare = function (frm, element, required, message, options) {
        if (!options || !options.target || !options.error) {
            alert(element.attr("id") + " : 필수속성이 누락되었습니다.\n필수속성 : target, error");
            return false;
        }

        var target = $("#" + options.target, frm)
            , error = options.error;

        if ($.trim(element.val()) === "" && required) {
            return validateText(frm, element, required, message, options);
        } else {
            if (target) {
                if ($.trim(element.val()) !== $.trim(target.val())) {
                    return validateAlert(frm, element, error);
                }
            }
        }
        return true;
    };

    // 라디오 체크
    var validateRadio = function (frm, element, required, message, options) {
        if (!options || !options.group) {
            alert("필수속성이 누락되었습니다.\n필수속성 : group");
            return false;
        }

        var group = options.group;

        if (required) {
            var checked = $("input[name='" + group + "']:radio:checked", frm).length;
            if (checked === 0) {
                var option = $("input[name='" + group + "']:radio:eq(0)", frm);
                return validateAlert(frm, option, message);
            }
        }
        return true;
    };

    // 체크박스 체크
    var validateCheckbox = function (frm, element, required, message, options) {
        if (!options || !options.group) {
            alert("필수속성이 누락되었습니다.\n필수속성 : group");
            return false;
        }

        var group = options.group
        , count = options.count;

        if (required) {
            var checked = $("input[name=" + group + "]:checkbox:checked", frm).length;
            if (checked === 0) {
                var option = $("input[name='" + group + "']:checkbox:eq(0)", frm);
                return validateAlert(frm, option, message);
            }

            if (count) {
                if (checked < count) {
                    alert(count + "개 이상 선택해주세요.");
                    return false;
                }
            }
        }
        return true;
    };

    // 이메일 체크
    var validateEmail = function (frm, element, required, message, options) {
        var regexp = /^([a-zA-Z0-9_\.\-])+\@(([a-zA-Z0-9\-])+\.)+([a-zA-Z0-9]{2,4})+$/;
        // element값이 하나로만 체크할때
        if (!options) {
            if ($.trim(element.val()) === "" && required) {
                return validateText(frm, element, required, message, options);
            }

            if (required) {
                if (!validateRegExp(regexp, element.val())) {
                    return validateAlert(frm, element, "이메일을 형식을 확인해주세요.");
                }
            } else {
                // 빈값이 아니라면 다시 유효성 검사
                if (element.val() !== "") {
                    return validateEmail(frm, element, true, message, options);
                }
            }
            // 이메일을 element값 2개로 받을때
        } else {
            if (!options.id || !options.domain) {
                alert("필수속성이 누락되었습니다.\n필수속성 : id, domain");
                return false;
            }

            var id = $("#" + options.id, frm)
            , domain = $("#" + options.domain, frm);
            if ($.trim(id.val()) === "" && required) {
                return validateText(frm, id, required, message, options);
            }

            if ($.trim(domain.val()) === "" && required) {
                return validateText(frm, domain, required, message, options);
            }

            if (required) {
                if (!validateRegExp(regexp, id.val() + "@" + domain.val())) {
                    return validateAlert(frm, id, "이메일을 형식을 확인해주세요.");
                }
            } else {
                // 빈값이 아니라면 다시 유효성 검사
                if (id.val() !== "" || domain.val() !== "") {
                    return validateEmail(frm, element, true, message, options);
                }
            }
        }
        return true;
    };

    // 정규식 패턴 체크
    var validateRegular = function (frm, element, required, message, options) {
        if (!options || !options.express || !options.error) {
            alert(element.attr("id") + " : 필수속성이 누락되었습니다.\n필수속성 : express, error");
            return false;
        }

        var express = options.express
        , error = options.error;

        if ($.trim(element.val()) === "" && required) {
            return validateText(frm, element, required, message, options);
        }

        if (required) {
            if (!validateRegExp(express, element.val())) {
                return validateAlert(frm, element, error);
            }
        } else {
            if (element.val() !== "") {
                if (!validateRegExp(express, element.val())) {
                    return validateAlert(frm, element, error);
                }
            }
        }
        return true;
    };

    // 숫자체크
    var validateNumber = function (frm, element, required, message, options) {
        var length = $.trim(element.val()).length;
        if (options) {
            var from = options.from
            , to = options.to
            , minlength = options.minlength
            , maxlength = options.maxlength;
        }
        if ($.trim(element.val()) === "" && required) {
            return validateText(frm, element, required, message, options);
        }

        if (required) {
            var regexp = /^[0-9]+$/;
            if (!validateRegExp(regexp, element.val())) {
                return validateAlert(frm, element, "숫자만 사용해주세요.");
            } else {
                // 최소
                if (from) {
                    if (from > parseInt(element.val(), 10)) {
                        return validateAlert(frm, element, "최소 " + from + " 이상으로 입력해주세요.");
                    }
                }

                // 최대
                if (to) {
                    if (to < parseInt(element.val(), 10)) {
                        return validateAlert(frm, element, "최대 " + to + " 이하로 입력해주세요.");
                    }
                }

                //최소 단어
                if (minlength) {
                    if (minlength > length) {
                        return validateAlert(frm, element, "최소 " + minlength + "자 이상 입력해주세요.");
                    }
                }

                // 최대 단어
                if (maxlength) {
                    if (maxlength < length) {
                        return validateAlert(frm, element, "최대 " + maxlength + "자 이상 입력해주세요.");
                    }
                }
            }
        } else {
            // 빈값이 아니라면 다시 유효성 검사
            if (element.val() !== "") {
                return validateNumber(frm, element, true, message, options);
            }
        }
        return true;
    };

    // 텍스트
    var validateText = function (frm, element, required, message, options) {
        var length = $.trim(element.val()).length;
        if (options) {
            var minlength = options.minlength
            , maxlength = options.maxlength
            , lang = options.lang;
        }

        if ($.trim(element.val()) === "" && required) {
            return validateAlert(frm, element, message);
        }

        //최소 단어
        if (minlength) {
            if (minlength > length) {
                return validateAlert(frm, element, "최소 " + minlength + "자 이상 입력해주세요.");
            }
        }

        // 최대 단어
        if (maxlength) {
            if (maxlength < length) {
                return validateAlert(frm, element, "최대 " + maxlength + "자 이상 입력해주세요.");
            }
        }

        // 언어체크
        if (lang) {
            if (lang === "kor") {
                if (!validateKorean(element.val())) {
                    alert("한글로만 입력해주세요.");
                    element.focus();
                    return false;
                }
            }
        }
        return true;
    };

    // 정규식 비교
    var validateRegExp = function (regexp, value) {
        if (!regexp.test(value)) {
            return false;
        } else {
            return true;
        }
    };

    // 한글 체크
    var validateKorean = function (str) {
        var bool = true;
        for (var i = 0, max = str.length; i < max; i++) {
            if (str.charCodeAt(i) < 0xAC00 || str.charCodeAt(i) > 0xD7A3) {
                if ((str.charCodeAt(i) < 12593 || str.charCodeAt(i) > 12643) && (str.charCodeAt(i) != 32)) {
                    bool = false;
                }
            }
        }
        return bool;
    };

    // 얼럿
    var validateAlert = function (frm, element, message) {
        alert(message);
        if ($.type(element) === "object" ) {
            element.focus();
        }
        return false;
    };
})(jQuery);