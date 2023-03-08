
var is_webtv = UserAgent.indexOf('webtv') != -1;
var is_saf = UserAgent.indexOf('applewebkit') != -1 || navigator.vendor == 'Apple Computer, Inc.';
var is_opera = UserAgent.indexOf('opera') != -1 && opera.version();
var is_moz = (navigator.product == 'Gecko' && !is_saf) && UserAgent.substr(UserAgent.indexOf('firefox') + 8, 3);
var is_ie = (UserAgent.indexOf('msie') != -1 && !is_opera && !is_saf && !is_webtv) && UserAgent.substr(UserAgent.indexOf('msie') + 5, 3);

var controlid = null;
var currdate = null;
var startdate = null;
var enddate  = null;
var yy = null;
var mm = null;
var currday = null;
var addtime = false;
var today = new Date();
var lastcheckedyear = false;
var lastcheckedmonth = false;




function doane(event) {
	e = event ? event : window.event ;
	if(is_ie) {
		e.returnValue = false;
		e.cancelBubble = true;
	} else {
		e.stopPropagation();
		e.preventDefault();
	}
}

function loadcalendar() {
	s = '<style type="text/css">';
	s += '.calendar_header {text-align: center;BORDER: #698CC3 1px solid; BACKGROUND: white; outline: 1px solid #698CC3}';
	s += '.calendar_checked A { COLOR: #FFF; background:#698CC3;}';
	s += '.calendar_expire A:link { COLOR: #999;}';
	s += '.calendar_default A:link { COLOR: #000;}';
	s += '.calendar_today A:link { display: inline-block;COLOR: #9EB6D8; border:1px solid #698CC3; padding:1px;}';
	s += '.calendar_expire A:hover,.calendar_today A:hover,.calendar_default A:hover { COLOR: #FFF; background:#9EB6D8;}';
	s += '#calendar_year,#calendar_month { Z-INDEX: 10; DISPLAY: none; POSITION: absolute;}';
	s += '</style>';

	s += '<div id="calendar" style="display:none; position:absolute; z-index:9;" onclick="doane(event)">';
	s += '<div style="width: 180px;"><table cellspacing="0" cellpadding="0" width="100%" style="text-align: center;BORDER: #698CC3 1px solid; BACKGROUND: white; outline: 1px solid #698CC3">';
	s += '<tr align="center" bgcolor="#E3EFFF"><td height=20><a onclick="refreshcalendar(yy, mm-1)" title="上一月"><<</a></td>';
	s += '<td colspan="5" style="text-align: center"><a onclick="showdiv(\'year\');doane(event)" title="点击选择年份"><span id="year"></span>年</a>&nbsp;&nbsp;';
	s += '<a onclick="showdiv(\'month\');doane(event)" title="点击选择月份"><span id="month"></span>月</a></td><td><A onclick="refreshcalendar(yy, mm+1)" title="下一月">>></A></td></tr>';
	s += '<tr><td height=20>日</td><td>一</td><td>二</td><td>三</td><td>四</td><td>五</td><td>六</td></tr>';
	
	s += '<tr><td colspan="7" height="1"><hr width="100%"  size="1" /></td></tr>';
	for(var i = 0; i < 6; i++) {
		s += '<tr>';
		for(var j = 1; j <= 7; j++)
			s += '<td id=d' + (i * 7 + j) + ' height="19"></td>';
		s += "</tr>";
	}

	s += '</table></div></div>';
	s += '<div id="calendar_year" onclick="doane(event)">';
	s += '<select id=calendar_year_select size=10 onchange="refreshcalendar(this.options[this.selectedIndex].value,mm);$(\'calendar_year\').style.display=\'none\';">'
	for(var k = today.getFullYear()-100; k <= today.getFullYear()+100; k++) {
		s += '<option value='+ k +'>'+ k +'</option>';
	}
	s += '</select>';
	s += '</div>';
	s += '<div id="calendar_month" onclick="doane(event)" style="height:210px;OVERFLOW:auto;padding-left:5px">';
	s += '<select id=calendar_month_select size=10 onchange="refreshcalendar(yy,this.options[this.selectedIndex].value-1);$(\'calendar_month\').style.display=\'none\';">'
	for(var k = 1; k <= 12; k++) {
		s += '<option value=' + (k<10? '0' : '') + k +'>' + (k<10? '0' : '') + k +'</option>';
	}
	s += '</select>';
	s += '</div>';
	if(is_ie && is_ie < 7) {
		s += '<iframe id="calendariframe" frameborder="0" style="display:none;position:absolute;filter:progid:DXImageTransform.Microsoft.Alpha(style=0,opacity=0)"></iframe>';
		s += '<iframe id="calendariframe_year" frameborder="0" style="display:none;position:absolute;filter:progid:DXImageTransform.Microsoft.Alpha(style=0,opacity=0)"></iframe>';
		s += '<iframe id="calendariframe_month" frameborder="0" style="display:none;position:absolute;filter:progid:DXImageTransform.Microsoft.Alpha(style=0,opacity=0)"></iframe>';
	}

	
	document.write(s);
	document.onclick = function(event) {
		$('calendar').style.display = 'none';
		$('calendar_year').style.display = 'none';
		$('calendar_month').style.display = 'none';
		if(is_ie && is_ie < 7) {
			$('calendariframe').style.display = 'none';
			$('calendariframe_year').style.display = 'none';
			$('calendariframe_month').style.display = 'none';
		}
	}
	$('calendar').onclick = function(event) {
		doane(event);
		$('calendar_year').style.display = 'none';
		$('calendar_month').style.display = 'none';
		if(is_ie && is_ie < 7) {
			$('calendariframe_year').style.display = 'none';
			$('calendariframe_month').style.display = 'none';
		}
	}
}

function parsedate(s) {
	/(\d+)\-(\d+)\-(\d+)\s*/.exec(s);
	var m1 = (RegExp.$1 && RegExp.$1 > today.getFullYear()-120 && RegExp.$1 < today.getFullYear()) ? parseFloat(RegExp.$1) : today.getFullYear();
	var m2 = (RegExp.$2 && (RegExp.$2 > 0 && RegExp.$2 < 13)) ? parseFloat(RegExp.$2) : today.getMonth() + 1;
	var m3 = (RegExp.$3 && (RegExp.$3 > 0 && RegExp.$3 < 32)) ? parseFloat(RegExp.$3) : today.getDate();
	/(\d+)\-(\d+)\-(\d+)\s*(\d*):?(\d*)/.exec("0000-00-00");
	return new Date(m1, m2 - 1, m3);
}

function settime(d) {
	yy = yy;
	mm = mm;
	$('calendar').style.display = 'none';
	$('calendar_month').style.display = 'none';
	if(is_ie && is_ie < 7) {
		$('calendariframe').style.display = 'none';
	}
	controlid.value = yy + "-" + zerofill(mm + 1) + "-" + zerofill(d);
}

function showcalendar(event, controlid1, addtime1, startdate1, enddate1) {
	controlid = controlid1;
	addtime = addtime1;
	startdate = startdate1 ? parsedate(startdate1) : false;
	enddate = enddate1 ? parsedate(enddate1) : false;
	currday = controlid.value ? parsedate(controlid.value) : today;
	
	yy = currday.getFullYear();
	mm = currday.getMonth();
	
	$('calendar').style.display = 'block';
	$('calendar').style.left = ie_x(controlid)+'px';
	$('calendar').style.top	= (ie_y(controlid) + 16)+'px';

	doane(event);
	refreshcalendar(currday.getFullYear(), currday.getMonth());
	lastcheckedyear = currday.getFullYear();
	lastcheckedmonth = currday.getMonth() + 1;
	if(is_ie && is_ie < 7) {
		$('calendariframe').style.top = $('calendar').style.top;
		$('calendariframe').style.left = $('calendar').style.left;
		$('calendariframe').style.width = $('calendar').offsetWidth;
		$('calendariframe').style.height = $('calendar').offsetHeight;
		$('calendariframe').style.display = 'block';
	}
}

function refreshcalendar(y, m) {
	var x = new Date(y, m, 1);
	var mv = x.getDay();
	var d = x.getDate();
	var dd = null;
	yy = x.getFullYear();
	mm = x.getMonth();
	$("year").innerHTML = yy;
	$("month").innerHTML = mm + 1 > 9  ? (mm + 1) : '0' + (mm + 1);
	$("calendar_year_select").value = yy;
	$("calendar_month_select").value= mm + 1 > 9  ? (mm + 1) : '0' + (mm + 1);

	for(var i = 1; i <= mv; i++) {
		dd = $("d" + i);
		dd.innerHTML = "&nbsp;";
		dd.className = "";
	}

	while(x.getMonth() == mm) {
		dd = $("d" + (d + mv));
		dd.innerHTML = '<a href="#" onclick="settime(' + d + ');return false">' + d + '</a>';
		if(x.getTime() < today.getTime() || (enddate && x.getTime() > enddate.getTime()) || (startdate && x.getTime() < startdate.getTime())) {
			dd.className = 'calendar_expire';
		} else {
			dd.className = 'calendar_default';
		}
		if(x.getFullYear() == today.getFullYear() && x.getMonth() == today.getMonth() && x.getDate() == today.getDate()) {
			dd.className = 'calendar_today';
			dd.firstChild.title = '今天';
		}
		if(x.getFullYear() == currday.getFullYear() && x.getMonth() == currday.getMonth() && x.getDate() == currday.getDate()) {
			dd.className = 'calendar_checked';
		}
		x.setDate(++d);
	}

	while(d + mv <= 42) {
		dd = $("d" + (d + mv));
		dd.innerHTML = "&nbsp;";
		d++;
	}
}

function showdiv(id) {
	$('calendar_' + id).style.left = ie_x($(id))+'px';
	$('calendar_' + id).style.top = (ie_y($(id)) + 16)+'px';
	$('calendar_' + id).style.display = 'block';
	if(is_ie && is_ie < 7) {
		$('calendariframe_' + id).style.top = $('calendar_' + id).style.top;
		$('calendariframe_' + id).style.left = $('calendar_' + id).style.left;
		$('calendariframe_' + id).style.width = $('calendar_' + id).offsetWidth;
		$('calendariframe_' + id ).style.height = $('calendar_' + id).offsetHeight;
		$('calendariframe_' + id).style.display = 'block';
	}
}

function zerofill(s) {
	var s = parseFloat(s.toString().replace(/(^[\s0]+)|(\s+$)/g, ''));
	s = isNaN(s) ? 0 : s;
	return (s < 10 ? '0' : '') + s.toString();
}

loadcalendar();