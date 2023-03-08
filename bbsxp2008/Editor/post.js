//Emoticon_Part Start
var YuZiCellWidth = 20;
var YuZiCellHeight = 20;
var YuZiRowNum = 4;
var YuZiColNum = 9;
var YuZiCurTab = 0;
var YuZiCurPage = 0;
var YuZiShowWidth = 75;
var YuZiShowHeight = 75;
var YuZiSnapThumb = 0; 				//需要截取缩略图时候，把这个改为1
var YuZiCustomSizeLock = 1; 		//是否开启自定义大小
var IsLoaded = false;

var EmoticonNode,YuZiTabs;

function LoadEmoticonXML(){
	var XmlDomRoot;
	ReadXMLFile("Xml/Emoticons.XML");
	XmlDomRoot = XmlDom.documentElement;
	EmoticonNode = XmlDomRoot.getElementsByTagName('Emoticon');

	YuZiTabs = EmoticonNode.length	//表情——组数
	IsLoaded = true;
}

var gbMoNeedHidden = false;

function Gel(id, obj) {
	obj = (obj == null ? document : obj);
	return obj.getElementById(id);
}

function BbsxpCalcCurPages() {//每组表情页数
	var iDLen = EmoticonNode[YuZiCurTab].getElementsByTagName('ICON').length;		//每组表情数
	
	var iPLen = YuZiRowNum * YuZiColNum;
	var fv = iDLen / iPLen;
	if (iDLen % iPLen == 0) {
		return fv == 0 ? 1 : parseInt(fv);
	}
	else {
		return Math.round(fv + 0.5);
	}
}

function BbsxpJustifyImg(oImg, iWidth, iHeight, bPos) {//单元格表情按比例缩放
	var oNewImg = new Image();
	oNewImg.src = oImg.src;
	
	var w = oNewImg.width;
	var h = oNewImg.height;
	var wb = w / iWidth;
	var hb = h / iHeight;

	if (wb > 1 || hb > 1) {//等比例缩少
		if (wb > hb) {
			oImg.width = iWidth;
			oImg.height = h * iHeight / w;
		}
		else {
			oImg.width = w * iWidth / h;
			oImg.height = iHeight;
		}
		h = oImg.height;
	}
	else {
		oImg.width = w;
		oImg.height = h;
	}

	if (bPos) {
		oImg.style.top = ((iHeight - h) / 2) + "px";
	}

	return;
}

function EmoticonHidePanel() {
	try{
		if (gbMoNeedHidden) {
			Gel("moShowPanel").style.display = "none";
		}
	}catch(e){}
}

function BbsxpEmoticonShow(sUrl, iOffset, oldObj) {
	var obj = Gel("moShowPanel");
	if (sUrl == null || sUrl == "") {
		gbMoNeedHidden = true;
		setTimeout("EmoticonHidePanel()", 300);
	}
	else if (EmoticonCurObj == oldObj){
		gbMoNeedHidden = false;
		var div = Gel("moDivContainer");
		obj.style.top = parseInt(div.offsetTop) + "px";
		if (iOffset < YuZiColNum / 2) {
			obj.style.left = (parseInt(div.offsetLeft) + parseInt(div.offsetWidth) - YuZiShowWidth) + "px";
		}
		else {
			obj.style.left = div.offsetLeft + "px";
		}
		obj.innerHTML = "<img src='" + sUrl + "' width=" + YuZiShowWidth + " height=" + YuZiShowHeight + " onload='BbsxpJustifyImg(this, " + YuZiShowWidth + "," + YuZiShowHeight + ", true);this.style.zIndex = 1;Gel(\"moShowLoadDiv\").style.display = \"none\";' style='position:relative;'></img><div id='moShowLoadDiv' style='text-align:center;position:absolute;left:0;top:0;background:#fff;width:" + YuZiShowWidth + "px;height:" + YuZiShowHeight + "px;'><div style='font:normal 12px Verdana;color:#888;margin-top:" + (YuZiShowHeight/2 - 6) + "px;'>loading...</div></div>";
		obj.style.display = "";
	}
}

var EmoticonCurObj = null;
function BbsxpOnMouseOver(obj, iOffset) {
	if (YuZiSnapThumb) {return ;}
	obj.style.border='1px solid #000080';
	EmoticonCurObj = obj;
	setTimeout(function() {try{BbsxpEmoticonShow(obj.getElementsByTagName("IMG")[0].getAttribute("sval"), iOffset, obj);}catch(e){}}, 300);
}

function BbsxpOnMouseOut(obj) {
	if (YuZiSnapThumb) {return ;}
	obj.style.border='1px solid #F6F6F6';
	EmoticonCurObj = null;
	BbsxpEmoticonShow();
}

function BbsxpGetPageText() {
	return (YuZiCurPage + 1) + "/" + BbsxpCalcCurPages();
}

function BbsxpRefreshData() {
	Gel("moTabContainer").innerHTML = BbsxpTab();
	Gel("moDivContainer").innerHTML = BbsxpTable();
	Gel("moPageText").innerHTML = BbsxpGetPageText();
}

function BbsxpChangeTab(iTab) {
	if (iTab >= YuZiTabs || iTab == YuZiCurTab) {
		return;
	}
	YuZiCurPage = 0;
	YuZiCurTab = iTab;
	BbsxpRefreshData();
}

function BbsxpNextPage() {
	if (YuZiCurPage + 1 >= BbsxpCalcCurPages()) {
		return ;
	}
	YuZiCurPage += 1;
	BbsxpRefreshData();
}

function BbsxpPrevPage() {
	if (YuZiCurPage <= 0) {
		return ;
	}
	YuZiCurPage -= 1;
	BbsxpRefreshData();
}

function BbsxpCell(iCurTab, iCurPage, iNum, iOffset) {
	iNum = iCurPage * YuZiRowNum * YuZiColNum + iNum;
	var data = EmoticonNode[iCurTab].getElementsByTagName('ICON');	//第 iCurTab 组表情数据
	if (iNum >= data.length) {
		return "";
	}
	
	var sSrc = EmoticonNode[iCurTab].getAttributeNode('PathName').nodeValue + data[iNum].getAttributeNode('FileName').nodeValue;		//获取 第 iCurTab 组、第 iNum 个 表情符号的路径
	
	return "<img onclick='BBSXPExecute(\"YuZi_ICON_END\", \"" + sSrc + "\");' sval='" + sSrc + "' src='" + sSrc + "' title='" + GetNodeValue(data[iNum]) + "' width=" + YuZiCellWidth + " height=" + YuZiCellHeight + " onload='BbsxpJustifyImg(this, " + YuZiCellWidth + "," + YuZiCellHeight + ", true);' style='position:relative;cursor:pointer'></img>";
}

function BbsxpTable() {
	var pdivh = "<div style='width:" + YuZiCellWidth + "px;height:" + YuZiCellHeight + "px;'>";
	var code = "<table cellpadding=0 cellspacing=" + (YuZiSnapThumb ? 0 : 1) + " bgcolor=#DFE6F6>";
	for (var i = 0; i < YuZiRowNum; i++) {
		code += "<tr>";
		for (var j = 0; j < YuZiColNum; j++) {
			var cell = BbsxpCell(YuZiCurTab, YuZiCurPage, i * YuZiColNum + j, j);
			code += "<td align=center valign=middle width=" + (YuZiCellWidth + (YuZiSnapThumb ? 0 : 2)) + "px height=" + (YuZiCellHeight + (YuZiSnapThumb ? 0 : 2)) + "px style='background:#f6f6f6;padding:" + (YuZiSnapThumb ? 0 : 1) + "px;border:" + (YuZiSnapThumb ? 0 : 1) + "px solid #F6F6F6;' " + (cell != "" ? "onmouseover='BbsxpOnMouseOver(this, " + j + ");' onmouseout='BbsxpOnMouseOut(this);'>" + pdivh + cell : ">" + pdivh) + "</div></td>";
		}
		code += "</tr>";
	}
	return code + "</table>";
}

function BbsxpTab() {
	if (YuZiCustomSizeLock) {//根据定义大小，重定向:宽,高,行,列
		YuZiCellWidth = parseInt(EmoticonNode[YuZiCurTab].getAttributeNode('Width').nodeValue);
		YuZiCellHeight = parseInt(EmoticonNode[YuZiCurTab].getAttributeNode('Height').nodeValue);
		YuZiRowNum = parseInt(EmoticonNode[YuZiCurTab].getAttributeNode('TableRow').nodeValue);
		YuZiColNum = parseInt(EmoticonNode[YuZiCurTab].getAttributeNode('TableCol').nodeValue);
	}

	var code = "";
	for (var i = 0; i < YuZiTabs; i++) {
		code += "<span style='color:" + (YuZiCurTab == i ? "#000;border:1px solid #DFE6F6;font-weight:bold;border-bottom:1px solid #f6f6f6;border-top:2px solid #FFC83C" : "#000" ) + ";padding:3px 7px 2px 7px;cursor:pointer;' onclick='BbsxpChangeTab(" + i + ")'  unselectable='on'>" + EmoticonNode[i].getAttributeNode('CategoryName').nodeValue + "</span>";
	}
	return code;
}

function BbsxpBtnMouse(obj, down) {
	var s = obj.style;
	s.position = "relative";
	s.top = down ? "1px" : "0px";
	s.left = down ? "1px" : "0px";
}

function BBSXPCube() {
	if(!IsLoaded)LoadEmoticonXML();
	var str;
	str = "<div style='padding:10px 5px 0 5px;' unselectable='on'>";
	str += "<div id='moTabContainer' style='font:normal 12px Verdana;color:#000;padding:2px 6px' unselectable='on'>" + BbsxpTab() + "</div>";
	str += "<div id='moDivContainer' unselectable='on'>" + BbsxpTable() + "</div>";
	
	str += "<div align=right style='text-aling:right;font-size:12px;padding:5px 0 5px 0;color:#000;' unselectable='on'>";
	str += "<span style='float:left'><a href='http://www.biaoqing.com/' target=_blank title='表情网'>更多表情</a></span>"
	str += "<span style='margin:0 10px 0 0;color:#000;font:normal 12px Verdana;' id='moPageText' unselectable='on'>" + BbsxpGetPageText() + "</span>";
	str += "<span style='cursor:pointer;margin:0 2px 0 2px;font-weight:bold;background:#eff7ff;border:1px solid #8ba8c8;color:#000;padding:2px 8px 0 8px' onclick='BbsxpPrevPage();' onmousedown='BbsxpBtnMouse(this,1);' onmouseup='BbsxpBtnMouse(this,0);' title='上一页' unselectable='on'>&lt;</span>";
	str += "<span style='cursor:pointer;margin:0 2px 0 2px;font-weight:bold;background:#eff7ff;border:1px solid #8ba8c8;color:#000;padding:2px 8px 0 8px' onclick='BbsxpNextPage();' onmousedown='BbsxpBtnMouse(this,1);' onmouseup='BbsxpBtnMouse(this,0);' title='下一页' unselectable='on'>&gt;</span>";
	str += "</div>";
	
	str += "<div id='moShowPanel' style='background:#fff;position:absolute;left:0;top:0;border:1px solid #004B97;width:" + YuZiShowWidth + "px;height:" + YuZiShowHeight + "px;text-align:center;display:none' unselectable='on'></div>";
	str += "<div id='moMainLoadDiv' style='font:normal 12px Verdana;display:none;position:absolute;text-valign:center;' unselectable='on'><img src='images/loading.gif' width=16 height=16>loading...</div>";
	str += "</div>";
	
	str += "<img id='moPPL' style='display:none;'/>";
	return str;
}
//Emoticon_Part End

var YuZi_SAFE_MODE = true; 		// true or false
var YuZi_CurrentMode = "DESIGN";
var YuZi_IsChangeMode = false;
var YuZi_TdHeight = -1;
var YuZi_FONT_FAMILY = "Courier New";
var YuZi_WIDTH = "100%";
var YuZi_HEIGHT = (UserAgent.indexOf("chrome")>-1) ? "60%": "100%";	//兼容Chrome浏览器
var YuZi_SKIN_PATH  = "Editor/Images/";
var YuZi_MENU_BORDER_COLOR = '#AAAAAA';
var YuZi_MENU_BG_COLOR = '#EFEFEF';
var YuZi_MENU_TEXT_COLOR = '#222222';
var YuZi_MENU_SELECTED_COLOR = '#CCCCCC';
var YuZi_TOOLBAR_BORDER_COLOR = '#DDDDDD';
var YuZi_TOOLBAR_BG_COLOR = '#EFEFEF';
var YuZi_FORM_BORDER_COLOR = '#DDDDDD';
var YuZi_FORM_BG_COLOR = '#FFFFFF';
var YuZi_BUTTON_COLOR = '#AAAAAA';
var YuZi_ForeColor = '#000000';
var YuZi_BackColor = '#FFFFFF';

var YuZi_LANG = {
	INPUT_URL		: "请输入正确的URL地址。",
	CONFIRM			: "确定",
	CANCEL			: "取消",
	PREVIEW			: "预览",
	REMOTE			: "URL：",
	NEW_WINDOW		: "新窗口",
	CURRENT_WINDOW	: "当前窗口",
	TARGET			: "目标",
	SUBJECT			: "标题",
	CurrentMode		: "编辑器处于代码状态下不能发送内容！",
	SubjectMinLen	: "标题不能小于 2 个字符！",
	SubjectMaxLen	: "标题不能多于 100 个字符！",
	CategoryMaxLen	: "类别不能多于 25 个字符！",
	BodyMinLen		: "内容不能小于 2 个字符！",
	BodyMaxLen		: "内容不能超过 60000 个字符！",
	VerifyCode		: "请输入验证码！",
	DesignMode		: "设计模式",
	CodeMode		: "代码模式",
	AddEditorArea	: "增大编辑区",
	ReduceEditorArea: "减小编辑区",
	Search			: "查　找：",
	Replace			: "替换为：",
	ReplaceButton	: "替换",
	Case			: "大小写：",
	CaseChkBox		: "区分大小写",
	WIDTH			: "宽度",
	HEIGHT			: "高度",
	Transparent		: "透明",
	AutoStart		: "自动播放",
	StatusBar		: "显示状态栏",
	ED2K		: "ED2K链接地址(每行表示一个ED2K链接地址)",
	CopyPaste		: "请用鼠标右击完成 剪切/复制/粘贴 的操作"
}

var YuZi_FONT_NAME = Array(
	Array('宋体', '宋体'),
	Array('黑体', '黑体'),
	Array('楷体_GB2312', '楷体_GB2312'), 
	Array('隶书', '隶书'),
	Array('幼圆', '幼圆'),
	Array('新宋体', '新宋体'),
	Array('Arial', 'Arial'),
	Array('Arial Black', 'Arial Black'),
	Array('Courier New','Courier New'),
	Array('Garamond','Garamond'),
	Array('Georgia','Georgia'),
	Array('Tahoma','Tahoma'),
	Array('Verdana', 'Verdana'),
	Array('Times New Roman', 'Times New Roman'),
	Array('GulimChe', 'GulimChe'),
	Array('MS Gothic', 'MS Gothic')
);
var YuZi_SPECIAL_CHARACTER = Array(
	'§','№','☆','★','○','●','◎','◇','◆','□','℃','‰','■','△','▲','※',
	'→','←','↑','↓','〓','¤','°','＃','＆','＠','＼','︿','＿','￣','―','α',
	'β','γ','δ','ε','ζ','η','θ','ι','κ','λ','μ','ν','ξ','ο','π','ρ',
	'σ','τ','υ','φ','χ','ψ','ω','≈','≡','≠','＝','≤','≥','＜','＞','≮',
	'≯','∷','±','＋','－','×','÷','／','∫','∮','∝','∞','∧','∨','∑','∏',
	'∪','∩','∈','∵','∴','⊥','∥','∠','⌒','⊙','≌','∽','〖','〗',
	'【','】','（','）','［','］'
);

var YuZi_TOP_TOOLBAR_ICON = Array(
	Array('YuZi_FONTNAME', 'font.gif', '字体'),
	Array('YuZi_FONTSIZE', 'fontsize.gif', '文字大小'),
	Array('YuZi_TEXTCOLOR', 'fbcolor.gif', '文字颜色'),
	Array('YuZi_BOLD', 'bold.gif', '粗体'),
	Array('YuZi_ITALIC', 'italic.gif', '斜体'),
	Array('YuZi_UNDERLINE', 'underline.gif', '下划线'),
	Array('YuZi_STRIKE', 'strikethrough.gif', '删除线'),
	Array('YuZi_REMOVE', 'RemoveFormat.gif', '清除格式'),
	Array('YuZi_JUSTIFYLEFT', 'justifyleft.gif', '左对齐'),
	Array('YuZi_JUSTIFYCENTER', 'justifycenter.gif', '居中'),
	Array('YuZi_JUSTIFYRIGHT', 'justifyright.gif', '右对齐'),
	Array('YuZi_JUSTIFYFULL', 'justifyfull.gif', '两端对齐'),
	Array('YuZi_NUMBEREDLIST', 'numberedlist.gif', '编号'),
	Array('YuZi_UNORDERLIST', 'unorderedlist.gif', '项目符号'),
	Array('YuZi_SELECTALL', 'selectall.gif', '全选'),
	Array('YuZi_UNSELECT', 'unselect.gif', '取消全选'),
	Array('YuZi_Delete','delete.gif','删除当前选中区域')
);

var YuZi_BOTTOM_TOOLBAR_ICON = Array(
	Array('YuZi_LINK', 'link.gif', '创建超级链接'),
	Array('YuZi_UNLINK', 'unlink.gif', '删除超级链接'),
	Array('YuZi_ED2K', 'ed2k.gif', '创建ED2K链接'),
	Array('YuZi_SPECIALCHAR', 'specialchar.gif', '插入特殊字符'),
	Array('YuZi_ICON', 'emoticons.gif', '插入表情'),
	Array('YuZi_TABLE', 'table.gif', '插入表格'),
	Array('YuZi_IMAGE', 'image.gif', '插入图片'),
	Array('YuZi_Media', 'Media.gif', '插入 FLASH、MediaPlayer、RealPlayer 文件'),
	Array('Break','Break.gif','换行符'),
	Array('YuZi_Replace','replace.gif','替换'),
	Array('YuZi_CUT', 'cut.gif', '剪切'),
	Array('YuZi_COPY', 'copy.gif', '复制'),
	Array('YuZi_PASTE', 'paste.gif', '粘贴'),
	Array('YuZi_ClearUP','cleanup.gif','清洁代码'),
	Array('YuZi_UNDO', 'undo.gif', '撤消'),
	Array('YuZi_REDO', 'redo.gif', '恢复'),
	Array('YuZi_DATE', 'date.gif', '日期'),
	Array('YuZi_TIME', 'time.gif', '时间')
);


var YuZi_FONT_SIZE = Array(
	Array(1,'8pt'), 
	Array(2,'10pt'), 
	Array(3,'12pt'), 
	Array(4,'14pt'), 
	Array(5,'18pt'), 
	Array(6,'24pt'), 
	Array(7,'36pt')
);

var YuZi_COLOR_TABLE = Array(
	"#000000","#993300","#333300","#003300","#003366","#000080","#333399","#333333",
	"#800000","#FF6600","#808000","#008000","#008080","#0000FF","#666699","#808080",
	"#FF0000","#FF9900","#99CC00","#339966","#33CCCC","#3366FF","#800080","#999999",
	"#FF00FF","#FFCC00","#FFFF00","#00FF00","#00FFFF","#00CCFF","#993366","#C0C0C0",
	"#FF99CC","#FFCC99","#FFFF99","#CCFFCC","#CCFFFF","#99CCFF","#CC99FF","#FFFFFF"
);

var YuZi_OBJ_NAME;
var YuZi_SELECTION;
var YuZi_RANGE;
var YuZi_RANGE_TEXT;
var YuZi_EDITFORM_DOCUMENT;
var YuZi_Replace_DOCUMENT;
var YuZi_ED2K_DOCUMENT;
var YuZi_IMAGE_DOCUMENT;
var YuZi_Media_DOCUMENT;
var YuZi_LINK_DOCUMENT;
var YuZi_BROWSER;
var YuZi_TOOLBAR_ICON;

//浏览器类型
function BBSXPGetBrowser()
{
	var browser = '';
	if (UserAgent.indexOf("msie") > -1) {
		var re = new RegExp("msie\\s?([\\d\\.]+)","ig");
		var arr = re.exec(UserAgent);
		if (parseInt(RegExp.$1) >= 5.5) {
			browser = 'IE';
		}
	} else if (UserAgent.indexOf("firefox") > -1) {
		browser = 'FF';
	} else if (UserAgent.indexOf("netscape") > -1) {
		var temp1 = UserAgent.split(' ');
		var temp2 = temp1[temp1.length-1].split('/');
		if (parseInt(temp2[1]) >= 7) {
			browser = 'NS';
		}
	} else if (UserAgent.indexOf("gecko") > -1) {
		browser = 'ML';
	} else if (UserAgent.indexOf("opera") > -1) {
		var temp1 = UserAgent.split(' ');
		var temp2 = temp1[0].split('/');
		if (parseInt(temp2[1]) >= 9) {
			browser = 'OPERA';
		}
	}
	return browser;
}




//BBCode与HTML切换 Start
function isUndefined(variable) {
	return typeof variable == 'undefined' ? true : false;
}
function trim(str) {
	return (str + '').replace(/(\s+)$/g, '').replace(/^\s+/g, '');
}
function fetchCheckbox(cbn) {
	return $(cbn) && $(cbn).checked == true ? 1 : 0;
}
function parseurl(str, mode, parsecode) {
	if(!parsecode) str= str.replace(/\s*\[code\]([\s\S]+?)\[\/code\]\s*/ig, function($1, $2) {return codetag($2);});
	str = str.replace(/([^>=\]"'\/]|^)((((https?|ftp):\/\/)|www\.)([\w\-]+\.)*[\w\-\u4e00-\u9fa5]+\.([\.a-zA-Z0-9]+|\u4E2D\u56FD|\u7F51\u7EDC|\u516C\u53F8)((\?|\/|:)+[\w\.\/=\?%\-&~`@':+!]*)+\.(jpg|gif|png|bmp))/ig, mode == 'html' ? '$1<img src="$2" border="0">' : '$1[img]$2[/img]');
	str = str.replace(/([^>=\]"'\/@]|^)((((https?|ftp|gopher|news|telnet|rtsp|mms|callto|bctp|ed2k|thunder|synacast):\/\/))([\w\-]+\.)*[:\.@\-\w\u4e00-\u9fa5]+\.([\.a-zA-Z0-9]+|\u4E2D\u56FD|\u7F51\u7EDC|\u516C\u53F8)((\?|\/|:)+[\w\.\/=\?%\-&~`@':+!#]*)*)/ig, mode == 'html' ? '$1<a href="$2" target="_blank">$2</a>' : '$1[url]$2[/url]');
	str = str.replace(/([^\w>=\]"'\/@]|^)((www\.)([\w\-]+\.)*[:\.@\-\w\u4e00-\u9fa5]+\.([\.a-zA-Z0-9]+|\u4E2D\u56FD|\u7F51\u7EDC|\u516C\u53F8)((\?|\/|:)+[\w\.\/=\?%\-&~`@':+!#]*)*)/ig, mode == 'html' ? '$1<a href="$2" target="_blank">$2</a>' : '$1[url]$2[/url]');
	str = str.replace(/([^\w->=\]:"'\.\/]|^)(([\-\.\w]+@[\.\-\w]+(\.\w+)+))/ig, mode == 'html' ? '$1<a href="mailto:$2">$2</a>' : '$1[email]$2[/email]');
	if(!parsecode) {
		for(var i = 0; i <= codecount; i++) {
			str = str.replace("[\tBB_CODE_" + i + "\t]", codehtml[i]);
		}
	}
	return str;
}

var is_mac = UserAgent.indexOf('mac') != -1;
var is_opera = UserAgent.indexOf('opera') != -1 && opera.version();
var re;
if(isUndefined(codecount)) var codecount = '-1';
if(isUndefined(codehtml)) var codehtml = new Array();

function addslashes(str) {
	return preg_replace(['\\\\', '\\\'', '\\\/', '\\\(', '\\\)', '\\\[', '\\\]', '\\\{', '\\\}', '\\\^', '\\\$', '\\\?', '\\\.', '\\\*', '\\\+', '\\\|'], ['\\\\', '\\\'', '\\/', '\\(', '\\)', '\\[', '\\]', '\\{', '\\}', '\\^', '\\$', '\\?', '\\.', '\\*', '\\+', '\\|'], str);
}

function atag(aoptions, text) {
	if(trim(text) == '') {
		return '';
	}

	href = getoptionvalue('href', aoptions);

	if(href == '"'){href ='' ;}

	if(href.substr(0, 11) == 'javascript:') {
		return trim(recursion('a', text, 'atag'));
	} else if(href.substr(0, 7) == 'mailto:') {
		tag = 'email';
		href = href.substr(7);
	} else {
		tag = 'url';
	}

	return '[' + tag + '=' + href + ']' + trim(recursion('a', text, 'atag')) + '[/' + tag + ']';
}

function bbcode2html(str,t) {
	str = trim(str);

	if(str == '') {return '';}
	if (!t){
		str = str.replace(/</g, '&lt;');
		str = str.replace(/>/g, '&gt;');
	}

	if(!fetchCheckbox('DisableBBCode')) {
		str= str.replace(/\[url\]\s*(www.|https?:\/\/|ftp:\/\/|gopher:\/\/|news:\/\/|telnet:\/\/|rtsp:\/\/|mms:\/\/|callto:\/\/|bctp:\/\/|ed2k:\/\/){1}([^\[\"']+?)\s*\[\/url\]/ig, function($1, $2, $3) {return cuturl($2 + $3);});
		str= str.replace(/\[url=www.([^\[\"']+?)\](.+?)\[\/url\]/ig, '<a href="http://www.$1" target="_blank">$2</a>');
		str= str.replace(/\[url=([^\[\"']*?)\](.*?)\[\/url\]/ig, '<a href="$1">$2</a>');
		str= str.replace(/\[url=(https?|ftp|gopher|news|telnet|rtsp|mms|callto|bctp|ed2k){1}:\/\/([^\[\"']+?)\]([\s\S]+?)\[\/url\]/ig, '<a href="$1://$2" target="_blank">$3</a>');
		str= str.replace(/\[(url|ed2k)\]\s*(www.|https?:\/\/|ftp:\/\/|gopher:\/\/|news:\/\/|telnet:\/\/|rtsp:\/\/|mms:\/\/|callto:\/\/|bctp:\/\/){1}([^\[\"']+?)\s*\[\/(url|ed2k)\]/ig, function($2, $3, $4) {return cuturl($2);});
		str= str.replace(/\[(url|ed2k)=www.([^\[\"']+?)\](.+?)\[\/(url|ed2k)\]/ig, '<a href="http://www.$2" target="_blank">$3</a>');
		str= str.replace(/\[(url|ed2k)=(https?|ftp|gopher|news|telnet|rtsp|mms|callto|bctp){1}:\/\/([^\[\"']+?)\]([\s\S]+?)\[\/(url|ed2k)\]/ig, '<a href="$2://$3" target="_blank">$4</a>');
		str= str.replace(/\[email\](.*?)\[\/email\]/ig, '<a href="mailto:$1">$1</a>');
		str= str.replace(/\[strike\](.*?)\[\/strike\]/ig, '<strike>$1</strike>');
		str= str.replace(/\[email=(.[^\[]*)\](.*?)\[\/email\]/ig, '<a href="mailto:$1" target="_blank">$2</a>');
		str = str.replace(/\[color=([^\[\<]+?)\]/ig, '<font color="$1">');
		str = str.replace(/\[size=(\d+?)\]/ig, '<font size="$1">');
		str = str.replace(/\[size=(\d+(\.\d+)?(px|pt|in|cm|mm|pc|em|ex|%)+?)\]/ig, '<font style="font-size: $1">');
		str = str.replace(/\[size=([^\[\<]+?)\]/ig, '<font style="font-size:$1">');
		str = str.replace(/\[font=([^\[\<]+?)\]/ig, '<font face="$1">');
		str = str.replace(/\[align=([^\[\<]+?)\]/ig, '<p align="$1">');
		str = str.replace(/\[float=([^\[\<]+?)\]/ig, '<br style="clear: both"><span style="float: $1;">');

		re = /\[table(?:=(\d{1,4}%?)(?:,([\(\)%,#\w ]+))?)?\]\s*([\s\S]+?)\s*\[\/table\]/ig;
		for (i = 0; i < 4; i++) {
			str = str.replace(re, function($1, $2, $3, $4) {return parsetable($2, $3, $4);});
		}

		str = preg_replace([
			'\\\[\\\/color\\\]', '\\\[\\\/size\\\]', '\\\[\\\/font\\\]', '\\\[\\\/align\\\]', '\\\[b\\\]', '\\\[\\\/b\\\]',
			'\\\[i\\\]', '\\\[\\\/i\\\]', '\\\[u\\\]', '\\\[\\\/u\\\]', '\\\[list\\\]', '\\\[list=1\\\]', '\\\[list=a\\\]',
			'\\\[list=A\\\]', '\\\[\\\*\\\]', '\\\[\\\/list\\\]', '\\\[indent\\\]', '\\\[\\\/indent\\\]', '\\\[\\\/float\\\]'
			], [
			'</font>', '</font>', '</font>', '</p>', '<b>', '</b>', '<i>',
			'</i>', '<u>', '</u>', '<ul>', '<ul type=1>', '<ul type=a>',
			'<ul type=A>', '<li>', '</ul>', '<blockquote>', '</blockquote>', '</span>'
			], str);
	}

	if(!fetchCheckbox('DisableBBCode')) {
		str = str.replace(/\[img\]\s*([^\[\<\r\n]+?)\s*\[\/img\]/ig, '<img src="$1" border="0" alt="" />');
		str = str.replace(/\[img=(\d{1,4})[x|\,](\d{1,4})\]\s*([^\[\<\r\n]+?)\s*\[\/img\]/ig, '<img width="$1" height="$2" src="$3" border="0" alt="" />');
	}
	for(var i = 0; i <= codecount; i++) {
		str = str.replace("[\tBB_CODE_" + i + "\t]", codehtml[i]);
	}
	str = preg_replace(['\t', '   ', '  ','(\r\n|\n|\r)'], ['&nbsp; &nbsp; &nbsp; &nbsp; ', '&nbsp; &nbsp;', '&nbsp;&nbsp;','<br />'], str);
	return str;
}

function cuturl(url) {
	var length = 65;
	var urllink = '<a href="' + (url.toLowerCase().substr(0, 4) == 'www.' ? 'http://' + url : url) + '" target="_blank">';
	if(url.length > length) {
		url = url.substr(0, parseInt(length * 0.5)) + ' ... ' + url.substr(url.length - parseInt(length * 0.3));
	}
	urllink += url + '</a>';
	return urllink;
}

function dpstag(options, text, tagname) {
	if(trim(text) == '') {
		return '\n';
	}
	var pend = parsestyle(options, '', '');
	var prepend = pend['prepend'];
	var append = pend['append'];
	if(InArray(tagname, ['div', 'p','span'])) {
		align = getoptionvalue('align', options);
		if(InArray(align, ['left', 'center', 'right'])) {
			prepend = '[align=' + align + ']' + prepend;
			append += '[/align]';
		} else {
			append += '\n\n';
		}
	}
	return prepend + recursion(tagname, text, 'dpstag') + append;
}

function fetchoptionvalue(option, text) {
	if((position = strpos(text, option)) !== false) {
		delimiter = position + option.length;
		if(text.charAt(delimiter) == '"') {
			delimchar = '"';
		} else if(text.charAt(delimiter) == '\'') {
			delimchar = '\'';
		} else {
			delimchar = ' ';
		}
		delimloc = strpos(text, delimchar, delimiter + 1);
		if(delimloc === false) {
			delimloc = text.length;
		} else if(delimchar == '"' || delimchar == '\'') {
			delimiter++;
		}
		return trim(text.substr(delimiter, delimloc - delimiter));
	} else {
		return '';
	}
}

function fonttag(fontoptions, text) {
	var prepend = '';
	var append = '';
	var tags = new Array();
	tags = {'font' : 'face=', 'size' : 'size=', 'color' : 'color='};
	for(bbcode in tags) {
		optionvalue = fetchoptionvalue(tags[bbcode], fontoptions);
		if(optionvalue) {
			prepend += '[' + bbcode + '=' + optionvalue + ']';
			append = '[/' + bbcode + ']' + append;
		}
	}

	var pend = parsestyle(fontoptions, prepend, append);
	return pend['prepend'] + recursion('font', text, 'fonttag') + pend['append'];
}

function getoptionvalue(option, text) {
	re = new RegExp(option + "(\s+?)?\=(\s+?)?[\"']?(.+?)([\"']|$|>)", "ig");
	var matches = re.exec(text);
	if(matches != null) {
		return trim(matches[3]);
	}
	return '';
}

function html2bbcode(str,t) {
	str = preg_replace(['<style.*?>[\\\s\\\S]*?<\/style>', '<script.*?>[\\\s\\\S]*?<\/script>', '<noscript.*?>[\\\s\\\S]*?<\/noscript>', '<select.*?>[\s\S]*?<\/select>', '<!--[\\\s\\\S]*?-->', ' on[a-zA-Z]{3,16}\\\s?=\\\s?"[\\\s\\\S]*?"'], '', str);
	str= str.replace(/(\r\n|\n|\r)/ig, '');
	str= trim(str.replace(/&((#(32|127|160|173))|shy|nbsp);/ig, ' '));
	str = parseurl(str, 'bbcode', false);
	str = str.replace(/<br\s+?style=(["']?)clear: both;?(\1)[^\>]*>/ig, '');
	str = str.replace(/<br[^\>]*>/ig, "\n");

	if(!fetchCheckbox('DisableBBCode')) {	//是否禁用BBCode
		str = preg_replace(['<table([^>]*(width|background|background-color|bgcolor)[^>]*)>', '<table[^>]*>', '<tr[^>]*(?:background|background-color|bgcolor)[:=]\\\s*(["\']?)([\(\)%,#\\\w]+)(\\1)[^>]*>', '<tr[^>]*>', '<t[dh]([^>]*(width|colspan|rowspan)[^>]*)>', '<t[dh][^>]*>', '<\/t[dh]>', '<\/tr>', '<\/table>'], [function($1, $2) {return tabletag($2);}, '[table]', function($1, $2, $3) {return '[tr=' + $3 + ']';}, '[tr]', function($1, $2) {return tdtag($2);}, '[td]', '[/td]', '[/tr]', '[/table]'], str);

		str = str.replace(/<img([^>]*src[^>]*)>/ig, function($1, $2) {return imgtag($2);});
		str = str.replace(/<a\s+?name=(["']?)(.+?)(\1)[\s\S]*?>([\s\S]*?)<\/a>/ig, '$4');
		str = recursion('b', str, 'simpletag', 'b');
		str = recursion('strong', str, 'simpletag', 'b');
		str = recursion('strike', str, 'simpletag', 'strike');
		str = recursion('i', str, 'simpletag', 'i');
		str = recursion('em', str, 'simpletag', 'i');
		str = recursion('u', str, 'simpletag', 'u');
		str = recursion('a', str, 'atag');
		str = recursion('font', str, 'fonttag');
		str = recursion('blockquote', str, 'simpletag', 'indent');
		str = recursion('ol', str, 'listtag');
		str = recursion('ul', str, 'listtag');
		str = recursion('div', str, 'dpstag');
		str = recursion('p', str, 'dpstag');
		str = recursion('span', str, 'dpstag');
	}
	str = str.replace(/<[\/\!]*?[^<>]*?>/ig, '');	//清除剩下多余的HTML代码

	for(var i = 0; i <= codecount; i++) {
		str = str.replace("[\tBB_CODE_" + i + "\t]", codehtml[i]);
	}

	if (t){		//最后提交
		str = preg_replace(['(\r\n|\n|\r)'], ['<br />'], str);
	}
	else {
		str = preg_replace(['&amp;', '&nbsp;', '&lt;', '&gt;'], ['&', ' ', '<', '>'], str);
	}
	return str;
}

function htmlspecialchars(str) {
	return preg_replace([(is_mac && is_ie ? '&' : '&(?!#[0-9]+;)'), '<', '>', '"'], ['&amp;', '&lt;', '&gt;', '&quot;'], str);
}

function imgtag(attributes) {
	var width = '';
	var height = '';

	re = /src=(["']?)([\s\S]*?)(\1)/i;
	var matches = re.exec(attributes);
	if(matches != null) {
		var src = matches[2];
	} else {
		return '';
	}

	re = /width\s?:\s?(\d{1,4})(px)?/ig;
	var matches = re.exec(attributes);
	if(matches != null) {
		width = matches[1];
	}

	re = /height\s?:\s?(\d{1,4})(px)?/ig;
	var matches = re.exec(attributes);
	if(matches != null) {
		height = matches[1];
	}

	if(!width || !height) {
		re = /width=(["']?)(\d+)(\1)/i;
		var matches = re.exec(attributes);
		if(matches != null) {
			width = matches[2];
		}

		re = /height=(["']?)(\d+)(\1)/i;
		var matches = re.exec(attributes);
		if(matches != null) {
			height = matches[2];
		}
	}

	var imgtag = 'img';

	return width > 0 && height > 0 ?
		'[' + imgtag + '=' + width + ',' + height + ']' + src + '[/' + imgtag + ']' :
		'[img]' + src + '[/img]';
}

function listtag(listoptions, text, tagname) {
	text = text.replace(/<li>(([\s\S](?!<\/li))*?)(?=<\/?ol|<\/?ul|<li|\[list|\[\/list)/ig, '<li>$1</li>') + (is_opera ? '</li>' : '');
	text = recursion('li', text, 'litag');
	var opentag = '[list]';
	var listtype = fetchoptionvalue('type=', listoptions);
	listtype = listtype != '' ? listtype : (tagname == 'ol' ? '1' : '');
	if(InArray(listtype, ['1', 'a', 'A'])) {
		opentag = '[list=' + listtype + ']';
	}
	return text ? opentag + recursion(tagname, text, 'listtag') + '[/list]' : '';
}
function litag(listoptions, text) {
	return '[*]' + text.replace(/(\s+)$/g, '');
}

function parsecode(text) {
	codecount++;
	codehtml[codecount] = '[code]' + htmlspecialchars(text) + '[/code]';
	return "[\tBB_CODE_" + codecount + "\t]";
}

function parsestyle(tagoptions, prepend, append) {
	var searchlist = [
		['align', true, 'text-align:\\s*(left|center|right);?', 1],
		['float', true, 'float:\\s*(left|right);?', 1],
		['color', true, '^(?:\\s|)color:\\s*([^;]+);?', 1],
		['font', true, 'font-family:\\s*([^;]+);?', 1],
		['size', true, 'font-size:\\s*([^;]+);?', 1],
		['b', false, 'font-weight:\\s*(bold);?'],
		['i', false, 'font-style:\\s*(italic);?'],
		['u', false, 'text-decoration:\\s*(underline);?']
	];
	var style = getoptionvalue('style', tagoptions);
	re = /^(?:\s|)color:\s*rgb\((\d+),\s*(\d+),\s*(\d+)\)(;?)/ig;
	style = style.replace(re, function($1, $2, $3, $4, $5) {return("color:#" + parseInt($2).toString(16) + parseInt($3).toString(16) + parseInt($4).toString(16) + $5);});
	var len = searchlist.length;
	for(var i = 0; i < len; i++) {
		re = new RegExp(searchlist[i][2], "ig");
		match = re.exec(style);
		if(match != null) {
			opnvalue = match[searchlist[i][3]];
			prepend += '[' + searchlist[i][0] + (searchlist[i][1] == true ? '=' + opnvalue + ']' : ']');
			append = '[/' + searchlist[i][0] + ']' + append;
		}
	}
	return {'prepend' : prepend, 'append' : append};
}

function parsetable(width, bgcolor, str) {

	if(isUndefined(width)) {
		var width = '';
	} else {
		width = width.substr(width.length - 1, width.length) == '%' ? (width.substr(0, width.length - 1) <= 98 ? width : '98%') : (width <= 560 ? width : '98%');
	}

	str = str.replace(/\[tr(?:=([\(\)%,#\w]+))?\]\s*\[td(?:=(\d{1,2}),(\d{1,2})(?:,(\d{1,4}%?))?)?\]/ig, function($1, $2, $3, $4, $5) {
		return '<tr' + ($2 ? ' style="background: ' + $2 + '"' : '') + '><td' + ($3 ? ' colspan="' + $3 + '"' : '') + ($4 ? ' rowspan="' + $4 + '"' : '') + ($5 ? ' width="' + $5 + '"' : '') + '>';
	});
	str = str.replace(/\[\/td\]\s*\[td(?:=(\d{1,2}),(\d{1,2})(?:,(\d{1,4}%?))?)?\]/ig, function($1, $2, $3, $4) {
		return '</td><td' + ($2 ? ' colspan="' + $2 + '"' : '') + ($3 ? ' rowspan="' + $3 + '"' : '') + ($4 ? ' width="' + $4 + '"' : '') + '>';
	});
	str = str.replace(/\[\/td\]\s*\[\/tr\]/ig, '</td></tr>');

	return '<table border=1 ' + (width == '' ? '' : 'width="' + width + '" ') + (isUndefined(bgcolor) ? '' : ' style="background: ' + bgcolor + '"') + '>' + str + '</table>';
}

function preg_replace(search, replace, str) {
	var len = search.length;
	for(var i = 0; i < len; i++) {
		re = new RegExp(search[i], "ig");
		str = str.replace(re, typeof replace == 'string' ? replace : (replace[i] ? replace[i] : replace[0]));
	}
	return str;
}

function recursion(tagname, text, dofunction, extraargs) {
	if(extraargs == null) {
		extraargs = '';
	}
	tagname = tagname.toLowerCase();

	var open_tag = '<' + tagname;
	var open_tag_len = open_tag.length;
	var close_tag = '</' + tagname + '>';
	var close_tag_len = close_tag.length;
	var beginsearchpos = 0;

	do {
		var textlower = text.toLowerCase();
		var tagbegin = textlower.indexOf(open_tag, beginsearchpos);
		if(tagbegin == -1) {
			break;
		}

		var strlen = text.length;

		var inquote = '';
		var found = false;
		var tagnameend = false;
		var optionend = 0;
		var t_char = '';

		for(optionend = tagbegin; optionend <= strlen; optionend++) {
			t_char = text.charAt(optionend);
			if((t_char == '"' || t_char == "'") && inquote == '') {
				inquote = t_char;
			} else if((t_char == '"' || t_char == "'") && inquote == t_char) {
				inquote = '';
			} else if(t_char == '>' && !inquote) {
				found = true;
				break;
			} else if((t_char == '=' || t_char == ' ') && !tagnameend) {
				tagnameend = optionend;
			}
		}

		if(!found) {
			break;
		}
		if(!tagnameend) {
			tagnameend = optionend;
		}

		var offset = optionend - (tagbegin + open_tag_len);
		var tagoptions = text.substr(tagbegin + open_tag_len, offset)
		var acttagname = textlower.substr(tagbegin * 1 + 1, tagnameend - tagbegin - 1);

		if(acttagname != tagname) {
			beginsearchpos = optionend;
			continue;
		}

		var tagend = textlower.indexOf(close_tag, optionend);
		if(tagend == -1) {
			break;
		}

		var nestedopenpos = textlower.indexOf(open_tag, optionend);
		while(nestedopenpos != -1 && tagend != -1) {
			if(nestedopenpos > tagend) {
				break;
			}
			tagend = textlower.indexOf(close_tag, tagend + close_tag_len);
			nestedopenpos = textlower.indexOf(open_tag, nestedopenpos + open_tag_len);
		}

		if(tagend == -1) {
			beginsearchpos = optionend;
			continue;
		}

		var localbegin = optionend + 1;
		var localtext = eval(dofunction)(tagoptions, text.substr(localbegin, tagend - localbegin), tagname, extraargs);

		text = text.substring(0, tagbegin) + localtext + text.substring(tagend + close_tag_len);

		beginsearchpos = tagbegin + localtext.length;

	} while(tagbegin != -1);

	return text;
}

function simpletag(options, text, tagname, parseto) {
	if(trim(text) == '') {
		return '';
	}
	text = recursion(tagname, text, 'simpletag', parseto);
	return '[' + parseto + ']' + text + '[/' + parseto + ']';
}

function strpos(haystack, needle, offset) {
	if(isUndefined(offset)) {
		offset = 0;
	}

	index = haystack.toLowerCase().indexOf(needle.toLowerCase(), offset);

	return index == -1 ? false : index;
}

function tabletag(attributes) {
	var width = '';
	re = /width=(["']?)(\d{1,4}%?)(\1)/i;
	var matches = re.exec(attributes);

	if(matches != null) {
		width = matches[2].substr(matches[2].length - 1, matches[2].length) == '%' ?
			(matches[2].substr(0, matches[2].length - 1) <= 98 ? matches[2] : '98%') :
			(matches[2] <= 560 ? matches[2] : '98%');
	} else {
		re = /width\s?:\s?(\d{1,4})([px|%])/ig;
		var matches = re.exec(attributes);
		if(matches != null) {
			width = matches[2] == '%' ? (matches[1] <= 98 ? matches[1] : '98%') : (matches[1] <= 560 ? matches[1] : '98%');
		}
	}

	var bgcolor = '';
	re = /(?:background|background-color|bgcolor)[:=]\s*(["']?)((rgb\(\d{1,3}%?,\s*\d{1,3}%?,\s*\d{1,3}%?\))|(#[0-9a-fA-F]{3,6})|([a-zA-Z]{1,20}))(\1)/i;
	var matches = re.exec(attributes);
	if(matches != null) {
		bgcolor = matches[2];
		width = width ? width : '98%';
	}

	return bgcolor ? '[table=' + width + ',' + bgcolor + ']' : (width ? '[table=' + width + ']' : '[table]');
}

function tdtag(attributes) {

	var colspan = 1;
	var rowspan = 1;
	var width = '';

	re = /colspan=(["']?)(\d{1,2})(\1)/ig;
	var matches = re.exec(attributes);
	if(matches != null) {
		colspan = matches[2];
	}

	re = /rowspan=(["']?)(\d{1,2})(\1)/ig;
	var matches = re.exec(attributes);
	if(matches != null) {
		rowspan = matches[2];
	}

	re = /width=(["']?)(\d{1,4}%?)(\1)/ig;
	var matches = re.exec(attributes);
	if(matches != null) {
		width = matches[2];
	}

	return InArray(width, ['', '0', '100%']) ?
		(colspan == 1 && rowspan == 1 ? '[td]' : '[td=' + colspan + ',' + rowspan + ']') :
		'[td=' + colspan + ',' + rowspan + ',' + width + ']';
}
//BBCode与HTML切换 End


function BBSXPDisplayMenu(cmd)
{
	BBSXPEditorForm.focus();
	BBSXPSelection();
	var Top = ie_y($(cmd))+25;
	var Left = ie_x($(cmd));

	$('Yuzi_TempDiv').style.top = Top.toString(10) + 'px';
	$('Yuzi_TempDiv').style.left = Left.toString(10) + 'px';
	$('Yuzi_TempDiv').style.display = 'block';
	$('Yuzi_TempDiv').innerHTML = BBSXPRealTimePopup(cmd);
}

function BBSXPDisableMenu()
{
	$('Yuzi_TempDiv').innerHTML = 'Loading...';
	$('Yuzi_TempDiv').style.display = 'none';
}

function BBSXPGetMenuCommonStyle()
{
	var str = 'position:absolute;top:1px;left:1px;font-size:12px;color:'+YuZi_MENU_TEXT_COLOR+
			';background-color:'+YuZi_MENU_BG_COLOR+';border:solid 1px '+YuZi_MENU_BORDER_COLOR+';z-index:1;display:none;';
	return str;
}

function BBSXPCreateColorTable(cmd, eventStr)
{
	var str = '';
	str += '<table cellpadding="0" cellspacing="2" border="0">';
	for (i = 0; i < YuZi_COLOR_TABLE.length; i++) {
		if (i == 0 || (i >= 8 && i%8 == 0)) {
			str += '<tr>';
		}
		str += '<td style="width:12px;height:12px;border:1px solid #AAAAAA;font-size:1px;cursor:pointer;background-color:' +
		YuZi_COLOR_TABLE[i] + ';" onmouseover="this.style.borderColor=\'#000000\';' + ((eventStr) ? eventStr : '') + '" ' +
		'onmouseout="this.style.borderColor=\'#AAAAAA\';" ' + 
		'onclick="BBSXPExecute(\''+cmd+'_END\', \'' + YuZi_COLOR_TABLE[i] + '\');">&nbsp;</td>';
		if (i >= 7 && i%(i-1) == 0) {
			str += '</tr>';
		}
	}
	str += '</table>';
	return str;
}



//弹出菜单
function BBSXPRealTimePopup(cmd){
	switch (cmd)
	{
		case 'YuZi_FONTNAME':
			var str = '';
			for (i = 0; i < YuZi_FONT_NAME.length; i++) {
				str += '<div style="font-family:' + YuZi_FONT_NAME[i][0] + 
				';padding:2px;cursor:pointer;" ' + 
				'onclick="BBSXPExecute(\'YuZi_FONTNAME_END\', \'' + YuZi_FONT_NAME[i][0] + '\');" ' + 
				'onmouseover="this.style.backgroundColor=\''+YuZi_MENU_SELECTED_COLOR+'\';" ' +
				'onmouseout="this.style.backgroundColor=\''+YuZi_MENU_BG_COLOR+'\';">' + 
				YuZi_FONT_NAME[i][1] + '</div>';
			}
			return str;
			break;
		case 'YuZi_FONTSIZE':
			var str = '';
			for (i = 0; i < YuZi_FONT_SIZE.length; i++) {
				str += '<div style="font-size:' + YuZi_FONT_SIZE[i][1] + 
				';padding:2px;cursor:pointer;" ' + 
				'onclick="BBSXPExecute(\'YuZi_FONTSIZE_END\', \'' + YuZi_FONT_SIZE[i][0] + '\');" ' + 
				'onmouseover="this.style.backgroundColor=\''+YuZi_MENU_SELECTED_COLOR+'\';" ' +
				'onmouseout="this.style.backgroundColor=\''+YuZi_MENU_BG_COLOR+'\';">' + 
				YuZi_FONT_SIZE[i][1] + '</div>';
			}
			return str;
			break;
		case 'YuZi_TEXTCOLOR':
			var str = '';
			str = BBSXPCreateColorTable('YuZi_TEXTCOLOR');
			return str;
			break;
		case 'YuZi_SPECIALCHAR':
			var str = '';
			str += '<table cellpadding="0" cellspacing="2">';
			for (i = 0; i < YuZi_SPECIAL_CHARACTER.length; i++) {
				if (i == 0 || (i >= 10 && i%10 == 0)) {
					str += '<tr>';
				}
				str += '<td style="padding:2px;border:1px solid #AAAAAA;cursor:pointer;" ' + 
				'onclick="BBSXPExecute(\'YuZi_SPECIALCHAR_END\', \'' + YuZi_SPECIAL_CHARACTER[i] + '\');" ' +
				'onmouseover="this.style.borderColor=\'#000000\';" ' +
				'onmouseout="this.style.borderColor=\'#AAAAAA\';">' + YuZi_SPECIAL_CHARACTER[i] + '</td>';
				if (i >= 9 && i%(i-1) == 0) {
					str += '</tr>';
				}
			}
			str += '</table>';
			return str;
			break;
		case 'YuZi_TABLE':
			var str = '';
			var num = 10;
			str += '<table cellpadding="0" cellspacing="0">';
			for (i = 1; i <= num; i++) {
				str += '<tr>';
				for (j = 1; j <= num; j++) {
					var value = i.toString(10) + ',' + j.toString(10);
					str += '<td id="BBSXPTableTd' + i.toString(10) + '_' + j.toString(10) + 
					'" style="width:15px;height:15px;background-color:#FFFFFF;border:1px solid #DDDDDD;cursor:pointer;" ' + 
					'onclick="BBSXPExecute(\'YuZi_TABLE_END\', \'' + value + '\');" ' +
					'onmouseover="BBSXPDrawTableSelected(\''+i.toString(10)+'\', \''+j.toString(10)+'\');" ' + 
					'onmouseout="">&nbsp;</td>';
				}
				str += '</tr>';
			}
			str += '<tr><td colspan="10" id="tableLocation" style="text-align:center;height:20px;"></td></tr>';
			str += '</table>';
			return str;
			break;
		case 'YuZi_ICON':
			return BBSXPCube();
			break;
		case 'YuZi_Replace':
			var str = '';
			str += '<iframe name="BBSXP'+cmd+'Iframe" id="BBSXP'+cmd+'Iframe" frameborder="0" style="width:250px;height:120px;padding:0;margin:0;border:0;"></iframe>';
			return str;
			break;
		case 'YuZi_ED2K':
			var str = '';
			str += '<iframe name="BBSXP'+cmd+'Iframe" id="BBSXP'+cmd+'Iframe" frameborder="0" style="width:510px;height:180px;padding:0;margin:0;border:0;"></iframe>';
			return str;
			break;
		case 'YuZi_IMAGE':
			var str = '';
			str += '<iframe name="BBSXP'+cmd+'Iframe" id="BBSXP'+cmd+'Iframe" frameborder="0" style="width:250px;height:240px;padding:0;margin:0;border:0;"></iframe>';
			return str;
			break;
		case 'YuZi_Media':
			var str = '';
			str += '<iframe name="BBSXP'+cmd+'Iframe" id="BBSXP'+cmd+'Iframe" frameborder="0" style="width:250px;height:300px;padding:0;margin:0;border:0;"></iframe>';
			return str;
			break;
		case 'YuZi_LINK':
			var str = '';
			str += '<iframe name="BBSXP'+cmd+'Iframe" id="BBSXP'+cmd+'Iframe" frameborder="0" style="width:250px;height:85px;padding:0;margin:0;border:0;"></iframe>';
			return str;
			break;
		default: 
			break;
	}
}

function BBSXPDrawIframe(cmd)
{
	switch (cmd)
	{
		case 'YuZi_IMAGE':
			YuZi_IMAGE_DOCUMENT = (YuZi_BROWSER == 'IE') ? document.frames('BBSXP'+cmd+'Iframe').document : $('BBSXP'+cmd+'Iframe').contentDocument;
			var str = '';
			str += '<div align="center">' +
				'<form name="uploadForm" style="margin:0;padding:0;" method="post" onsubmit="javascript:if(parent.BBSXPDrawImageEnd()==false){return false;};">' +
				'<input type="hidden" name="fileName" id="fileName" value="" />' + 
				'<table cellpadding="0" cellspacing="0" style="width:100%;font-size:12px;">' + 
				'<tr><td colspan="2"><table border="0" style="margin-bottom:3px;"><tr><td id="imgPreview" style="width:240px;height:180px;border:1px solid #AAAAAA;background-color:#FFFFFF;" align="center" valign="middle">&nbsp;</td></tr></table></td></tr>' +  	
				'<tr><td style="width:40px;padding-left:5px;">';
				
			str += YuZi_LANG['REMOTE'];
			str += '</td><td style="width:210px;padding-bottom:3px;">';

			str += '<input type="text" id="imgLink" value="http://" maxlength="255" style="width:95%;border:1px solid #555555;" />';
			str += '</td></tr><tr><td colspan="2" style="margin:5px;padding-bottom:5px;" align="center">' +
				'<input type="button" name="button" value="'+YuZi_LANG['PREVIEW']+'" onclick="parent.BBSXPImagePreview();" /> ' +
				'<input type="submit" name="button" id="'+cmd+'submitButton" value="'+YuZi_LANG['CONFIRM']+'" /> ' +
				'<input type="button" name="button" value="'+YuZi_LANG['CANCEL']+'" onclick="parent.BBSXPDisableMenu();" /></td></tr>' + 
				'</table></form></div>';
			BBSXPDrawMenuIframe(YuZi_IMAGE_DOCUMENT, str);
			break;
		case 'YuZi_Replace':
			YuZi_Replace_DOCUMENT = YuZi_BROWSER == 'IE' ? document.frames('BBSXP'+cmd+'Iframe').document : $('BBSXP'+cmd+'Iframe').contentDocument;
			var str = '<table border=0 cellpadding=2 style="background:'+YuZi_MENU_BG_COLOR+';width:100%;height:100%;font-size:12px;">';
			str += '<form name=Replace method="Post">';
			str += '<tr><td>'+YuZi_LANG['Search']+'<input type=text name=TextOne value=></td></tr>';
			str += '<tr><td>'+YuZi_LANG['Replace']+'<input type=text name=TextTwo value=></td></tr>';
			str += '<tr><td>　　<input type=checkbox id=Case /> <label for="Case">'+YuZi_LANG['CaseChkBox']+'</label></td></tr>';
			str += '<tr><td align=center><input id="YuZi_ReplacesubmitButton" type="button" value=" '+YuZi_LANG['ReplaceButton']+' " onclick="parent.BBSXPReplace();">　<input type=button value=" '+YuZi_LANG['CANCEL']+' " onclick="parent.BBSXPDisableMenu();"></td></tr>';
			str += '</form></table>';
			BBSXPDrawMenuIframe(YuZi_Replace_DOCUMENT, str);
			break;
		case 'YuZi_ED2K':
			YuZi_ED2K_DOCUMENT = YuZi_BROWSER == 'IE' ? document.frames('BBSXP'+cmd+'Iframe').document : $('BBSXP'+cmd+'Iframe').contentDocument;
			var str = '<table border=0 cellpadding=2 style="background:'+YuZi_MENU_BG_COLOR+';width:100%;height:100%;font-size:12px;">';
			str += '<form name=ED2K method="Post">';
			str += '<tr><td>'+YuZi_LANG['ED2K']+'</td><td align=right><a href="http://www.ed2000.com" title="ED2000资源分享" target="_blank">更多ED2K资源</a></td></tr>';
			str += '<tr><td colspan=2><textarea name=ED2KURL cols=58 rows=5></textarea></td></tr>';
			str += '<tr><td colspan=2 align=center><input id="YuZi_ED2KsubmitButton" type="button" value=" '+YuZi_LANG['CONFIRM']+' " onclick="parent.BBSXPInsertED2K();">　<input type=button value=" '+YuZi_LANG['CANCEL']+' " onclick="parent.BBSXPDisableMenu();"></td></tr>';
			str += '</form></table>';
			BBSXPDrawMenuIframe(YuZi_ED2K_DOCUMENT, str);
			break;
		case 'YuZi_Media':
			YuZi_Media_DOCUMENT = YuZi_BROWSER == 'IE' ? document.frames('BBSXP'+cmd+'Iframe').document : $('BBSXP'+cmd+'Iframe').contentDocument;
			var str = '<table cellpadding="0" cellspacing="0" style="width:100%;font-size:12px;">' + 
			'<tr><td colspan="2"><table border="0"><tr><td id="MediaPreview" style="width:240px;height:180px;border:1px solid #AAAAAA;background-color:#FFFFFF;" align="center" valign="middle">&nbsp;</td></tr></table></td></tr>' +
			'<tr><td style="width:20px;padding:5px;">'+YuZi_LANG['REMOTE']+'</td>' +
			'<td style="width:230px;padding-bottom:5px;padding-top:5px"><input type="text" id="MediaLink" value="http://" style="width:195px;border:1px solid #555555;" /></td></tr>' +
			'<tr>' +
			'<td colspan="2" style="margin:5px;padding-bottom:5px;padding-left:5px">'+YuZi_LANG['WIDTH']+'：<input type=text id="MediaWidth" value=300 size=4 style="border:1px solid #555555;">　　'+YuZi_LANG['HEIGHT']+'：<input type=text id=MediaHeigth value=250 size=4 style="border:1px solid #555555;"></td>' +
			'</tr>' +
			
			'<tr><td colspan="2" style="padding-bottom:5px;padding-left:5px">'+YuZi_LANG['AutoStart']+'：<input type="checkbox" id="AutoStart" value="1" checked />　　　'+YuZi_LANG['StatusBar']+'：<input type="checkbox" value="1" id=ShowStatusBar checked></td></tr>' +
			
			
			'<tr><td colspan="2" style="margin:5px;padding-bottom:5px;" align="center">' +
			'<input type="button" name="button" value="'+YuZi_LANG['PREVIEW']+'" onclick="parent.BBSXPMediaPreview();" /> ' +
			'<input type="submit" name="button" id="'+cmd+'submitButton" value="'+YuZi_LANG['CONFIRM']+'" onclick="parent.BBSXPDrawMediaEnd();" /> ' +
			'<input type="button" name="button" value="'+YuZi_LANG['CANCEL']+'" onclick="parent.BBSXPDisableMenu();" /></td></tr>' + 
			'</table>';
			BBSXPDrawMenuIframe(YuZi_Media_DOCUMENT, str);
			break;
		case 'YuZi_LINK':
			YuZi_LINK_DOCUMENT = YuZi_BROWSER == 'IE' ? document.frames('BBSXP'+cmd+'Iframe').document : $('BBSXP'+cmd+'Iframe').contentDocument;
			var str = '';
			str += '<table cellpadding="0" cellspacing="0" style="width:100%;font-size:12px;">' + 
				'<tr><td style="width:50px;padding:5px;">URL</td>' +
				'<td style="width:200px;padding-top:5px;padding-bottom:5px;"><input type="text" id="hyperLink" value="http://" style="width:190px;border:1px solid #555555;background-color:#FFFFFF;"></td>' +
				'<tr><td style="padding:5px;">'+YuZi_LANG['TARGET']+'</td>' +
				'<td style="padding-bottom:5px;"><select id="hyperLinkTarget"><option value="_blank" selected="selected">'+YuZi_LANG['NEW_WINDOW']+'</option><option value="">'+YuZi_LANG['CURRENT_WINDOW']+'</option></select></td></tr>' + 
				'<tr><td colspan="2" style="padding-bottom:5px;" align="center">' +
				'<input type="submit" name="button" id="'+cmd+'submitButton" value="'+YuZi_LANG['CONFIRM']+'" onclick="parent.BBSXPDrawLinkEnd();" /> ' +
				'<input type="button" name="button" value="'+YuZi_LANG['CANCEL']+'" onclick="parent.BBSXPDisableMenu();" /></td></tr>';
			str += '</table>';
			BBSXPDrawMenuIframe(YuZi_LINK_DOCUMENT, str);
			break;
		default:
			break;
	}
}
function BBSXPDrawMenuIframe(obj, str)
{
	obj.open();
	obj.write(str);
	obj.close();
	obj.body.style.color = YuZi_MENU_TEXT_COLOR;
	obj.body.style.backgroundColor = YuZi_MENU_BG_COLOR;
	obj.body.style.margin = 0;
	obj.body.scroll = 'no';
}
function BBSXPDrawTableSelected(i, j)
{
	var text = i.toString(10) + ' 乘 ' + j.toString(10) + ' 表格';
	$('tableLocation').innerHTML = text;
	var num = 10;
	for (m = 1; m <= num; m++) {
		for (n = 1; n <= num; n++) {
			var obj = $('BBSXPTableTd' + m.toString(10) + '_' + n.toString(10) + '');
			if (m <= i && n <= j) {
				obj.style.backgroundColor = YuZi_MENU_SELECTED_COLOR;
			} else {
				obj.style.backgroundColor = '#FFFFFF';
			}
		}
	}
}
//ED2K_开始
function BBSXPInsertED2K()
{
	var Ed2kUrl = YuZi_ED2K_DOCUMENT.ED2K.elements['ED2KURL'].value;
	var re = /^ed2k:\/\/\|file\|[^\\\/:*?<>""|]+[\.]?[^\\\/:*?<>""|]+\|\d+\|[0-9a-zA-Z]{32}\|/;
	if (Ed2kUrl != null){
		var Ed2kUrlArray = new Array();
		Ed2kUrlArray=Ed2kUrl.split('\n');
		Ed2kUrl = '';
		for(var i=0;i<Ed2kUrlArray.length;i++){
			if (Ed2kUrlArray[i]!='' && re.test(Ed2kUrlArray[i])){
				if (Ed2kUrl == '')Ed2kUrl = '[ed2k]';
				Ed2kUrl += '[Link]'+Ed2kUrlArray[i].replace(/(^\s*)|(\s*$)/g,"")+'[/Link]';
			}
		}
		if (Ed2kUrl != '')Ed2kUrl += '[/ed2k]';
		var element = document.createElement("span");
		element.appendChild(document.createTextNode(Ed2kUrl));
		BBSXPInsertItem(element);
	}
	else {
		BBSXPEditorForm.focus();
	}
	BBSXPDisableMenu();
}

//替换内容_开始
function BBSXPReplace()
{
	var TextOne = YuZi_Replace_DOCUMENT.Replace.elements['TextOne'].value
	var TextTwo = YuZi_Replace_DOCUMENT.Replace.elements['TextTwo'].value
	var IgnoreCase = YuZi_Replace_DOCUMENT.Replace.elements['Case'].value
	if (TextOne != null && TextTwo != null){
		con = YuZi_EDITFORM_DOCUMENT.body.innerHTML;
		if (IgnoreCase.toLowerCase() == "on"){
			con = BBSXPRegEx(con,TextOne,TextTwo,true);
		}
		else{
			con = BBSXPRegEx(con,a,b);
		}
		YuZi_EDITFORM_DOCUMENT.body.innerHTML = con;
	}
	else
		BBSXPEditorForm.focus();
	
	BBSXPDisableMenu();
}
function BBSXPRegEx(s,a,b,i){
	a = a.replace("?","\\?");
	if (i==null){
		var r = new RegExp(a,"gi");
	}
	else if (i) {
		var r = new RegExp(a,"g");
	}
	else{
		var r = new RegExp(a,"gi");
	}
	return s.replace(r,b); 
}
//替换内容_结束

//插入图片 预览
function BBSXPImagePreview()
{
	var url = YuZi_IMAGE_DOCUMENT.getElementById('imgLink').value;
	if (url=="http://" || !url || url.length<10) {
		alert(YuZi_LANG['INPUT_URL']);
		return false;
	}
	var imgObj = YuZi_IMAGE_DOCUMENT.createElement("IMG");
	imgObj.src = url;
	var width = parseInt(imgObj.width);
	var height = parseInt(imgObj.height);
	var rate = parseInt(width/height);
	if (width >240 && height <= 180) {
		width = 240;
		height = parseInt(width/rate);
	} else if (width <=240 && height > 180) {
		height = 180;
		width = parseInt(height*rate);
	} else if (width >240 && height > 180) {
		height = 180;
		width = parseInt(height*rate);
		if (width >240) width = 240;
	}
	imgObj.style.width = width;
	imgObj.style.height = height;
	var el = YuZi_IMAGE_DOCUMENT.getElementById('imgPreview');
	if (el.hasChildNodes()) {
		el.removeChild(el.childNodes[0]);
	}
	el.appendChild(imgObj);
	return imgObj;
}
//插入图片
function BBSXPDrawImageEnd()
{
	var url = YuZi_IMAGE_DOCUMENT.getElementById('imgLink').value;
	if (url=="http://" || !url || url.length<10) {
		alert(YuZi_LANG['INPUT_URL']);
		return false;
	}
		
	BBSXPEditorForm.focus();
	var element = document.createElement("img");
	element.src = url;
	element.border = 0;
	BBSXPSelect();
	BBSXPInsertItem(element);
	BBSXPDisableMenu();
}
function BBSXPGetMediaHtmlTag(url)
{
	var str = '<embed src="'+url+'" quality="high" AutoStart="true"></embed>';
	return str;
}
//插入Media 预览
function BBSXPMediaPreview()
{
	var url = YuZi_Media_DOCUMENT.getElementById('MediaLink').value;
	if (url=="http://" || !url || url.length<10) {
		alert(YuZi_LANG['INPUT_URL']);
		return false;
	}
	var el = YuZi_Media_DOCUMENT.getElementById('MediaPreview');
	el.innerHTML = BBSXPGetMediaHtmlTag(url);
}
//插入Media
function BBSXPDrawMediaEnd()
{
	var url = YuZi_Media_DOCUMENT.getElementById('MediaLink').value;
	var Width = YuZi_Media_DOCUMENT.getElementById('MediaWidth').value;
	var Height = YuZi_Media_DOCUMENT.getElementById('MediaHeigth').value;
	var autostart = YuZi_Media_DOCUMENT.getElementById('AutoStart').checked;
	var ShowStatusBar = YuZi_Media_DOCUMENT.getElementById('ShowStatusBar').checked;
	if (url=="http://" || !url || url.length<10) {
		alert(YuZi_LANG['INPUT_URL']);
		return false;
	}

	var EmbedStr = '[media='+Width+','+Height+','+autostart+','+ShowStatusBar+']'+url+'[/media]';
	BBSXPEditorForm.focus();
	BBSXPSelect();
	var element = document.createElement("span");
	element.appendChild(document.createTextNode(EmbedStr));
	BBSXPInsertItem(element);
	BBSXPDisableMenu();
}

//插入链接
function BBSXPDrawLinkEnd()
{
	var range;
	var url = YuZi_LINK_DOCUMENT.getElementById('hyperLink').value;
	var target = YuZi_LINK_DOCUMENT.getElementById('hyperLinkTarget').value;
	if (url.match(/http|ftp|https|mailto|wais|telnet|news|gopher|mms:\/\/.{3,}/) == null) {
		alert(YuZi_LANG['INPUT_URL']);
		return false;
	}
	BBSXPEditorForm.focus();
	BBSXPSelect();
	var element;
    if (YuZi_BROWSER != 'IE') {
		BBSXPExecuteValue('CreateLink', url);
		element = YuZi_RANGE.startContainer.previousSibling;
		element.target = target;
		if (target) {
			element.target = target;
		}
    }
	BBSXPDisableMenu();
}

function BBSXPSelection()
{
	if (YuZi_BROWSER == 'IE') {
		YuZi_SELECTION = YuZi_EDITFORM_DOCUMENT.selection;
		YuZi_RANGE = YuZi_SELECTION.createRange();
		YuZi_RANGE_TEXT = YuZi_RANGE.text;
	}
	else {
		YuZi_SELECTION = $("BBSXPEditorForm").contentWindow.getSelection();
        YuZi_RANGE = YuZi_SELECTION.getRangeAt(0);
		YuZi_RANGE_TEXT = YuZi_RANGE.toString();
	}
}
function BBSXPSelect()
{
	if (YuZi_BROWSER == 'IE') {
		YuZi_RANGE.select();
	}
}

function BBSXPInsertItem(insertNode)
{
	if (YuZi_BROWSER == 'IE') {
		if (YuZi_SELECTION.type.toLowerCase() == 'control') {
			YuZi_RANGE.item(0).outerHTML = insertNode.outerHTML;
		}
		else {
			YuZi_RANGE.pasteHTML(insertNode.outerHTML);
		}
	}
	else {
        YuZi_SELECTION.removeAllRanges();
		YuZi_RANGE.deleteContents();
        var startRangeNode = YuZi_RANGE.startContainer;
        var startRangeOffset = YuZi_RANGE.startOffset;
        var newRange = document.createRange();
		if (startRangeNode.nodeType == 3 && insertNode.nodeType == 3) {
            startRangeNode.insertData(startRangeOffset, insertNode.nodeValue);
            newRange.setEnd(startRangeNode, startRangeOffset + insertNode.length);
            newRange.setStart(startRangeNode, startRangeOffset + insertNode.length);
        }
		else {
            var afterNode;
            if (startRangeNode.nodeType == 3) {
                var textNode = startRangeNode;
                startRangeNode = textNode.parentNode;
                var text = textNode.nodeValue;
                var textBefore = text.substr(0, startRangeOffset);
                var textAfter = text.substr(startRangeOffset);
                var beforeNode = document.createTextNode(textBefore);
                var afterNode = document.createTextNode(textAfter);
                startRangeNode.insertBefore(afterNode, textNode);
                startRangeNode.insertBefore(insertNode, afterNode);
                startRangeNode.insertBefore(beforeNode, insertNode);
                startRangeNode.removeChild(textNode);
            }
			else {
				if (startRangeNode.tagName.toLowerCase() == 'html') {
					startRangeNode = startRangeNode.childNodes[0].nextSibling;
					afterNode = startRangeNode.childNodes[0];
				}
				else {
					afterNode = startRangeNode.childNodes[startRangeOffset];
				}
				startRangeNode.insertBefore(insertNode, afterNode);
            }
            newRange.setEnd(afterNode, 0);
            newRange.setStart(afterNode, 0);
        }
        YuZi_SELECTION.addRange(newRange);
	}
}

//清洁代码
function ClearCode(htmlStr,Type){
	htmlStr = htmlStr.replace(/\<p>/gi,"[$p]");
	htmlStr = htmlStr.replace(/\<\/p>/gi,"[$\/p]");
	htmlStr = htmlStr.replace(/\&lt;/gi,"<");
	htmlStr = htmlStr.replace(/\&gt;/gi,">");
	htmlStr = htmlStr.replace(/\<br>/gi,"[$br]");
	htmlStr = htmlStr.replace(/\<br \/>/gi,"[$br]");
	htmlStr = htmlStr.replace(/\<[^>]*>/g,"");
	if (Type == 1){
		htmlStr = htmlStr.replace(/\[\$p\]/gi,"<p>");
		htmlStr = htmlStr.replace(/\[\$\/p\]/gi,"<\/p>");
		htmlStr = htmlStr.replace(/\[\$br\]/gi,"<br>");
	}
	else {
		htmlStr = htmlStr.replace(/\[\$p\]/gi,"");
		htmlStr = htmlStr.replace(/\[\$\/p\]/gi,",");
		htmlStr = htmlStr.replace(/\[\$br\]/gi,",");
	}
	return htmlStr;
}

function BBSXPExecuteValue(cmd, value)
{
	if (value == 'YuZi_ForeColor') value = YuZi_ForeColor;
	YuZi_EDITFORM_DOCUMENT.execCommand(cmd, false, value);
}
function BBSXPSimpleExecute(cmd,option)
{
	BBSXPEditorForm.focus();
	YuZi_EDITFORM_DOCUMENT.execCommand(cmd,false,option);
	BBSXPDisableMenu();
}

//格式实现
function BBSXPExecute(cmd, value)
{
	switch (cmd)
	{
		case 'Break':
			BBSXPEditorForm.focus();
			BBSXPSelection();
			var element = document.createElement("br");
			BBSXPInsertItem(element);
			BBSXPSelect();
			BBSXPDisableMenu();
			break;
		case 'YuZi_UNDO':
			BBSXPSimpleExecute('undo');
			break;
		case 'YuZi_REDO':
			BBSXPSimpleExecute('redo');
			break;
		case 'YuZi_CUT':
			YuZi_BROWSER == 'IE' ? BBSXPSimpleExecute('cut') : alert(YuZi_LANG['CopyPaste']);
			break;
		case 'YuZi_COPY':
			YuZi_BROWSER == 'IE' ? BBSXPSimpleExecute('copy') : alert(YuZi_LANG['CopyPaste']);
			break;
		case 'YuZi_PASTE':
			YuZi_BROWSER == 'IE' ? BBSXPSimpleExecute('paste') : alert(YuZi_LANG['CopyPaste']);
			break;
		case 'YuZi_SELECTALL':
			BBSXPSimpleExecute('selectall');
			break;
		case 'YuZi_UNSELECT':
			BBSXPSimpleExecute('unselect');
			break;
		case 'YuZi_Delete':
			BBSXPSimpleExecute('Delete');
			break;
		case 'YuZi_BOLD':
			BBSXPSimpleExecute('bold');
			break;
		case 'YuZi_ITALIC':
			BBSXPSimpleExecute('italic');
			break;
		case 'YuZi_UNDERLINE':
			BBSXPSimpleExecute('underline');
			break;
		case 'YuZi_STRIKE':
			BBSXPSimpleExecute('strikethrough');
			break;
		case 'YuZi_JUSTIFYLEFT':
			BBSXPSimpleExecute('justifyleft');
			break;
		case 'YuZi_JUSTIFYCENTER':
			BBSXPSimpleExecute('justifycenter');
			break;
		case 'YuZi_JUSTIFYRIGHT':
			BBSXPSimpleExecute('justifyright');
			break;
		case 'YuZi_JUSTIFYFULL':
			BBSXPSimpleExecute('justifyfull');
			break;
		case 'YuZi_NUMBEREDLIST':
			BBSXPSimpleExecute('insertorderedlist');
			break;
		case 'YuZi_UNORDERLIST':
			BBSXPSimpleExecute('insertunorderedlist');
			break;
		case 'YuZi_REMOVE':
			BBSXPSimpleExecute('removeformat');
			break;
		case 'YuZi_ClearUP':
			YuZi_EDITFORM_DOCUMENT.body.innerHTML = ClearCode(YuZi_EDITFORM_DOCUMENT.body.innerHTML,1);
			break;
		case 'YuZi_FONTNAME':
			BBSXPDisplayMenu(cmd);
			break;
		case 'YuZi_FONTNAME_END':
			BBSXPEditorForm.focus();
			BBSXPSelect();
			BBSXPExecuteValue('fontname', value);
			BBSXPDisableMenu();
			break;
		case 'YuZi_FONTSIZE':
			BBSXPDisplayMenu(cmd);
			break;
		case 'YuZi_FONTSIZE_END':
			BBSXPEditorForm.focus();
			value = value.substr(0, 1);
			BBSXPSelect();
			BBSXPExecuteValue('fontsize', value);
			BBSXPDisableMenu();
			break;
		case 'YuZi_TEXTCOLOR':
			BBSXPDisplayMenu(cmd);
			break;
		case 'YuZi_TEXTCOLOR_END':
			BBSXPEditorForm.focus();
			BBSXPSelect();
			BBSXPExecuteValue('ForeColor', value);
			BBSXPDisableMenu();
			$('YuZi_TEXTCOLORBar').style.background = value;
			YuZi_ForeColor = value;
			break;
		case 'YuZi_ICON':
			BBSXPDisplayMenu(cmd);
			break;
		case 'YuZi_ICON_END':
			BBSXPEditorForm.focus();
			var element = document.createElement("img");
			element.src = value;
			element.border = 0;
			BBSXPSelect();
			BBSXPInsertItem(element);
			BBSXPDisableMenu();
			break;
		case 'YuZi_Replace':
			BBSXPDisplayMenu(cmd);
			BBSXPDrawIframe(cmd);
			BBSXPYuZi_ReplaceIframe.focus();
			YuZi_Replace_DOCUMENT.getElementById(cmd+'submitButton').focus();
			break;
		case 'YuZi_ED2K':
			BBSXPDisplayMenu(cmd);
			BBSXPDrawIframe(cmd);
			BBSXPYuZi_ED2KIframe.focus();
			YuZi_ED2K_DOCUMENT.getElementById(cmd+'submitButton').focus();
			break;
		case 'YuZi_IMAGE':
			BBSXPDisplayMenu(cmd);
			BBSXPDrawIframe(cmd);
			BBSXPYuZi_IMAGEIframe.focus();
			YuZi_IMAGE_DOCUMENT.getElementById(cmd+'submitButton').focus();
			break;
		case 'YuZi_Media':
			BBSXPDisplayMenu(cmd);
			BBSXPDrawIframe(cmd);
			BBSXPYuZi_MediaIframe.focus();
			YuZi_Media_DOCUMENT.getElementById(cmd+'submitButton').focus();
			break;
		case 'YuZi_LINK':
			if (YuZi_BROWSER == 'IE') {
				BBSXPSimpleExecute('createLink');
			}
			else {
				BBSXPDisplayMenu(cmd);
				BBSXPDrawIframe(cmd);
				BBSXPYuZi_LINKIframe.focus();
				YuZi_LINK_DOCUMENT.getElementById(cmd+'submitButton').focus();
			}
			break;
		case 'YuZi_UNLINK':
			BBSXPSimpleExecute('unlink');
			break;
		case 'YuZi_SPECIALCHAR':
			BBSXPDisplayMenu(cmd);
			break;
		case 'YuZi_SPECIALCHAR_END':
			BBSXPEditorForm.focus();
			BBSXPSelect();
			var element = document.createElement("span");
			element.appendChild(document.createTextNode(value));
			BBSXPInsertItem(element);
			BBSXPDisableMenu();
			break;
		case 'YuZi_TABLE':
			BBSXPDisplayMenu(cmd);
			break;
		case 'YuZi_TABLE_END':
			BBSXPEditorForm.focus();
			var location = value.split(',');
			var element = document.createElement("table");
			element.cellPadding = 0;
			element.cellSpacing = 0;
			element.border = 1;
			element.style.width = "100%";
			for (var i = 0; i < location[0]; i++) {
				var rowElement = element.insertRow(i);
				for (var j = 0; j < location[1]; j++) {
					var cellElement = rowElement.insertCell(j);
					cellElement.innerHTML = "&nbsp;";
				}
			}
			BBSXPSelect();
			BBSXPInsertItem(element);
			BBSXPDisableMenu();
			break;
		case 'YuZi_DATE':
			BBSXPEditorForm.focus();
			BBSXPSelection();
			var element = document.createElement("span");
			element.appendChild(document.createTextNode(new Date().toLocaleDateString()));
			BBSXPInsertItem(element);
			BBSXPDisableMenu();
			break;
		case 'YuZi_TIME':
			BBSXPEditorForm.focus();
			BBSXPSelection();
			var element = document.createElement("span");
			element.appendChild(document.createTextNode(new Date().toLocaleTimeString()));
			BBSXPInsertItem(element);
			BBSXPDisableMenu();
			break;
		case 'YuZi_QUOTE':
			BBSXPEditorForm.focus();
			var element = document.createElement("span");
			element.innerHTML = value;
			BBSXPSelection();
			BBSXPInsertItem(element);
			BBSXPDisableMenu();	
			break;
		default:
			break;
	}
}
// 改变模式：代码、编辑
function setMode(NewMode){
	if(NewMode!=YuZi_CurrentMode){
		YuZi_IsChangeMode = true;
		var obj=$("BBSXPEditorIframe");
		if(YuZi_TdHeight == -1)YuZi_TdHeight = parseInt(obj.offsetHeight);
		switch(NewMode){
			case 'DESIGN':
				obj.style.height = (parseInt(obj.offsetHeight) - 55) + 'px';
				YuZi_EDITFORM_DOCUMENT.body.innerHTML = bbcode2html($("BBSXPCodeForm").value,false);
				$("BBSXPEditorForm").style.display = 'block';
				$("BBSXPToolBar").style.display = 'block';
				$("BBSXPCodeForm").style.display = 'none';
				
				var EditorFormObject =$("BBSXPEditorForm");
				EditorFormObject.style.height = parseInt(obj.offsetHeight) + 'px';
				if (parseInt(EditorFormObject.offsetHeight)< YuZi_TdHeight){EditorFormObject.style.height = YuZi_TdHeight + 'px';}

				break;
			case 'CODE':
				obj.style.height = (parseInt(obj.offsetHeight) + 51) + 'px';
				
				$("BBSXPCodeForm").value = html2bbcode(YuZi_EDITFORM_DOCUMENT.body.innerHTML,false);
				$("BBSXPToolBar").style.display = 'none';
				$("BBSXPEditorForm").style.display = 'none';
				$("BBSXPCodeForm").style.display = 'block';
				$("BBSXPCodeForm").focus();
				
				var CodeFormObject=$("BBSXPCodeForm");
				CodeFormObject.style.height = parseInt(obj.offsetHeight) + 'px';
				if (parseInt(CodeFormObject.offsetHeight)< YuZi_TdHeight){CodeFormObject.style.height = YuZi_TdHeight + 'px';}

				break;
			default:
				break;
		}
		if (parseInt(obj.offsetHeight)< YuZi_TdHeight){obj.style.height = YuZi_TdHeight + 'px';}
		try{
			$("HtmlEditor_DESIGN").className = "StatusBarBtnOff";
			$("HtmlEditor_CODE").className = "StatusBarBtnOff";
			$("HtmlEditor_"+NewMode).className = "StatusBarBtnOn";
		}
		catch(e){
		}
		YuZi_CurrentMode = NewMode;
	}
}

// 调整编辑器的大小
function ChangeSize(size){
	var obj=$("BBSXPEditorIframe");
	if(YuZi_TdHeight == -1)YuZi_TdHeight = parseInt(obj.offsetHeight);
	obj.style.height = (parseInt(obj.offsetHeight) + size) + 'px';
	if (parseInt(obj.offsetHeight)< YuZi_TdHeight){obj.style.height = YuZi_TdHeight + 'px';}
	
	if (YuZi_IsChangeMode) {
		if (YuZi_CurrentMode == 'DESIGN') {
			obj=$("BBSXPEditorForm");
			obj.style.height = (parseInt(obj.offsetHeight) + size) + 'px';
			if (parseInt(obj.offsetHeight)< YuZi_TdHeight){obj.style.height = YuZi_TdHeight + 'px';}
		}
		else {
			if (parseInt(obj.offsetHeight)< YuZi_TdHeight+77){obj.style.height = (YuZi_TdHeight+77) + 'px';}
			
			obj=$("BBSXPCodeForm");
			obj.style.height = (parseInt(obj.offsetHeight) + size) + 'px';
			if (parseInt(obj.offsetHeight)< YuZi_TdHeight+77){obj.style.height = (YuZi_TdHeight+77) + 'px';}
		}
	}
}

//编辑器工具栏各图标按钮
function BBSXPCreateIcon(icon)
{
	var str;
	if (icon[0] == 'YuZi_TEXTCOLOR') {
		ExecuteCMD = 'ForeColor';
		
		str = '<div style="POSITION: relative;border:1px solid ' + YuZi_TOOLBAR_BG_COLOR +';cursor:pointer;height:20px;width:28px" onmouseover="$(\''+icon[0]+'RightBar\').style.borderLeft = this.style.border=\'1px solid ' + YuZi_MENU_BORDER_COLOR + '\';" onmouseout="$(\''+icon[0]+'RightBar\').style.borderLeft = this.style.border=\'1px solid ' + YuZi_TOOLBAR_BG_COLOR + '\';" title="' + icon[2]  + '">';
		str += '<img src="' + YuZi_SKIN_PATH  + icon[1] +'" onclick="BBSXPExecuteValue(\''+ExecuteCMD+'\',\'YuZi_'+ExecuteCMD+'\');">';
		str += '<img id="'+icon[0]+'Bar" style="BACKGROUND: '+eval("YuZi_"+ExecuteCMD)+'; LEFT: 1px; POSITION: absolute; TOP: 15px" height=4 src="' + YuZi_SKIN_PATH  +'clear.gif" width=16 onclick="BBSXPExecuteValue(\''+ExecuteCMD+'\',\'YuZi_'+ExecuteCMD+'\');">';
		str += '<div id="'+icon[0]+'RightBar" style="POSITION: absolute;left:17px;top:0px;border-left:1px solid '+YuZi_TOOLBAR_BG_COLOR+';height:18px"><img id="'+ icon[0] +'" src="' + YuZi_SKIN_PATH + 'PopMenu.gif" title="' + icon[2]  + '" align="absmiddle" onclick="BBSXPExecute(\''+ icon[0] +'\');"></div>';
		str += '</div>';
	}
	else {
		str = '<img id="'+ icon[0] +'" src="' + YuZi_SKIN_PATH + icon[1] + '" title="' + icon[2]  + 
			'" align="absmiddle" style="border:1px solid ' + YuZi_TOOLBAR_BG_COLOR +';cursor:pointer;height:20px;';
		str += '" onclick="BBSXPExecute(\''+ icon[0] +'\');" '+
			'onmouseover="this.style.border=\'1px solid ' + YuZi_MENU_BORDER_COLOR + '\';" ' +
			'onmouseout="this.style.border=\'1px solid ' + YuZi_TOOLBAR_BG_COLOR + '\';" ';
		str += '>';
	}
	return str;
}
//编辑器工具栏
function BBSXPCreateToolbar()
{
	var htmlData = '<table cellpadding="0" cellspacing="0" border="0" height="26"><tr>';
	
	for (i = 0; i < YuZi_TOP_TOOLBAR_ICON.length; i++) {
		htmlData += '<td style="padding:2px;">' + BBSXPCreateIcon(YuZi_TOP_TOOLBAR_ICON[i]) + '</td>';
	}
		
	htmlData += '</tr></table><table cellpadding="0" cellspacing="0" border="0" height="26"><tr>';
	for (i = 0; i < YuZi_BOTTOM_TOOLBAR_ICON.length; i++) {
		htmlData += '<td style="padding:2px;">' + BBSXPCreateIcon(YuZi_BOTTOM_TOOLBAR_ICON[i]) + '</td>';
	}
		
	htmlData += '</tr></table>';
	return htmlData;
}

function BBSXPWriteFullHtml(documentObj, content)
{
	var editHtmlData = '';
	editHtmlData += '<html>\r\n<head>\r\n<title>BBSXP Editor --Powered By BBSXP</title>\r\n';
	editHtmlData += '<style type="text/css">\r\n\tbody { font-family:Courier New;font-size:12px;margin:2px;}\r\n\tBLOCKQUOTE {BORDER: #cccccc 1px dotted; PADDING: 4px; MARGIN: 16px;}\r\n</style>\r\n';
	editHtmlData += '</head>\r\n<body>\r\n';
	editHtmlData += bbcode2html(content,true);
	editHtmlData += '\r\n</body>\r\n</html>\r\n';
	documentObj.open();
	documentObj.write(editHtmlData);
	documentObj.close();
}

//编辑器主函数
function BBSXPEditor(objName) 
{
	this.objName = objName;
	this.hiddenName = objName;
	this.safeMode;
	this.editorWidth;
	this.editorHeight;
	this.iconPath;
	this.menuBorderColor;
	this.menuBgColor;
	this.menuTextColor;
	this.menuSelectedColor;
	this.toolbarBorderColor;
	this.toolbarBgColor;
	this.formBorderColor;
	this.formBgColor;
	this.buttonColor;
	this.init = function()
	{
		if (this.safeMode) YuZi_SAFE_MODE = this.safeMode;
		if (this.editorWidth) YuZi_WIDTH = this.editorWidth;
		if (this.editorHeight) YuZi_HEIGHT = this.editorHeight;
		if (this.menuBorderColor) YuZi_MENU_BORDER_COLOR = this.menuBorderColor;
		if (this.menuBgColor) YuZi_MENU_BG_COLOR = this.menuBgColor;
		if (this.menuTextColor) YuZi_MENU_TEXT_COLOR = this.menuTextColor;
		if (this.menuSelectedColor) YuZi_MENU_SELECTED_COLOR = this.menuSelectedColor;
		if (this.toolbarBorderColor) YuZi_TOOLBAR_BORDER_COLOR = this.toolbarBorderColor;
		if (this.toolbarBgColor) YuZi_TOOLBAR_BG_COLOR = this.toolbarBgColor;
		if (this.formBorderColor) YuZi_FORM_BORDER_COLOR = this.formBorderColor;
		if (this.formBgColor) YuZi_FORM_BG_COLOR = this.formBgColor;
		if (this.buttonColor) YuZi_BUTTON_COLOR = this.buttonColor;
		YuZi_OBJ_NAME = this.objName;
		YuZi_BROWSER = BBSXPGetBrowser();
	}
	this.show = function()
	{
		this.init();
		var widthStyle = 'width:' + YuZi_WIDTH + ';';
		var heightStyle = 'height:' + YuZi_HEIGHT + ';';
		if (YuZi_BROWSER == '') {
			var htmlData = '<div id="BBSXPEditorIframe" style="' + widthStyle + heightStyle +';background-color:'+ YuZi_FORM_BG_COLOR +'">' +
			'<textarea name="BBSXPCodeForm" id="BBSXPCodeForm" style="' + widthStyle + heightStyle + 
			'padding:0;margin:0;border:1px solid '+ YuZi_FORM_BORDER_COLOR + 
			';font-size:12px;line-height:16px;font-family:'+YuZi_FONT_FAMILY+';background-color:'+ 
			YuZi_FORM_BG_COLOR +';">' + document.getElementsByName(this.hiddenName)[0].value + '</textarea></div>';
			document.open();
			document.write(htmlData);
			document.close();
			return;
		}
		var htmlData = '<style>TD.StatusBarBtnOff {padding:1px 5px;border:1px outset;cursor:pointer;}TD.StatusBarBtnOn {padding:1px 5px;border:1px inset;background-color: #EEEEEE;}</style>';
		
		htmlData += '<div id=BBSXPToolBar style="'+widthStyle+';border:1px solid ' + YuZi_TOOLBAR_BORDER_COLOR + ';border-bottom:0;background-color:'+ YuZi_TOOLBAR_BG_COLOR +'">';
		htmlData += BBSXPCreateToolbar();
		htmlData += '</div><div id="BBSXPEditorIframe" style="' + widthStyle + heightStyle + 
			'border:1px solid '+ YuZi_FORM_BORDER_COLOR +';background-color:'+ YuZi_FORM_BG_COLOR +'">' +
			'<iframe name="BBSXPEditorForm" id="BBSXPEditorForm" style="' + widthStyle + 
			'height:100%;padding:0;margin:0;border:0;"></iframe>';

		
		htmlData += '<textarea name="BBSXPCodeForm" id="BBSXPCodeForm" style="' + widthStyle + 
				'height:100%;padding:0;margin:0;border:0;font-size:12px;display:none;line-height:16px;font-family:'+YuZi_FONT_FAMILY+';background-color:'+ 
		YuZi_FORM_BG_COLOR +';" onclick="javascirit:parent.BBSXPDisableMenu();"></textarea></div>';
		
		htmlData += '<table border=0 cellpadding=0 cellspacing=0 height=20 style="margin-top:5px;">';
		htmlData += '<tr>';
		htmlData += '<td class=StatusBarBtnOn id=HtmlEditor_DESIGN onclick="setMode(\'DESIGN\')"><img border=0 src='+YuZi_SKIN_PATH+'modeedit.gif align=absmiddle title="'+YuZi_LANG['DesignMode']+'"></td>';
		htmlData += '<td width=5>&nbsp;</td>';
		htmlData += '<td class=StatusBarBtnOff id=HtmlEditor_CODE onclick="setMode(\'CODE\')"><img border=0 src='+YuZi_SKIN_PATH+'modecode.gif align=absmiddle title="'+YuZi_LANG['CodeMode']+'"></td>';
		htmlData += '<td width=5>&nbsp;</td>';
		htmlData += '<td align=right width=100%><img border=0 src='+YuZi_SKIN_PATH+'expand.gif align=absmiddle onclick=ChangeSize(100) style=cursor:pointer title="'+YuZi_LANG['AddEditorArea']+'"><img border=0 src='+YuZi_SKIN_PATH+'contract.gif align=absmiddle onclick=ChangeSize(-100) style=cursor:pointer title="'+YuZi_LANG['ReduceEditorArea']+'"></td>';
		htmlData += '</tr></table>';

		htmlData +='<div id="Yuzi_TempDiv" style="padding:2px;'+BBSXPGetMenuCommonStyle()+'">Loading...</div>';

		document.open();
		document.write(htmlData);
		document.close();
		if (YuZi_BROWSER == 'IE') {
			YuZi_EDITFORM_DOCUMENT = document.frames("BBSXPEditorForm").document;
		} else {
			YuZi_EDITFORM_DOCUMENT = $('BBSXPEditorForm').contentDocument;
		}

		YuZi_EDITFORM_DOCUMENT.designMode = 'On';
		YuZi_CurrentMode = "DESIGN";
		BBSXPWriteFullHtml(YuZi_EDITFORM_DOCUMENT, document.getElementsByName(eval(YuZi_OBJ_NAME).hiddenName)[0].value);
		var el = YuZi_EDITFORM_DOCUMENT.body;
		
		if (YuZi_EDITFORM_DOCUMENT.addEventListener){
			YuZi_EDITFORM_DOCUMENT.addEventListener('click', BBSXPDisableMenu, false);
			YuZi_EDITFORM_DOCUMENT.addEventListener('keypress', ctlent, true);
		} else if (el.attachEvent){
			el.attachEvent('onclick', BBSXPDisableMenu);
			el.attachEvent('onkeypress', ctlent);
		}
	}
	this.data = function()
	{
		var htmlResult;
		if (YuZi_BROWSER == '') {
			htmlResult = $("BBSXPCodeForm").value;
		}
		else {
			if (YuZi_CurrentMode.toLowerCase() == 'design') {
				htmlResult = html2bbcode(YuZi_EDITFORM_DOCUMENT.body.innerHTML,true);	//BBCode模式提交
			}
			else {
				htmlResult = $("BBSXPCodeForm").value;	//BBCode模式提交;
			}
		}
		BBSXPDisableMenu();
		
		htmlResult = ChangeImgPath(htmlResult);
		
		document.getElementsByName(this.hiddenName)[0].value = htmlResult;
		return htmlResult;
	}
}

//初始化编辑器
var editor = new BBSXPEditor("editor");
editor.hiddenName = "Body";
editor.show();

function ChangeImgPath(str){
	var url = location.href;
	var siteUrl = url.split("/");
	siteUrl.length = siteUrl.length-1;
	var siteStr = siteUrl.join("/") +"/";
	var re=new RegExp(siteStr,'ig');
	
	return str.replace(re,"");
}

// 预览
function Preview()
{
	editor.data();
	Subject=document.form.Subject;
	Body=document.form.Body;

 	if(typeof(Subject) != "undefined" ){
		if(Subject.value.length<2){alert(YuZi_LANG['SubjectMinLen']);return false;}
  	}

 	if(typeof(Body) != "undefined" ){
		if(Body.value.length<2){alert(YuZi_LANG['BodyMinLen']);return false;}
		if(Body.value.length>60000){alert(YuZi_LANG['BodyMaxLen']);return false;}
  	}

	$('Preview').style.display = '';
	Ajax_CallBack('form','Preview','Loading.asp?menu=Preview');
	window.location="#Preview";
}

//内容提交
function CheckForm() {
	editor.data();
	Subject=document.form.Subject;
	Body=document.form.Body;
	VerifyCode=document.form.VerifyCode;
	Category=document.form.Category;

 	if(typeof(Category) != "undefined" ){
		if(Category.value.length>25){alert(YuZi_LANG['CategoryMaxLen']);return false;}
  	}

 	if(typeof(Subject) != "undefined" ){
		if(Subject.value.length<2){alert(YuZi_LANG['SubjectMinLen']);return false;}
		if(Subject.value.length>100){alert(YuZi_LANG['SubjectMaxLen']);return false;}
  	}

 	if(typeof(Body) != "undefined" ){
		if(Body.value.length<2){alert(YuZi_LANG['BodyMinLen']);return false;}
		if(Body.value.length>60000){alert(YuZi_LANG['BodyMaxLen']);return false;}
  	}
	
	if(typeof(document.form.Description) != "undefined"){
		document.form.Description.value = ClearCode(bbcode2html(Body.value,false),2);
	}
	
	if(typeof(document.form.Tags) != "undefined"){
		Tags = document.form.Tags.value;
		TagsArray = Tags.split(',');
		if (TagsArray.length > 5){
			alert('您最多只能添加 5 个标签。');
			return false;
		}
		for (var i=0; i<TagsArray.length;i++) {
			if(TagsArray[i].length > 25){
				alert('每个标签的长度不能大于 25 个字符');
				return false;
			}
		}
	}
	
 	if(typeof(document.form.RequireEditNotes) != "undefined" ){
		if(document.form.EditNotes.value.length<2){alert('请输入帖子编辑原因');return false;}
  	}
	
	if(typeof(VerifyCode) != "undefined"){
		if(VerifyCode.value.length<4){alert(YuZi_LANG['VerifyCode']);return false;}
  	}
}
//检测文章长度
function CheckLength(){
	editor.data();
	MessageLength=document.form.Body.value.length;
	alert("最大字符为 "+60000+ " 字节\n您的内容已有 "+MessageLength+" 字节");
}

//Ctrl + Enter 发帖
function ctlent(event){
	if((event.ctrlKey && (event.keyCode==13 || event.keyCode==10)) || (event.altKey && event.keyCode == 83)){
		parent.document.form.EditSubmit.click();
	}
}



// Upadate at 2008-1-16
document.write('<style type="text/css">.autosave { behavior: url(#default#userdata);}</style>');
document.write('<div id=AutoSave class="autosave"></div>');
var UserDataObject = document.getElementById('AutoSave');
window.onbeforeunload = function () {saveData()};

function saveData() {
	editor.data();
	PostBody = document.form.Body.value;
	if(!PostBody) return;
	
	var formdata = '';
	formdata += 'Body' + String.fromCharCode(9) + 'Editor_Frame' + String.fromCharCode(9) + 'DocumentFrame' + String.fromCharCode(9) + PostBody + String.fromCharCode(9, 9);
	for(var i = 0; i < document.form.elements.length; i++) {
		var el = document.form.elements[i];
		if(el.name != '' && (el.tagName == 'TEXTAREA' || el.tagName == 'INPUT' && (el.type == 'text' || el.type == 'checkbox' || el.type == 'radio')) && el.name.substr(0, 6) != 'attach') {
			var elvalue = el.value;
			if((el.type == 'checkbox' || el.type == 'radio') && !el.checked) {
				continue;
			}
			formdata += el.name + String.fromCharCode(9) + el.tagName + String.fromCharCode(9) + el.type + String.fromCharCode(9) + elvalue + String.fromCharCode(9, 9);
		}
	}
	ResponseUserData(formdata);
}

function RestoreData() {
	var message = RequestUserData();
	if(InArray(message, ['', 'null', 'false', null, false])) {
		alert('对不起，暂时没有可以恢复的数据！');
		return;
	}
	if(!window.confirm('此操作将覆盖当前表单内容，确定要恢复数据吗？')) {
		return;
	}

	var formdata = message.split(/\x09\x09/);
	for(var i = 0; i < document.form.elements.length; i++) {
		var el = document.form.elements[i];
		if(el.name != '' && (el.tagName == 'TEXTAREA' || el.tagName == 'INPUT' && (el.type == 'text' || el.type == 'checkbox' || el.type == 'radio'))) {
			for(var j = 0; j < formdata.length; j++) {
				var ele = formdata[j].split(/\x09/);
				if(ele[0] == 'Body') {
					YuZi_EDITFORM_DOCUMENT.designMode = 'On';
					setMode('DESIGN');
					BBSXPWriteFullHtml(YuZi_EDITFORM_DOCUMENT, ele[3]);
				}
				else if(ele[0] == el.name) {
					elvalue = typeof(ele[3]) != "undefined" ? ele[3] : '';
					if(ele[1] == 'INPUT') {
						if(ele[2] == 'text') {
							el.value = elvalue;
						} else if((ele[2] == 'checkbox' || ele[2] == 'radio') && ele[3] == el.value) {
							el.checked = true;
							EvalEvent(el);
						}
					}
					else if(ele[1] == 'TEXTAREA') {
						el.value = elvalue;
					}
					break;
				}
			}
		}
	}
}
function RequestUserData(){
	var message = '';
	if(BBSXPGetBrowser() == 'IE') {
		try {
			UserDataObject.load('BBSXP');
			var oXMLDoc = UserDataObject.XMLDocument;
			var nodes = oXMLDoc.documentElement.childNodes;
			message = nodes.item(nodes.length - 1).getAttribute('message');
		} catch(e) {}
	}
	else if(window.sessionStorage) {
		try {
			message = sessionStorage.getItem('BBSXP');
		} catch(e) {}
	}
	message = message.toString().replace(/(\s+)$/g, '').replace(/^\s+/g, '');
	return message;
}
function ResponseUserData(formdata){
	if(BBSXPGetBrowser() == 'IE') {
		try {
			var oXMLDoc = UserDataObject.XMLDocument;
			var root = oXMLDoc.firstChild;
			if(root.childNodes.length > 0) {
				root.removeChild(root.firstChild);
			}
			var node = oXMLDoc.createNode(1, 'POST', '');
			var oTimeNow = new Date();
			oTimeNow.setHours(oTimeNow.getHours() + 24);
			UserDataObject.expires = oTimeNow.toUTCString();
			node.setAttribute('message', formdata);
			oXMLDoc.documentElement.appendChild(node);
			UserDataObject.save('BBSXP');
		} catch(e) {}
	}
	else if(window.sessionStorage) {
		try {
			sessionStorage.setItem('BBSXP', formdata);
		} catch(e) {}
	}
}

function EvalEvent(obj) {
	var script = obj.parentNode.innerHTML;
	var re = /onclick="(.+?)["|>]/ig;
	var matches = re.exec(script);
	if(matches != null) {
		matches[1] = matches[1].replace(/this\./ig, 'obj.');
		eval(matches[1]);
	}
}
function InArray(needle, haystack) {
	if(typeof(needle) == 'string' || typeof (needle) == 'number') {
		for(var i in haystack) {
			if(haystack[i] == needle) {
					return true;
					break;
			}
		}
	}
	return false;
}