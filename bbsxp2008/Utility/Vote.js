document.write("<img id=imgStar1 src=Images/Star/star-off.gif><img id=imgStar2 src=Images/Star/star-off.gif><img id=imgStar3 src=Images/Star/star-off.gif><img id=imgStar4 src=Images/Star/star-off.gif><img id=imgStar5 src=Images/Star/star-off.gif>");

if (UserAgent.indexOf("firefox")!=-1 || UserAgent.indexOf("chrome")!=-1) {//如果浏览器不支持 attachEvent/detachEvent 请在此处添加
	Event.prototype.__defineGetter__("srcElement", function () {
		var node = this.target;
		while (node.nodeType != 1) node = node.parentNode;
		return node;
	});
}
function voteBase(nValue)
{
	if(ThreadID != null)
	{
		var xmlHttp = Ajax_GetXMLHttpRequest();
		var RequestUrl = "Loading.asp?Menu=Threadstar&ThreadID=" + ThreadID + "&Rate=" + nValue + "&" + Math.random();
		xmlHttp.open("POST", RequestUrl);
		xmlHttp.onreadystatechange = function(){
			if (xmlHttp.readyState == 4 && xmlHttp.status == 200)
			{
			showData(xmlHttp.responseText);
			}
		};
		xmlHttp.send('');
	}
}

function showData(avgVoteScore)
{
	var canBeFull = false;
	for(var i=5; i>=1; i--)
	{
		temp = $("imgStar" + i);
		if(temp != null)
		{
			if(canBeFull == false)
			{
				if(i - avgVoteScore <= 0.5 && temp != null)
				{
					canBeFull = true;
					temp.src = i - avgVoteScore <0.5 ? "Images/Star/star-on.gif" : "Images/Star/star-half.gif";
				}
				else
					temp.src = "Images/Star/star-off.gif"
			}
			else
			{
				temp.src = "Images/Star/star-on.gif";
			}

			temp.nValue = i;
			temp.defaultImage = temp.src;
		}
	}
	setEvent();
}

function votePoint(event)
{
	var evt = window.event ? window.event : event;
	var element = evt.srcElement;
	if(element != null && confirm('您确定将其评为 '+element.nValue+' 星级?'))
	{
		voteBase(element.nValue);
	}
}
var TigArray = ["非常差", "差", "好", "很好", "非常好"];

function OnMouseOverImgStar(event)
{
	if(typeof(imgStarMouseEvent) != "undefined")
		clearTimeout(imgStarMouseEvent);
	var evt = window.event ? window.event : event;
	var element = evt.srcElement;
	if(element != null)
	{
		var n = element.nValue;
		for(var i=1; i<=5; i++)
		{
			var temp = $("imgStar" + i);
			if(temp != null)
			{
				temp.src = i <= n ? "Images/Star/star-on.gif" : "Images/Star/star-off.gif";
			}
			temp.alt=TigArray[i-1];
		}
	}
}

var imgStarMouseEvent;

function OnMouseOutImgStar(event)
{
	imgStarMouseEvent = setTimeout(function(){
	for(var i=1; i<=5; i++)
	{
		var temp = $("imgStar" + i);
		if(temp != null)
			temp.src = temp.defaultImage;
		}
	}, 100);
}

function setEvent()
{
	for(var i=1; i<=5; i++)
	{
		var element = $("imgStar" + i);
		if(element != null)
		{
			if (UserAgent.indexOf("firefox")!=-1 || UserAgent.indexOf("chrome")!=-1){//Mozilla, Netscape, Firefox
				element.removeEventListener("click", votePoint, false);
				element.removeEventListener("mouseover", OnMouseOverImgStar, false);
				element.removeEventListener("mouseout", OnMouseOutImgStar, false);	
				element.addEventListener("click", votePoint, false);
				element.addEventListener("mouseover", OnMouseOverImgStar, false);
				element.addEventListener("mouseout", OnMouseOutImgStar, false);
			}
			else {
				element.detachEvent("onclick", votePoint);
				element.detachEvent("onmouseover", OnMouseOverImgStar);
				element.detachEvent("onmouseout", OnMouseOutImgStar);	
				element.attachEvent("onclick", votePoint);
				element.attachEvent("onmouseover", OnMouseOverImgStar);
				element.attachEvent("onmouseout", OnMouseOutImgStar);
			}
			element.style.cursor = "pointer";
		}
	}
}