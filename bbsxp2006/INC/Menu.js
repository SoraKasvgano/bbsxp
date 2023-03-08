
var icon=["images/expand.gif","images/collapse.gif","images/node.gif"]
for(i=0;i<icon.length;i++){
var tem=new Image()
tem.src=icon[i]
}

function menutree(obj,target,check){
	this.obj=obj;
	this.target=target
	this.child=0
	this.node=0
	this.msg=[]
	this.showCheck=check
	this.html="<table id='deeptree' onselectstart='return false' cellspacing=0 cellpadding=0 border=0>"
}
	
menutree.prototype.addFolder=function(txt,link,show){
	this.msg[this.node]=[txt,link?link:'']
	this.html+="<tr><td class='node'><nobr>&nbsp;&nbsp;<img src='"+(show?icon[1]:icon[0])+"' id='img"+this.child+"' border=0 align='absmiddle' onclick='"+this.obj+".expand("+this.child+")'><input type='checkbox' name='treeFolder' onclick='"+this.obj+".checkAll(this,"+this.child+")' style='display:"+(this.showCheck?'':'none')+"'><span onmouseover='doOver(this)' onmouseout='doOut(this)' onmousedown='"+this.obj+".Light(this,"+this.node+");"+this.obj+".expand("+this.child+")' title='"+txt+"'>"+txt+"</span></nobr></td></tr><tr id='child"+this.child+"' style='display:"+(show?'':'none')+"'><td class='node'>"
	this.html+="<table cellspacing=0 cellpadding=0 border=0 style='margin-left:18;'>"
	this.child++
	this.node++;
}

menutree.prototype.addNode=function(txt,link){
	this.msg[this.node]=[txt,link?link:'']
	this.html+="<tr><td class='node'><nobr>&nbsp;&nbsp;<img src='"+icon[2]+"' align='absmiddle' onclick='"+this.obj+".Light(this.nextSibling.nextSibling,"+this.node+")'><input type='checkbox' name='treeNode' onclick='"+this.obj+".parentCheck(this)' style='display:"+(this.showCheck?'':'none')+"'><span onmouseover='doOver(this)' onmouseout='doOut(this)' onmousedown='"+this.obj+".Light(this,"+this.node+")' title='"+txt+"'>"+txt+"</span></nobr></td></tr>"
	this.node++;
}

menutree.prototype.endFolder=function(){
	this.html+="</table></td></tr>"
}


menutree.prototype.expand=function(childNum,flag){
var isExpand=document.getElementById("child"+childNum).style.display
document.getElementById("img"+childNum).src=isExpand=='none'?icon[1]:icon[0]
document.getElementById("child"+childNum).style.display=isExpand=='none'?'':'none'
}

menutree.prototype.expandAll=function(flag){
	if(this.child>0)
	for(i=0;i<this.child;i++){
	document.getElementById("img"+i).src=flag?icon[1]:icon[0]
	document.getElementById("child"+i).style.display=flag?'':'none'
}
}

menutree.prototype.checkAll=function(obj,childNum){
	obj.blur()
	var child=document.getElementById("child"+childNum)
	var node=child.getElementsByTagName("INPUT")
	for(i=0;i<node.length;i++)node[i].checked=obj.checked
	this.parentCheck(obj)
}

menutree.prototype.parentCheck=function(obj){
	obj.blur()
	for(i=this.child-1;i>=0;i--){
	var checkParent=true
	var c=document.getElementById("child"+i)
	var node=c.getElementsByTagName("INPUT")
	for(j=0;j<node.length;j++)if(!node[j].checked)checkParent=false
	document.getElementById("img"+i).nextSibling.checked=checkParent
	}
}

menutree.prototype.getCheckedValue=function(){
	var value=[]
	var node=document.getElementById('deeptree').getElementsByTagName("INPUT")
	for(i=0;i<node.length;i++)if(node[i].checked&&node[i].name=="treeNode")value[value.length]=this.msg[i][0]
	return value

}

menutree.prototype.init=function(){
	this.html+=(this.node>0?"":"<tr><td>暂时没有数据</td></tr>")+"</table>"
	document.write(this.html)
}

var tem=null
function doOver(o){
o.className='NodeOver'
}

function doOut(o){
o.className=(tem==o?'NodeFocus':'')
}

menutree.prototype.Light=function(o,nodeNum){
if(!tem)tem=o
tem.className=''
o.className='NodeFocus'
tem=o
if(this.msg[nodeNum][1])window.open(this.msg[nodeNum][1],this.target);
}
