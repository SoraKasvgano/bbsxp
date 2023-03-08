document.writeln('<style type="text/css">');
document.writeln('.ImgStyle{ margin-left:50px; width:16px; height:16px;}');
document.writeln('.ImgStyle2{margin-left:5px;}');
document.writeln('</style>');


function AddInputLabel(i){
	var InputObject=document.createElement('input');
	InputObject.type='text';
	InputObject.id='Bbsxp_Input_'+i;
	InputObject.name='Bbsxp_Input_'+i;
	InputObject.size=50;
	InputObject.className='';
	return InputObject;
}
function AddImgLabel(src,clsName,ImgTitle){
	var ElementImg = document.createElement('img');
	ElementImg.src = 'Images/'+src;
	ElementImg.className = clsName;
	ElementImg.style.cursor = 'pointer';
	ElementImg.title = ImgTitle;
	return ElementImg;
}
function AddLabel(parentNode){	
	if (RealCount>MaxCount-1){alert('投票选项不能超过'+MaxCount+'个');return;}
	var i=CurrentCount;
	var ElementDiv = document.createElement('div');
	ElementDiv.id='Bbsxp_Div_'+i;

	var oInput=AddInputLabel(i);

	var ElementDelImg=AddImgLabel('delete.gif','ImgStyle','删除此项');
	var ElementPreImg=AddImgLabel('i_previous.gif','ImgStyle2','上移一位');
	var ElementNextImg=AddImgLabel('i_next.gif','ImgStyle2','下移一位');
	AddEvent(ElementDelImg,'click',function(){RemoveLabel(i);});
	AddEvent(ElementPreImg,'click',function(){Move(parentNode,i,'Pre');});
	AddEvent(ElementNextImg,'click',function(){Move(parentNode,i,'Next');});
		
	ElementDiv.appendChild(oInput);
	ElementDiv.appendChild(ElementDelImg);
	ElementDiv.appendChild(ElementPreImg);
	ElementDiv.appendChild(ElementNextImg);
	document.getElementById(parentNode).appendChild(ElementDiv);

	CurrentCount++;
	RealCount++;
}
function RemoveLabel(i){
	if (RealCount<=MinCount){alert('投票选项不能少于'+MinCount+'个');return;}
	if (window.confirm('确实删除此项？')){
		var oa=document.getElementById('Bbsxp_Div_'+i);
		if(oa){oa.parentNode.removeChild(oa);RealCount--;}
	}
}
function Move(NodeID,i,direction){
	var k=i,j,TempText;
	var CountArray=new Array(),NodeArray=new Array();
	var oa=document.getElementById('Bbsxp_Div_'+i);
	var Node=document.getElementById(NodeID);
	var child=Node.childNodes.length;
	for (j=0;j<child;j++){
		CountArray[j]=(Node.childNodes[j].id).split('_')[2];
	}
	if (direction=='Next'){
		for (j=0;j<CountArray.length;j++){
			if(CountArray[j]>k){k=CountArray[j];break;}
		}
	}
	else{
		for (j=CountArray.length-1;j>=0;j--){
			if(CountArray[j]<k){k=CountArray[j];break;alert(k);}
		}
	}
	if (k != i){
		TempText=document.getElementById('Bbsxp_Input_'+k).value;
		document.getElementById('Bbsxp_Input_'+k).value=document.getElementById('Bbsxp_Input_'+i).value;
		document.getElementById('Bbsxp_Input_'+i).value=TempText;
	}
}

function AddEvent(o,t,h){	//添加事件
	if (o.addEventListener){
		o.addEventListener(t,h,false);
	}
	else if(o.attachEvent){
		o.attachEvent('on'+t,h);
	}
	else{
		try{o['on'+t]=h;}catch(e){;}
	}
}

function init(){
	for (var i=0;i<MinCount;i++){
		AddLabel('VoteOptionList');
	}
}