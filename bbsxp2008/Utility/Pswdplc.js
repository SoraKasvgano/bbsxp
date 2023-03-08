function isctype(character){
 if (character>=48 && character<=57) //Êý×Ö
  return 1;
 if (character>=65 && character<=90) //´óÐ´×ÖÄ¸
  return 2;
 if (character>=97 && character<=122) //Ð¡Ð´
  return 4;
 else
  return 8; //ÌØÊâ×Ö·û
}

function bitTotal(num){
 modes=0;
 for (i=0;i<4;i++){
  if (num & 1) modes++;
  num>>>=1;
 }
 return modes;
}

function checkStrong(sPW){
 Modes=0;
 for (i=0;i<sPW.length;i++){
  Modes|=isctype(sPW.charCodeAt(i));
 }
 return bitTotal(Modes);
}



function EvalPwd(pwd){
 O_color="#F1F1F1";
 L_color="#FF0000";
 M_color="#FFFF99";
 H_color="#00CC66";
 if (pwd.length<6){
  Lcolor=Mcolor=Hcolor=O_color;
 }
 else{
  S_level=checkStrong(pwd);
  switch(S_level)  {
   case 0:
    Lcolor=Mcolor=Hcolor=O_color;
   case 1:
    Lcolor=L_color;
    Mcolor=Hcolor=O_color;
    break;
   case 2:
    Lcolor=Mcolor=M_color;
    Hcolor=O_color;
    break;
   default:
    Lcolor=Mcolor=Hcolor=H_color;
    }
  }
 $("iWeak").style.background=Lcolor;
 $("iMedium").style.background=Mcolor;
 $("iStrong").style.background=Hcolor;
 return;
}