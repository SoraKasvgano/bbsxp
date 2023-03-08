function getpet (birthday) {
abc=birthday.split("-")
if (abc.length < 3){return("");}
birthyear=(abc[0]*1)
b="ºï|¼¦|¹·|Öí|Êó|Å£|»¢|ÍÃ|Áú|Éß|Âí|Ñò";
var list= b.split ('|');
return(list[birthyear%12]);
}



function astro(birth)
{
abc=birth.split("-")
if (abc.length < 3){return("");}
mm=(abc[1]*1)
dd=(abc[2]*1)
if (mm=='1'){
if(dd>=21){return('<img src=images/star/h.gif alt=Ë®Æ¿×ù>');}
else{return('<img src=images/star/g.gif alt=Ä§ôÉ×ù>');}
}

if (mm=='2'){
if(dd>=20){return('<img src=images/star/i.gif alt=Ë«Óã×ù>');}
else{return('<img src=images/star/h.gif alt=Ë®Æ¿×ù>');}
}

if (mm=='3'){
if(dd>=21){return('<img src=images/star/^.gif alt=°×Ñò×ù>');}
else{return('<img src=images/star/i.gif alt=Ë«Óã×ù>');}
}

if (mm=='4'){
if(dd>=21){return('<img src=images/star/_.gif alt=½ğÅ£×ù>');}
else{return('<img src=images/star/^.gif alt=°×Ñò×ù>');}
}

if (mm=='5'){
if(dd>=22){return('<img src=images/star/`.gif alt=Ë«×Ó×ù>');}
else{return('<img src=images/star/_.gif alt=½ğÅ£×ù>');}
}

if (mm=='6'){
if(dd>=22){return('<img src=images/star/a.gif alt=¾ŞĞ·×ù>');}
else{return('<img src=images/star/`.gif alt=Ë«×Ó×ù>');}
}

if (mm=='7'){
if(dd>=23){return('<img src=images/star/b.gif alt=Ê¨×Ó×ù>');}
else{return('<img src=images/star/a.gif alt=¾ŞĞ·×ù>');}
}

if (mm=='8'){
if(dd>=24){return('<img src=images/star/c.gif alt=´¦Å®×ù>');}
else{return('<img src=images/star/b.gif alt=Ê¨×Ó×ù>');}
}

if (mm=='9'){
if(dd>=24){return('<img src=images/star/d.gif alt=Ìì³Ó×ù>');}
else{return('<img src=images/star/c.gif alt=´¦Å®×ù>');}
}

if (mm=='10'){
if(dd>=24){return('<img src=images/star/e.gif alt=ÌìĞ«×ù>');}
else{return('<img src=images/star/d.gif alt=Ìì³Ó×ù>');}
}

if (mm=='11'){
if(dd>=23){return('<img src=images/star/f.gif alt=ÉäÊÖ×ù>');}
else{return('<img src=images/star/e.gif alt=ÌìĞ«×ù>');}
}

if (mm=='12'){
if(dd>=22){return('<img src=images/star/g.gif alt=Ä§ôÉ×ù>');}
else{return('<img src=images/star/f.gif alt=ÉäÊÖ×ù>');}
}

}