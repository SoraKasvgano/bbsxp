
/*****************************************************************************
                                   ????????
*****************************************************************************/

var lunarInfo=new Array(
0x04bd8,0x04ae0,0x0a570,0x054d5,0x0d260,0x0d950,0x16554,0x056a0,0x09ad0,0x055d2,
0x04ae0,0x0a5b6,0x0a4d0,0x0d250,0x1d255,0x0b540,0x0d6a0,0x0ada2,0x095b0,0x14977,
0x04970,0x0a4b0,0x0b4b5,0x06a50,0x06d40,0x1ab54,0x02b60,0x09570,0x052f2,0x04970,
0x06566,0x0d4a0,0x0ea50,0x06e95,0x05ad0,0x02b60,0x186e3,0x092e0,0x1c8d7,0x0c950,
0x0d4a0,0x1d8a6,0x0b550,0x056a0,0x1a5b4,0x025d0,0x092d0,0x0d2b2,0x0a950,0x0b557,
0x06ca0,0x0b550,0x15355,0x04da0,0x0a5d0,0x14573,0x052d0,0x0a9a8,0x0e950,0x06aa0,
0x0aea6,0x0ab50,0x04b60,0x0aae4,0x0a570,0x05260,0x0f263,0x0d950,0x05b57,0x056a0,
0x096d0,0x04dd5,0x04ad0,0x0a4d0,0x0d4d4,0x0d250,0x0d558,0x0b540,0x0b5a0,0x195a6,
0x095b0,0x049b0,0x0a974,0x0a4b0,0x0b27a,0x06a50,0x06d40,0x0af46,0x0ab60,0x09570,
0x04af5,0x04970,0x064b0,0x074a3,0x0ea50,0x06b58,0x055c0,0x0ab60,0x096d5,0x092e0,
0x0c960,0x0d954,0x0d4a0,0x0da50,0x07552,0x056a0,0x0abb7,0x025d0,0x092d0,0x0cab5,
0x0a950,0x0b4a0,0x0baa4,0x0ad50,0x055d9,0x04ba0,0x0a5b0,0x15176,0x052b0,0x0a930,
0x07954,0x06aa0,0x0ad50,0x05b52,0x04b60,0x0a6e6,0x0a4e0,0x0d260,0x0ea65,0x0d530,
0x05aa0,0x076a3,0x096d0,0x04bd7,0x04ad0,0x0a4d0,0x1d0b6,0x0d250,0x0d520,0x0dd45,
0x0b5a0,0x056d0,0x055b2,0x049b0,0x0a577,0x0a4b0,0x0aa50,0x1b255,0x06d20,0x0ada0)

var solarMonth=new Array(31,28,31,30,31,30,31,31,30,31,30,31);
var Gan=new Array("??","??","??","??","??","??","??","??","??","??");
var Zhi=new Array("??","??","??","î","??","??","??","δ","??","??","??","??");
var Animals=new Array("??","ţ","??","??","??","??","??","??","??","??","??","??");
var solarTerm = new Array("С??","????","????","??ˮ","????","????","????","????","????","С??","â??","????","С??","????","????","????","??¶","????","??¶","˪??","????","Сѩ","??ѩ","????")
var sTermInfo = new Array(0,21208,42467,63836,85337,107014,128867,150921,173149,195551,218072,240693,263343,285989,308563,331033,353350,375494,397447,419210,440795,462224,483532,504758)
var nStr1 = new Array('??','һ','??','??','??','??','??','??','??','??','ʮ')
var nStr2 = new Array('??','ʮ','إ','ئ','??')
var monthName = new Array("1 ??","2 ??","3 ??","4 ??","5 ??","6 ??","7 ??","8 ??","9 ??","10 ??","11 ??","12 ??");

//???????? *??ʾ?ż???
var sFtv = new Array(
"0101 Ԫ??",
"0214 ???˽?",
"0308 ??Ů??",
"0401 ???˽?",
"0501 ?Ͷ???",
"0504 ??????",
"0601 ??ͯ??",
"0701 ??????",
"0801 ??????",
"0910 ??ʦ??",
"1001 ??????",
"1120 ??????ͯ??",
"1224 ƽ??ҹ",
"1225 ʥ????")

//ũ?????? *??ʾ?ż???
var lFtv = new Array(
"0101 ????",
"0115 Ԫ????",
"0505 ??????",
"0707 ??Ϧ???˽?",
"0815 ??????",
"0909 ??????",
"1224 С??",
"0100 ??Ϧ")

//ĳ?µĵڼ??????ڼ?
var wFtv = new Array(
"0520 ĸ?׽?",
"0630 ???׽?")

/*****************************************************************************
                                      ???ڼ???
*****************************************************************************/

//====================================== ????ũ?? y??????????
function lYearDays(y) {
   var i, sum = 348
   for(i=0x8000; i>0x8; i>>=1) sum += (lunarInfo[y-1900] & i)? 1: 0
   return(sum+leapDays(y))
}

//====================================== ????ũ?? y?????µ?????
function leapDays(y) {
   if(leapMonth(y))  return((lunarInfo[y-1900] & 0x10000)? 30: 29)
   else return(0)
}

//====================================== ????ũ?? y?????ĸ??? 1-12 , û?򴫻? 0
function leapMonth(y) {
   return(lunarInfo[y-1900] & 0xf)
}

//====================================== ????ũ?? y??m?µ???????
function monthDays(y,m) {
   return( (lunarInfo[y-1900] & (0x10000>>m))? 30: 29 )
}

//====================================== ????ũ??, ????????????, ????ũ??????????
//                                       ???????????? .year .month .day .isLeap .yearCyl .dayCyl .monCyl
function Lunar(objDate) {

   var i, leap=0, temp=0
   var baseDate = new Date(1900,0,31)
   var offset   = (objDate - baseDate)/86400000

   this.dayCyl = offset + 40
   this.monCyl = 14

   for(i=1900; i<2050 && offset>0; i++) {
      temp = lYearDays(i)
      offset -= temp
      this.monCyl += 12
   }

   if(offset<0) {
      offset += temp;
      i--;
      this.monCyl -= 12
   }

   this.year = i
   this.yearCyl = i-1864

   leap = leapMonth(i) //???ĸ???
   this.isLeap = false

   for(i=1; i<13 && offset>0; i++) {
      //????
      if(leap>0 && i==(leap+1) && this.isLeap==false)
         { --i; this.isLeap = true; temp = leapDays(this.year); }
      else
         { temp = monthDays(this.year, i); }

      //????????
      if(this.isLeap==true && i==(leap+1)) this.isLeap = false

      offset -= temp
      if(this.isLeap == false) this.monCyl ++
   }

   if(offset==0 && leap>0 && i==leap+1)
      if(this.isLeap)
         { this.isLeap = false; }
      else
         { this.isLeap = true; --i; --this.monCyl;}

   if(offset<0){ offset += temp; --i; --this.monCyl; }

   this.month = i
   this.day = offset + 1
}

//==============================???ع??? y??ĳm+1?µ?????
function solarDays(y,m) {
   if(m==1)
      return(((y%4 == 0) && (y%100 != 0) || (y%400 == 0))? 29: 28)
   else
      return(solarMonth[m])
}
//============================== ???? offset ???ظ?֧, 0=????
function cyclical(num) {
   return(Gan[num%10]+Zhi[num%12])
}

//============================== ????????
function calElement(sYear,sMonth,sDay,week,lYear,lMonth,lDay,isLeap,cYear,cMonth,cDay) {

      this.isToday    = false;
      //????
      this.sYear      = sYear;
      this.sMonth     = sMonth;
      this.sDay       = sDay;
      this.week       = week;
      //ũ??
      this.lYear      = lYear;
      this.lMonth     = lMonth;
      this.lDay       = lDay;
      this.isLeap     = isLeap;
      //??֧
      this.cYear      = cYear;
      this.cMonth     = cMonth;
      this.cDay       = cDay;

      this.color      = '';

      this.lunarFestival = ''; //ũ??????
      this.solarFestival = ''; //????????
      this.solarTerms    = ''; //????

}

//===== ĳ???ĵ?n??????Ϊ????(??0С??????)
function sTerm(y,n) {
   var offDate = new Date( ( 31556925974.7*(y-1900) + sTermInfo[n]*60000  ) + Date.UTC(1900,0,6,2,5) )
   return(offDate.getUTCDate())
}

//============================== ???????????? (y??,m+1??)
function calendar(y,m) {

   var sDObj, lDObj, lY, lM, lD=1, lL, lX=0, tmp1, tmp2
   var lDPOS = new Array(3)
   var n = 0
   var firstLM = 0

   sDObj = new Date(y,m,1)            //????һ??????

   this.length    = solarDays(y,m)    //????????????
   this.firstWeek = sDObj.getDay()    //????????1?????ڼ?


   for(var i=0;i<this.length;i++) {

      if(lD>lX) {
         sDObj = new Date(y,m,i+1)    //????һ??????
         lDObj = new Lunar(sDObj)     //ũ??
         lY    = lDObj.year           //ũ????
         lM    = lDObj.month          //ũ????
         lD    = lDObj.day            //ũ????
         lL    = lDObj.isLeap         //ũ???Ƿ?????
         lX    = lL? leapDays(lY): monthDays(lY,lM) //ũ??????????һ??

         if(n==0) firstLM = lM
         lDPOS[n++] = i-lD+1
      }

      //sYear,sMonth,sDay,week,
      //lYear,lMonth,lDay,isLeap,
      //cYear,cMonth,cDay
      this[i] = new calElement(y, m+1, i+1, nStr1[(i+this.firstWeek)%7],
                               lY, lM, lD++, lL,
                               cyclical(lDObj.yearCyl) ,cyclical(lDObj.monCyl), cyclical(lDObj.dayCyl++) )


      if((i+this.firstWeek)%7==0)   this[i].color = 'red'  //??????ɫ
      if((i+this.firstWeek)%14==13) this[i].color = 'red'  //???ݶ?????ɫ
   }

   //????
   tmp1=sTerm(y,m*2  )-1
   tmp2=sTerm(y,m*2+1)-1
   this[tmp1].solarTerms = solarTerm[m*2]
   this[tmp2].solarTerms = solarTerm[m*2+1]
   if(m==3) this[tmp1].color = 'red' //??????ɫ

   //????????
   for(i in sFtv)
      if(sFtv[i].match(/^(\d{2})(\d{2})([\s\*])(.+)$/))
         if(Number(RegExp.$1)==(m+1)) {
            this[Number(RegExp.$2)-1].solarFestival += RegExp.$4 + ' '
            if(RegExp.$3=='*') this[Number(RegExp.$2)-1].color = 'red'
         }

   //???ܽ???
   for(i in wFtv)
      if(wFtv[i].match(/^(\d{2})(\d)(\d)([\s\*])(.+)$/))
         if(Number(RegExp.$1)==(m+1)) {
            tmp1=Number(RegExp.$2)
            tmp2=Number(RegExp.$3)
            this[((this.firstWeek>tmp2)?7:0) + 7*(tmp1-1) + tmp2 - this.firstWeek].solarFestival += RegExp.$5 + ' '
         }

   //ũ??????
   for(i in lFtv)
      if(lFtv[i].match(/^(\d{2})(.{2})([\s\*])(.+)$/)) {
         tmp1=Number(RegExp.$1)-firstLM
         if(tmp1==-11) tmp1=1
         if(tmp1 >=0 && tmp1<n) {
            tmp2 = lDPOS[tmp1] + Number(RegExp.$2) -1
            if( tmp2 >= 0 && tmp2<this.length) {
               this[tmp2].lunarFestival += RegExp.$4 + ' '
               if(RegExp.$3=='*') this[tmp2].color = 'red'
            }
         }
      }

  
   //????
   if(y==tY && m==tM) this[tD-1].isToday = true;

}

//====================== ????????
function cDay(d){
   var s;

   switch (d) {
      case 10:
         s = '??ʮ'; break;
      case 20:
         s = '??ʮ'; break;
         break;
      case 30:
         s = '??ʮ'; break;
         break;
      default :
         s = nStr2[Math.floor(d/10)];
         s += nStr1[d%10];
   }
   return(s);
}

///////////////////////////////////////////////////////////////////////////////

var cld;

function drawCld(SY,SM) {
   var i,sD,s,size;
   cld = new calendar(SY,SM);

   

   GZ.innerHTML = 'ũ??' + cyclical(SY-1900+36) + '?? &nbsp;&nbsp;??'+Animals[(SY-4)%12]+'??';


   for(i=0;i<42;i++) {

      sObj=eval('SD'+ i);
      lObj=eval('LD'+ i);


      sD = i - cld.firstWeek;

      if(sD>-1 && sD<cld.length) { //??????
         sObj.innerHTML = "<a href=calendar.asp?menu=show&lookdate=" + SY + "-" + (SM+1) + "-" + (sD+1)+"><font color=000000>"+(sD+1)+"</font></a>";


         if(cld[sD].lDay==1) //??ʾũ????
            lObj.innerHTML = '<b>'+(cld[sD].isLeap?'??':'') + cld[sD].lMonth + '??' + (monthDays(cld[sD].lYear,cld[sD].lMonth)==29?'С':'??')+'</b>';
         else //??ʾũ????
            lObj.innerHTML = cDay(cld[sD].lDay);

         s=cld[sD].lunarFestival;
         if(s.length>0) { //ũ??????
            if(s.length>6) s = s.substr(0, 4)+'??';
            s = s.fontcolor('red');
         }
         else { //????????
            s=cld[sD].solarFestival;
            if(s.length>0) {
               size = (s.charCodeAt(0)>0 && s.charCodeAt(0)<128)?8:4;
               if(s.length>size+2) s = s.substr(0, size)+'??';
               s = s.fontcolor('#13B0F4');
            }
            else { //إ?Ľ???
               s=cld[sD].solarTerms;
               if(s.length>0) s = s.fontcolor('#13B0F4');
            }
         }
         if(s.length>0) lObj.innerHTML = s;

   


      }
      else { //??????
         sObj.innerHTML = '';
         lObj.innerHTML = '';

         
      }
      

   }
   

   
}


function changeCld() {
   var y,m;
   y=CLD.SY.selectedIndex+1900;
   m=CLD.SM.selectedIndex;
   drawCld(y,m);
}

function pushBtm(K) {
   switch (K){
      case 'YU' :
         if(CLD.SY.selectedIndex>0) CLD.SY.selectedIndex--;
         break;
      case 'YD' :
         if(CLD.SY.selectedIndex<149) CLD.SY.selectedIndex++;
         break;
      case 'MU' :
         if(CLD.SM.selectedIndex>0) {
            CLD.SM.selectedIndex--;
         }
         else {
            CLD.SM.selectedIndex=11;
            if(CLD.SY.selectedIndex>0) CLD.SY.selectedIndex--;
         }
         break;
      case 'MD' :
         if(CLD.SM.selectedIndex<11) {
            CLD.SM.selectedIndex++;
         }
         else {
            CLD.SM.selectedIndex=0;
            if(CLD.SY.selectedIndex<149) CLD.SY.selectedIndex++;
         }
         break;
      default :
         CLD.SY.selectedIndex=tY-1900;
         CLD.SM.selectedIndex=tM;
   }
   changeCld();
}



var Today = new Date();
var tY = Today.getFullYear();
var tM = Today.getMonth();
var tD = Today.getDate();

///////////////////////////////////////////////////////////////////////////

function changeTZ() {
   CITY.innerHTML = CLD.TZ.value.substr(6)
   setCookie("TZ",CLD.TZ.selectedIndex)
}

function tick() {
   var today
   today = new Date()
   Clock.innerHTML = today.toLocaleString().replace(/(??|??)/g, "/").replace(/??/, "");
   Clock.innerHTML = TimeAdd(today.toGMTString(), CLD.TZ.value)
   window.setTimeout("tick()", 1000);
}

function setCookie(name, value) {
	var today = new Date()
	var expires = new Date()
	expires.setTime(today.getTime() + 1000*60*60*24*365)
	document.cookie = name + "=" + escape(value)	+ "; expires=" + expires.toGMTString()
}

function getCookie(Name) {
   var search = Name + "="
   if(document.cookie.length > 0) {
      offset = document.cookie.indexOf(search)
      if(offset != -1) {
         offset += search.length
         end = document.cookie.indexOf(";", offset)
         if(end == -1) end = document.cookie.length
         return unescape(document.cookie.substring(offset, end))
      }
      else return ""
   }
}

/////////////////////////////////////////////////////////

function initial() {
   dStyle = detail.style;
   CLD.SY.selectedIndex=tY-1900;
   CLD.SM.selectedIndex=tM;
   drawCld(tY,tM);

   CLD.TZ.selectedIndex=getCookie("TZ");
   changeTZ();
   tick();
}


