
/* 

  ================================================
  PVII TabPanel scripts
  Copyright (c) 2006 Project Seven Development
  www.projectseven.com
  Version: 1.0.3
  ================================================
  
*/

var p7tpa=new Array();
function P7_setTP(){ //v1.0.3 by PVII-www.projectseven.com
 var i,h="<sty"+"le type=\"text/css\">\n";h+=".p7TP_tabs{display: block;}.p7TPcontent div{display:none;}\n";
 h+=".p7TPcontent div div{display:block;}\n";if(document.getElementById){for(i=1;i<11;i++){
 h+="#p7tpc"+i+"_1 {display:block;}\n";}h+="\n<"+"/sty"+"le>";document.write(h);}
}
P7_setTP();
function P7_initTP(){ //v1.0.3 by PVII-www.projectseven.com
 var i,j,tb,tD,tP,tA,pb="p7TP";if(!document.getElementById){return;};p7tpa=arguments;
 for(i=1;i<11;i++){tb=pb+i;tP=document.getElementById(tb);if(tP){tD=tP.getElementsByTagName("DIV");
 if(tD){for(j=0;j<tD.length;j++){if(tD[j].id&&tD[j].id.indexOf("p7tpb")==0){
 tA=tD[j].getElementsByTagName("A");if(tA[0]){tA[0].p7tpn=new Array(i,tD[j].id);
 tA[0].onclick=function(){return P7_TPtrig(this);};}}}}}}
}
function P7_TPtrig(a){ //v1.0.3 by PVII-www.projectseven.com
 var i,tD,tA,tC,c,d,sb,an=p7tpa[1];if(typeof(a)!='object'){c=a.replace("p7tpc","p7tpb");
 d=document.getElementById(c);if(d){a=d.getElementsByTagName("A")[0];}}if(a.p7tpn){
 tD=document.getElementById(a.p7tpn[1]);if(tD){tA=tD.parentNode.getElementsByTagName("A");
 for(i=0;i<tA.length;i++){if(tA[i].p7tpn){sb=tA[i].p7tpn[1].replace("p7tpb","p7tpc");
 tC=document.getElementById(sb);if(tA[i]==a){tA[i].className="down";
 document.getElementById(tA[i].p7tpn[1]).className="down";if(tC){if(an>0){
 tC.style.visibility="hidden";tC.style.display="block";setTimeout("P7_TPanim('"+tC.id+"')",100);
 }else{tC.style.display="block";}}}else{tA[i].className='';
 document.getElementById(tA[i].p7tpn[1]).className='';if(tC){tC.style.display="none";}}}}}}return false;
}
function P7_TPanim(iD){ //v1.0.3 by PVII-www.projectseven.com
 var i,f,tC,g=new Array(),an=p7tpa[1];ob=document.getElementById(iD);tC=ob.parentNode;
 if(!tC.filters){ob.style.opacity="0.10";ob.style.visibility='visible';P7_TPfadeIn(ob.id,0.00);
 return;}f='progid:DXImageTransform.Microsoft.';d=' Duration=1';
 g[0]='Fade';g[1]='Fade';g[2]='Wipe(GradientSize=0.5, wipeStyle=0, motion="forward"'+d+')';
 g[3]='Pixelate(MaxSquare=50,Duration=1,Enabled=false'+d+')';
 g[4]='RandomDissolve('+d+')';g[5]='Iris(irisstyle="SQUARE", motion="in"'+d+')';
 g[6]='Iris(irisstyle="SQUARE", motion="out"'+d+')';g[7]='Iris(irisstyle="CIRCLE", motion="in"'+d+')';
 g[8]='Iris(irisstyle="CIRCLE", motion="out"'+d+')';g[9]='Blinds(direction="up", bands=1'+d+')';
 g[10]='Blinds(direction="down", bands=1'+d+')';g[11]='Blinds(direction="right", bands=1'+d+')';
 g[12]='Blinds(direction="left", bands=1'+d+')';g[13]='Barn(orientation="vertical", motion="in"'+d+')';
 g[14]='Barn(orientation="vertical", motion="out"'+d+')';g[15]='Barn(orientation="horizontal", motion="in"'+d+')'
 g[16]='Barn(orientation="horizontal", motion="out"'+d+')';g[17]='Strips(motion="leftdown"'+d+')';
 g[18]='Strips(motion="leftup"'+d+')';g[19]='Strips(motion="rightdown"'+d+')';
 g[20]='Strips(motion="rightup"'+d+')';g[21]='RadialWipe(wipeStyle="clock"'+d+')';
 g[22]='RadialWipe(wipeStyle="wedge"'+d+')';g[23]='RadialWipe(wipeStyle="radial"'+d+')';
 g[24]='Slide(slideStyle="PUSH", bands=1'+d+')';g[25]='Slide(slideStyle="SWAP", bands=5'+d+')';
 g[26]='Slide(slideStyle="HIDE", bands=1'+d+')';g[27]='Wheel(spokes=4'+d+')';g[28]='Wheel(spokes=16'+d+')';
 an=(an>g.length)?3:an;f+=g[an];tC.style.filter=f;if(tC.filters.length<1){p7tpa[1]=0;ob.style.visibility='visible';
 return;}tC.filters[0].Apply();ob.style.visibility='visible';tC.filters[0].Play();
}
function P7_TPfadeIn(id,op){ //v1.0.3 by PVII-www.projectseven.com
 var d=document.getElementById(id);op+=.05;op=(op>=1)?1:op;d.style.opacity=op;
 if(op<1){setTimeout("P7_TPfadeIn('"+id+"',"+op+")",10);}
}
