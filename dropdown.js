// JavaScript Document

<!--
/***********************************************
* AnyLink Drop Down Menu- © Dynamic Drive (www.dynamicdrive.com)
* This notice MUST stay intact for legal use
* Visit http://www.dynamicdrive.com/ for full source code
***********************************************/

var menu3=new Array()
menu3[0]='<a href="grower_report.asp">Grower</a>'
menu3[1]='<a href="grower_report_PURS.asp">PURS</a>'
menu3[2]='<a href="sprayrecords_list.asp">History</a>'

var menu4=new Array()
menu4[0]='<a href="packers_list.asp">Packers</a>'
menu4[1]='<a href="packerusers_list.asp"><i>Packer Users</i></a>'
menu4[2]='<a href="growers_list.asp">Growers</a>'
menu4[3]='<a href="growerusers_list.asp"><i>Grower Users</i></a>'
menu4[4]='<a href="growerunit.asp"><i>Grower Orgs</i></a>'
menu4[5]='<a href="methods_list.asp">Methods</a>'
menu4[6]='<a href="spraylist_list.asp">Spray List</a>'
menu4[7]='<a href="stages_list.asp">Stages</a>'
menu4[8]='<a href="targets_list.asp">Targets</a>'
menu4[9]='<a href="crops_list.asp">Crops/Varieties</a>'
menu4[10]='<a href="units_list.asp">Units</a>'
menu4[11]='<a href="sprayyears_list.asp">Spray Years</a>'
menu4[12]='<a href="administrators_list.asp">Admins</a>'

var menu5=new Array()
menu5[0]='<a href="SprayProgramInstructionsGROWERS.pdf" target=_blank>Growers Manual</a>'
menu5[1]='<a href="SprayProgramInstructionsPACKERS.pdf" target=_blank>Packers Manual</a>'

var menu6=new Array()
menu6[0]='<a href="grower_report.asp">Grower</a>'
menu6[1]='<a href="grower_report_PURS.asp">PURS</a>'
menu6[2]='<a href="spraylist_list.asp">Spray List</a>'

var menu7=new Array()
menu7[0]='<a href="GrowerCrop.asp">Crops/Varieties</a>'
menu7[1]='<a href="GrowerLocations.asp">Locations</a>'
menu7[2]='<a href="GrowerSupervisors.asp">Supervisors</a>'
menu7[3]='<a href="GrowerApplicators.asp">Applicators</a>'
menu7[4]='<a href="GrowerSuppliers.asp">Suppliers</a>'
menu7[5]='<a href="GrowerReferrers.asp">Referrals</a>'

var menu8=new Array()
menu8[0]='<a href="grower_report.asp">Grower</a>'
menu8[1]='<a href="spraylist_list.asp">Spray List</a>'
menu8[2]='<a href="sprayrecords_list.asp">History</a>'

var menu9=new Array()
menu9[0]='<a href="growers_list.asp">Growers</a>'
menu9[1]='<a href="growerusers_list.asp"><i>Grower Users</i></a>'
menu9[2]='<a href="packerusers_list.asp"><i>Packer Users</i></a>'
menu9[3]='<a href="spraylist_list.asp">Spray List</a>'

var menuwidth='450px' //default menu width
var menubgcolor='#eeeeee'  //menu bgcolor
var disappeardelay=250  //menu disappear speed onMouseout (in miliseconds)
var hidemenu_onclick="yes" //hide menu when user clicks within menu?

/////No further editting needed

var ie4=document.all
var ns6=document.getElementById&&!document.all

if (ie4||ns6)
document.write('<div id="dropmenudiv" style="visibility:hidden;width:'+menuwidth+';background-color:'+menubgcolor+'" onMouseover="clearhidemenu()" onMouseout="dynamichide(event)"></div>')

function getposOffset(what, offsettype){
var totaloffset=(offsettype=="left")? what.offsetLeft : what.offsetTop;
var parentEl=what.offsetParent;
while (parentEl!=null){
totaloffset=(offsettype=="left")? totaloffset+parentEl.offsetLeft : totaloffset+parentEl.offsetTop;
parentEl=parentEl.offsetParent;
}
return totaloffset;
}


function showhide(obj, e, visible, hidden, menuwidth){
if (ie4||ns6)
dropmenuobj.style.left=dropmenuobj.style.top="-500px"
if (menuwidth!=""){
dropmenuobj.widthobj=dropmenuobj.style
dropmenuobj.widthobj.width=menuwidth
}
if (e.type=="click" && obj.visibility==hidden || e.type=="mouseover")
obj.visibility=visible
else if (e.type=="click")
obj.visibility=hidden
}

function iecompattest(){
return (document.compatMode && document.compatMode!="BackCompat")? document.documentElement : document.body
}

function clearbrowseredge(obj, whichedge){
var edgeoffset=0
if (whichedge=="rightedge"){
var windowedge=ie4 && !window.opera? iecompattest().scrollLeft+iecompattest().clientWidth-15 : window.pageXOffset+window.innerWidth-15
dropmenuobj.contentmeasure=dropmenuobj.offsetWidth
if (windowedge-dropmenuobj.x < dropmenuobj.contentmeasure)
edgeoffset=dropmenuobj.contentmeasure-obj.offsetWidth
}
else{
var topedge=ie4 && !window.opera? iecompattest().scrollTop : window.pageYOffset
var windowedge=ie4 && !window.opera? iecompattest().scrollTop+iecompattest().clientHeight-15 : window.pageYOffset+window.innerHeight-18
dropmenuobj.contentmeasure=dropmenuobj.offsetHeight
if (windowedge-dropmenuobj.y < dropmenuobj.contentmeasure){ //move up?
edgeoffset=dropmenuobj.contentmeasure+obj.offsetHeight
if ((dropmenuobj.y-topedge)<dropmenuobj.contentmeasure) //up no good either?
edgeoffset=dropmenuobj.y+obj.offsetHeight-topedge
}
}
return edgeoffset
}

function populatemenu(what){
if (ie4||ns6)
dropmenuobj.innerHTML=what.join("")
}


function dropdownmenu(obj, e, menucontents, menuwidth){
if (window.event) event.cancelBubble=true
else if (e.stopPropagation) e.stopPropagation()
clearhidemenu()
dropmenuobj=document.getElementById? document.getElementById("dropmenudiv") : dropmenudiv
populatemenu(menucontents)

if (ie4||ns6){
showhide(dropmenuobj.style, e, "visible", "hidden", menuwidth)
dropmenuobj.x=getposOffset(obj, "left")
dropmenuobj.y=getposOffset(obj, "top")
dropmenuobj.style.left=dropmenuobj.x-clearbrowseredge(obj, "rightedge")+"px"
dropmenuobj.style.top=dropmenuobj.y-clearbrowseredge(obj, "bottomedge")+obj.offsetHeight+"px"
}

return clickreturnvalue()
}

function clickreturnvalue(){
if (ie4||ns6) return false
else return true
}

function contains_ns6(a, b) {
while (b.parentNode)
if ((b = b.parentNode) == a)
return true;
return false;
}

function dynamichide(e){
if (ie4&&!dropmenuobj.contains(e.toElement))
delayhidemenu()
else if (ns6&&e.currentTarget!= e.relatedTarget&& !contains_ns6(e.currentTarget, e.relatedTarget))
delayhidemenu()
}

function hidemenu(e){
if (typeof dropmenuobj!="undefined"){
if (ie4||ns6)
dropmenuobj.style.visibility="hidden"
}
}

function delayhidemenu(){
if (ie4||ns6)
delayhide=setTimeout("hidemenu()",disappeardelay)
}

function clearhidemenu(){
if (typeof delayhide!="undefined")
clearTimeout(delayhide)
}

if (hidemenu_onclick=="yes")
document.onclick=hidemenu

//-->


