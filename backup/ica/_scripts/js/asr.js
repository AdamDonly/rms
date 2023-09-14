var agt=navigator.userAgent.toLowerCase(); 

// *** BROWSER VERSION *** 
var is_major = parseInt(navigator.appVersion); 
var is_minor = parseFloat(navigator.appVersion); 

var is_nav  = ((agt.indexOf('mozilla')!=-1) && (agt.indexOf('spoofer')==-1) 
            && (agt.indexOf('compatible') == -1) && (agt.indexOf('opera')==-1) 
            && (agt.indexOf('webtv')==-1) && (agt.indexOf('gecko')==-1)); 
var is_nav2 = (is_nav && (is_major == 2)); 
var is_nav3 = (is_nav && (is_major == 3)); 
var is_nav4 = (is_nav && (is_major == 4)); 
var is_nav4up = (is_nav && (is_major >= 4)); 
var is_navonly      = (is_nav && ((agt.indexOf(";nav") != -1) || 
                          (agt.indexOf("; nav") != -1)) ); 
var is_nav5 = (is_nav && (is_major == 5)); 
var is_nav5up = (is_nav && (is_major >= 5)); 

var is_ie   = (agt.indexOf("msie") != -1); 
var is_ie3  = (is_ie && (is_major < 4)); 
var is_ie4  = (is_ie && (is_major == 4) && (agt.indexOf("msie 5.0")==-1) ); 
var is_ie4up  = (is_ie  && (is_major >= 4)); 
var is_ie5  = (is_ie && (is_major == 4) && (agt.indexOf("msie 5.0")!=-1) ); 
var is_ie5up  = (is_ie  && !is_ie3 && !is_ie4); 

var is_aol   = (agt.indexOf("aol") != -1); 
var is_aol3  = (is_aol && is_ie3); 
var is_aol4  = (is_aol && is_ie4); 

var is_opera = (agt.indexOf("opera") != -1);
var is_opera7up = (agt.indexOf("opera 7") != -1);
if (is_opera) {is_ie4up=false;}

var is_webtv = (agt.indexOf("webtv") != -1); 
var is_mozilla = (agt.indexOf("gecko") != -1); 
var jvb;

// for scrolling in Mozilla
is_ie4up=is_ie4up || is_mozilla

// for scrolling in Opera
// is_ie4up=is_ie4up || is_opera7up

function nwin(page,title_,w,h){
var win;
win=window.open(page,title_,'resizable=no,menubar=no,status=no,scrollbars=yes,width='+w+',height='+h);
}

function checkBrowser(){
	this.ver=navigator.appVersion;
	this.app=navigator.userAgent.toLowerCase();
	this.dom=document.getElementById?1:0
	this.nav=(navigator.appName == "Netscape" && this.app.indexOf("gecko")<0) ?1:0;
	this.ie5=(this.ver.indexOf("MSIE 5")>-1 && this.dom)?1:0;
	this.ie4=(document.all && !this.dom)?1:0;
	this.nav5=(this.dom && parseInt(this.ver) >= 5) ?1:0;
	this.nav4=(document.layers && !this.dom)?1:0;
	this.op5=(this.app.indexOf("opera") != -1);
	this.mzl=(this.app.indexOf("gecko")>0 && this.dom)
	this.bw=(this.ie5 || this.ie4 || this.nav4 || this.nav5 || this.op5 || this.mzl )
	return this
}
bw=new checkBrowser()

var speed=20;
var num
var loop, timer

function makeObj(obj,nest){
    nest=(!nest) ? '':'document.'+nest+'.'
	this.el=bw.dom?document.getElementById(obj):bw.ie4?document.all[obj]:bw.nav4?eval(nest+'document.'+obj):0;
  	this.css=bw.dom?document.getElementById(obj).style:bw.ie4?document.all[obj].style:bw.nav4?eval(nest+'document.'+obj):0;
	this.scrollHeight=bw.nav4?this.css.document.height:this.el.offsetHeight
	this.clipHeight=bw.nav4?this.css.clip.height:this.el.offsetHeight
	this.up=goUp;this.down=goDown;
	this.moveIt=moveIt; this.x; this.y;
    this.obj = obj + "Object"
    eval(this.obj + "=this")
    return this
}

function moveIt(x,y){
	this.x=x;this.y=y
	this.css.left=this.x
	this.css.top=this.y
}

function goDown(move,cnum){
var mv=0
	if(this.y>-this.scrollHeight+eval('oCont'+cnum+'.clipHeight')){
		if(this.y+this.scrollHeight-eval('oCont'+cnum+'.clipHeight')-move-0<0)
		{mv=this.y+this.scrollHeight-eval('oCont'+cnum+'.clipHeight')-move-0}
		else
		{mv=0}
		
		this.moveIt(0,this.y-move-mv)
			if(loop) setTimeout(this.obj+".down("+eval(move+mv)+","+cnum+")",speed)
	}
}

function goUp(move){
var mv=0
	if(this.y<0){
		if (this.y-move>0) {mv=this.y-move}
		this.moveIt(0,this.y-move-mv)
		if(loop) setTimeout(this.obj+".up("+eval(move+mv)+")",speed)
	}
}

function scroll(speed,num, idl){
	if(loaded){
		loop=true;
		if(speed>0) eval('oContScroll'+num+'.down(speed,'+num+')')
		else eval('oContScroll'+num+'.up(speed)')
	}
}

function noScroll(){
	loop=false;
	if(timer) clearTimeout(timer)
}

var loaded;
function scrollInit(lr1,lr2,lr3){
   if (is_ie4up || is_nav4up)
   {
	if (lr1==1)
	{
	oCont1=new makeObj('Div1')
	oContScroll1=new makeObj('DivText1','Div1')
	oContScroll1.moveIt(0,0)
	oCont1.css.visibility='visible'
	}

	if (lr2==1)
	{
	oCont2=new makeObj('Div2')
	oContScroll2=new makeObj('DivText2','Div2')
	oContScroll2.moveIt(0,0)
	oCont2.css.visibility='visible'
	}

	if (lr3==1)
	{
	oCont3=new makeObj('Div3')
	oContScroll3=new makeObj('DivText3','Div3')
	oContScroll3.moveIt(0,0)
	oCont3.css.visibility='visible'
	}
	loaded=true;
   }

}


function init(){
  if (navigator.appName == "Netscape") 
   {
    layerStyleRef="layer.";
    layerRef="document.layers";
    styleSwitch="";
   }
  else
   {
    layerStyleRef="layer.style.";
    layerRef="document.all";
    styleSwitch=".style";
   }
}

function showLayer(layerName)
{
  eval(layerRef+'["'+layerName+'"]'+styleSwitch+'.visibility="visible"');
}
        
function hideLayer(layerName)
{
  eval(layerRef+'["'+layerName+'"]'+styleSwitch+'.visibility="hidden"');
}

function StrToCurrency(strInput)
{
var strTemp=strInput.toString();
var cp=strTemp.indexOf('.');
if (cp>0)
{strTemp=strTemp+'0'; return strTemp.substr(0,cp+2);}
else
return strTemp+'.0';
}

