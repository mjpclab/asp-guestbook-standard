<%Response.ContentType="application/x-javascript"%>
<%if session("gotclientinfo")<>true then%>
<!-- #include file="xmlhttp.inc" -->

function trim(s)
{
	if(s.trim) {
		return s.trim();
	}
	else {
		s=s.replace(/^\s*/,'');
		s=s.replace(/\s*$/,'');
		return s;
	}
}

function xmlHttpHandler()
{
	if(xmlHttp && xmlHttp.readyState==4 && xmlHttp.status==200)
	{
		xmlHttp.abort();
		xmlHttp=null;
	}
}

function getsysname()
{
	if (navigator.userAgent)
	{
		var buffer;

		//for Mozilla
		if ((/Win95/i).test(navigator.userAgent)) return 'Windows 95';
		else if ((/Win98/i).test(navigator.userAgent)) return 'Windows 98';
		else if ((/WinNT/i).test(navigator.userAgent)) return 'Windows NT';

		//For Opera
		else if ((/Windows ME/i).test(navigator.userAgent)) return 'Windows ME';
	
		//for IE & Opera
		else if ((/Windows 95/i).test(navigator.userAgent)) return 'Windows 95';
		else if ((/Win 9x 4\.90/i).test(navigator.userAgent)) return 'Windows ME';
		else if ((/Windows 98/i).test(navigator.userAgent)) return 'Windows 98';

		//for IE & Opera & Mozilla
		else if(buffer=/OS (?:(\d+_\d+)\w+) like Mac OS X/i.exec(navigator.userAgent)) return 'iOS ' + buffer[1].replace(/_/g,'.');
		else if(buffer=/Android \d+\.\d+/.exec(navigator.userAgent)) return buffer[0];
		else if(buffer=/SymbianOS\/[\w\.]+/.exec(navigator.userAgent)) return buffer[0].replace('/',' ');
		else if(buffer=/Windows Phone(?: OS)? [\w\.]+/.exec(navigator.userAgent)) return buffer[0].replace(' OS','');
		else if ((/BlackBerry/i).test(navigator.userAgent)) return 'BlackBerry';
		else if ((/Windows NT 5\.0/i).test(navigator.userAgent)) return 'Windows 2000';
		else if ((/Windows NT 5\.0/i).test(navigator.userAgent)) return 'Windows 2000';
		else if ((/Windows NT 5\.1/i).test(navigator.userAgent)) return 'Windows XP';
		else if ((/Windows NT 5\.2/i).test(navigator.userAgent)) return 'Windows Server2003';
		else if ((/Windows NT 6\.0/i).test(navigator.userAgent)) return 'Windows Vista/Server2008';
		else if ((/Windows NT 6\.1/i).test(navigator.userAgent)) return 'Windows 7/Server2008R2';
		else if ((/Windows NT 6\.2/i).test(navigator.userAgent)) return 'Windows 8/Server2012';
		else if ((/Windows NT 6\.3/i).test(navigator.userAgent)) return 'Windows 8.1/Server2012R2';
		else if ((/Windows NT 10\.0/i).test(navigator.userAgent)) return 'Windows 10';
		else if ((/Windows NT/).test(navigator.userAgent)) return 'Windows NT';
		else if ((/Linux/i).test(navigator.userAgent)) return 'Linux';
		else if ((/NetWare/i).test(navigator.userAgent)) return 'NetWare';
		else if ((/Solaris/i).test(navigator.userAgent)) return 'Solaris';
		else if ((/BSD/i).test(navigator.userAgent)) return 'BSD';
		else if ((/Unix/i).test(navigator.userAgent)) return 'Unix';
		else if(buffer=/Mac OS X \d+_\d+/i.exec(navigator.userAgent)) return buffer[0].replace(/_/g,'.');
		else if ((/Mac\s?OS/i).test(navigator.userAgent)) return 'Mac OS';
		else return '';
	}
	else return '';
}

function getbrowsername()
{
	if (navigator.userAgent)
	{
		if ((/Opera/i).test(navigator.userAgent)) return 'Opera';
		else if ((/Trident\/7\.0/i).test(navigator.userAgent)) return 'MSIE 11';
		else if (navigator.appName=='Microsoft Internet Explorer')
		{
			var bname=navigator.appVersion.match(/\(.*\)/);
			if (bname) bname=bname[0].substring(1,bname[0].length-1).split(';')[1];
			bname=trim(bname);
			return bname;
		}
		else if ((/Edge/i).test(navigator.userAgent)) return 'Edge';
		else if ((/Chrome/i).test(navigator.userAgent)) return 'Chrome';
		else if ((/Safari/i).test(navigator.userAgent)) return 'Safari';
		else if ((/NetScape/i).test(navigator.userAgent)) return 'NetScape';
		else if ((/Firefox/i).test(navigator.userAgent)) return 'Firefox';
		else if ((/Mozilla/i).test(navigator.userAgent)) return 'Mozilla';
		else if (navigator.appName) return navigator.appName;
		else if (navigator.appCodeName) return navigator.appCodeName;
		else return '';
	}
	else return '';
}

function getscreenwidth() {return screen.width || 0;}

function getscreenheight() {return screen.height || 0;}

function getsourceaddr()
{
	if (document && document.referrer)
	{
		var sourceaddr;
		sourceaddr=trim(document.referrer);
		sourceaddr=sourceaddr.replace(/^([\w\d]+\:\/\/)?/,'')
		if (sourceaddr.indexOf('/')!=-1) sourceaddr=sourceaddr.substring(0,sourceaddr.indexOf('/'));
		//sourceaddr=sourceaddr.replace(/\:\d+$/,'');
		sourceaddr=trim(sourceaddr);
		return sourceaddr;
	}
	else return '';
}

var xmlHttp=createXmlHttp();
var url=location.href.substring(0,location.href.lastIndexOf('/')+1);
url+='saveclientinfo.asp?sys=' + encodeURIComponent(getsysname()) +
	'&brow=' + encodeURIComponent(getbrowsername()) +
	'&sw=' + encodeURIComponent(getscreenwidth()) +
	'&sh=' + encodeURIComponent(getscreenheight()) +
	'&src=' + encodeURIComponent(getsourceaddr()) +
	'&fsrc=' + encodeURIComponent(document.referrer);

if(xmlHttp)
{
	xmlHttp.onreadystatechange=xmlHttpHandler;
	xmlHttp.open('GET',url);
	xmlHttp.send(null);
}
else
{
	var script=document.createElement('script');
	script.type="text/javascript";
	script.src=url;
	script.defer="defer";
	script.async="async";
	document.body.appendChild(script);
}
<%end if%>