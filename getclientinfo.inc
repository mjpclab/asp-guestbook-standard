<%if session("gotclientinfo")<>true then%>
<script type="text/javascript">
//<![CDATA[

<!-- #include file="xmlhttp.inc" -->

function trim(s)
{
	s=s.replace(/^\s*/,'');
	s=s.replace(/\s*$/,'');
	return s;
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
		//for Mozilla
		if ((/Win95/i).test(navigator.userAgent)) return 'Windows 95';
		else if ((/Win98/i).test(navigator.userAgent)) return 'Windows 98';
		else if ((/WinNT/i).test(navigator.userAgent)) return 'Windows NT';

		//For Opera
		else if ((/Windows ME/i).test(navigator.userAgent)) return 'Windows ME';
	
		//for IE & Opera
		else if ((/Windows 95/i).test(navigator.userAgent)) return 'Windows 95';
		else if ((/Win 9x 4.90/i).test(navigator.userAgent)) return 'Windows ME';
		else if ((/Windows 98/i).test(navigator.userAgent)) return 'Windows 98';

		//for IE & Opera & Mozilla
		else if ((/Windows NT 5.0/i).test(navigator.userAgent)) return 'Windows 2000';
		else if ((/Windows NT 5.1/i).test(navigator.userAgent)) return 'Windows XP';
		else if ((/Windows NT 5.2/i).test(navigator.userAgent)) return 'Windows 2003';
		else if ((/Windows NT 6.0/i).test(navigator.userAgent)) return 'Windows Vista';
		else if ((/Windows NT/).test(navigator.userAgent)) return 'Windows NT';
		else if ((/Linux/i).test(navigator.userAgent)) return 'Linux';
		else if ((/NetWare/i).test(navigator.userAgent)) return 'NetWare';
		else if ((/Solaris/i).test(navigator.userAgent)) return 'Solaris';
		else if ((/FreeBSD/i).test(navigator.userAgent)) return 'FreeBSD';
		else if ((/Unix/i).test(navigator.userAgent)) return 'Unix';
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
		else if (navigator.appName=='Microsoft Internet Explorer')
		{
			var bname=navigator.appVersion.match(/\(.*\)/);
			if (bname) bname=bname[0].substring(1,bname[0].length-1).split(';')[1];
			bname=trim(bname);
			return bname;
		}
		else if ((/NetScape/i).test(navigator.userAgent)) return 'NetScape';
		else if ((/Mozilla.+Firefox/i).test(navigator.userAgent)) return 'Mozilla Firefox';
		else if ((/Mozilla/i).test(navigator.userAgent)) return 'Mozilla';
		else if ((/Safari/i).test(navigator.userAgent)) return 'Safari';
		else if (navigator.appName) return navigator.appName;
		else if (navigator.appCodeName) return navigator.appCodeName;
		else return '';
	}
	else return '';
}

function getscreenwidth() {if (screen.width) return screen.width; else return 0;}

function getscreenheight() {if (screen.height) return screen.height; else return 0;}

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
var url=window.location.href.substring(0,window.location.href.lastIndexOf('/')+1);
url+='saveclientinfo.asp?sys=' +getsysname()+ '&brow=' +getbrowsername()+ '&sw=' +getscreenwidth()+ '&sh=' +getscreenheight()+ '&src=' +getsourceaddr()+ '&fsrc=' +escape(document.referrer);

if(xmlHttp)
{
	xmlHttp.onreadystatechange=xmlHttpHandler;
	xmlHttp.open('GET',url);
	xmlHttp.send(null);
}
else
{
	var spt,output;
	spt='script';
	output='<' +spt+ ' type="text/javascript" src="' +url+ '" defer="defer"><\/script>';
	document.write(output);
}

//]]>
</script>
<%end if%>