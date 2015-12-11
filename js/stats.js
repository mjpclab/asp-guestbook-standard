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

function getsysname()
{
	var buffer;
	var strUserAgent=navigator.userAgent;

	//for Mozilla
	if ((/Win95/i).test(strUserAgent)) return 'Windows 95';
	else if ((/Win98/i).test(strUserAgent)) return 'Windows 98';
	else if ((/WinNT/i).test(strUserAgent)) return 'Windows NT';

	//For Opera
	else if ((/Windows ME/i).test(strUserAgent)) return 'Windows ME';

	//for IE & Opera
	else if ((/Windows 95/i).test(strUserAgent)) return 'Windows 95';
	else if ((/Win 9x 4\.90/i).test(strUserAgent)) return 'Windows ME';
	else if ((/Windows 98/i).test(strUserAgent)) return 'Windows 98';

	//for IE & Opera & Mozilla
	else if(buffer=/OS (?:(\d+_\d+)\w+) like Mac OS X/i.exec(strUserAgent)) return 'iOS ' + buffer[1].replace(/_/g,'.');
	else if(buffer=/Android \d+\.\d+/.exec(strUserAgent)) return buffer[0];
	else if(buffer=/SymbianOS\/[\w\.]+/.exec(strUserAgent)) return buffer[0].replace('/',' ');
	else if(buffer=/Windows Phone(?: OS)? [\w\.]+/.exec(strUserAgent)) return buffer[0].replace(' OS','');
	else if ((/BlackBerry/i).test(strUserAgent)) return 'BlackBerry';
	else if ((/Windows NT 5\.0/i).test(strUserAgent)) return 'Windows 2000';
	else if ((/Windows NT 5\.0/i).test(strUserAgent)) return 'Windows 2000';
	else if ((/Windows NT 5\.1/i).test(strUserAgent)) return 'Windows XP';
	else if ((/Windows NT 5\.2/i).test(strUserAgent)) return 'Windows Server2003';
	else if ((/Windows NT 6\.0/i).test(strUserAgent)) return 'Windows Vista/Server2008';
	else if ((/Windows NT 6\.1/i).test(strUserAgent)) return 'Windows 7/Server2008R2';
	else if ((/Windows NT 6\.2/i).test(strUserAgent)) return 'Windows 8/Server2012';
	else if ((/Windows NT 6\.3/i).test(strUserAgent)) return 'Windows 8.1/Server2012R2';
	else if ((/Windows NT 10\.0/i).test(strUserAgent)) return 'Windows 10';
	else if ((/Windows NT/).test(strUserAgent)) return 'Windows NT';
	else if ((/Linux/i).test(strUserAgent)) return 'Linux';
	else if ((/NetWare/i).test(strUserAgent)) return 'NetWare';
	else if ((/Solaris/i).test(strUserAgent)) return 'Solaris';
	else if ((/BSD/i).test(strUserAgent)) return 'BSD';
	else if ((/Unix/i).test(strUserAgent)) return 'Unix';
	else if(buffer=/Mac OS X \d+_\d+/i.exec(strUserAgent)) return buffer[0].replace(/_/g,'.');
	else if ((/Mac\s?OS/i).test(strUserAgent)) return 'Mac OS';
	else return '';
}

function getbrowsername()
{
	var strUserAgent=navigator.userAgent;

	if ((/Opera/i).test(strUserAgent)) return 'Opera';
	else if ((/\sOPR\//i).test(strUserAgent)) return 'Opera Webkit';
	else if ((/Trident\/7\.0/i).test(strUserAgent)) return 'MSIE 11';
	else if (navigator.appName=='Microsoft Internet Explorer')
	{
		return strUserAgent.substring(strUserAgent.indexOf('(')+1, strUserAgent.indexOf(')')).split(/;\s*/,2)[1];
	}
	else if ((/Edge/i).test(strUserAgent)) return 'Edge';
	else if ((/Chrome/i).test(strUserAgent)) return 'Chrome';
	else if ((/Safari/i).test(strUserAgent)) return 'Safari';
	else if ((/NetScape/i).test(strUserAgent)) return 'NetScape';
	else if ((/Firefox/i).test(strUserAgent)) return 'Firefox';
	else if ((/Mozilla/i).test(strUserAgent)) return 'Mozilla';
	else if (navigator.appName) return navigator.appName;
	else if (navigator.appCodeName) return navigator.appCodeName;
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

function createXmlHttp() {
	if (window.XMLHttpRequest) {
		return new XMLHttpRequest();
	}
	else if (window.ActiveXObject) {
		return new ActiveXObject('Microsoft.XMLHTTP');
	}
}

function xmlHttpHandler()
{
	if(this.readyState===4 && this.status===200)
	{
		this.abort();
	}
}

function RequestScript(url) {
	var xmlHttp=createXmlHttp();

	if(xmlHttp)
	{
		xmlHttp.onreadystatechange=xmlHttpHandler;
		xmlHttp.open('GET',url);
		xmlHttp.send(null);
	}
	else
	{
		var script=document.createElement('script');
		script.type='text/javascript';
		script.src=url;
		script.defer='defer';
		script.async='async';
		document.body.appendChild(script);
	}
}