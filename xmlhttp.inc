function createXmlHttp()
{
	var obj=null;
	if(window.XMLHttpRequest)
	{
		obj=new XMLHttpRequest();
	}
	else if(window.ActiveXObject)
	{
		obj=new ActiveXObject("Microsoft.XMLHTTP");
	}
	return obj;
}

function clearChildren(element)
{
	if(element.childNodes)
		while(element.hasChildNodes()) element.removeChild(element.lastChild);
}

function setPureText(element,text)
{
	clearChildren(element);
	if(document && element.appendChild) element.appendChild(document.createTextNode(text));
}