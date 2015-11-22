function createXmlHttp() {
	if (window.XMLHttpRequest) {
		return new XMLHttpRequest();
	}
	else if (window.ActiveXObject) {
		return new ActiveXObject('Microsoft.XMLHTTP');
	}
}

function clearChildren(element) {
	if (element.childNodes) {
		while (element.hasChildNodes()) element.removeChild(element.lastChild);
	}
}

function setPureText(element, text) {
	clearChildren(element);
	element.appendChild(document.createTextNode(text));
}