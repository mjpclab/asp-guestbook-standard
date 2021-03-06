function toArray(arrayObj) {
	var result;
	try{
		result=Array.prototype.slice.call(arrayObj);
	}
	catch(ex) {
		result=arrayObj;
	}

	return result;
}

function storeRange()
{
	if (document.selection)
		this.selRange=document.selection.createRange().duplicate();
}

function sendUbbCode(tagOpen)
{
	var tarea=document.getElementById('icontent') ||
		document.getElementById('rcontent') ||
		document.getElementById('econtent') ||
		document.getElementById('abulletin');
	if (tarea)
	{
		var tagBegin = '[' + tagOpen + ']';
		var tagEnd = tagOpen.match(/\w+/);
		tagEnd = tagEnd[0] || '';
		tagEnd = '[/' + tagEnd + ']';

		if (typeof tarea.selectionStart !== 'undefined') {
			var selectedText = tarea.value.slice(tarea.selectionStart, tarea.selectionEnd);
			var insertText = tagBegin + selectedText + tagEnd;
			var prevSelectionStart = tarea.selectionStart;
			tarea.value = tarea.value.slice(0, tarea.selectionStart) + insertText + tarea.value.slice(tarea.selectionEnd);
			tarea.selectionStart = tarea.selectionEnd = prevSelectionStart + insertText.length;
			tarea.focus();
		}
		else if (document.selection && typeof tarea.selRange!=='undefined') {
			var selectedText = tarea.selRange.text;
			var insertText = tagBegin + selectedText + tagEnd;
			tarea.selRange.text = insertText;
			tarea.focus();
		}
		else {
			tarea.value += tagBegin + tagEnd;
		}

		if(tarea.form) {
			if(tarea.form.ubb1) {
				tarea.form.ubb1.checked=true;
			}
			else if (tarea.form.ubb2)
			{
				tarea.form.ubb2.checked=true;
			}
		}
	}
}

function sendUbbFace()
{
	var tarea=document.getElementById('icontent') ||
		document.getElementById('rcontent') ||
		document.getElementById('econtent') ||
		document.getElementById('abulletin');
	if (tarea)
	{
		var insertText=this.title;
		if (typeof tarea.selectionStart !== 'undefined') {
			var prevSelectionEnd = tarea.selectionEnd;
			tarea.value = tarea.value.slice(0, tarea.selectionEnd) + insertText + tarea.value.slice(tarea.selectionEnd);
			tarea.selectionStart = tarea.selectionEnd = prevSelectionEnd + insertText.length;
			tarea.focus();
		}
		else if (document.selection && typeof tarea.selRange!=='undefined') {
			tarea.selRange.text += insertText;
			tarea.focus();
		}
		else {
			tarea.value += insertText;
		}

		if(tarea.form) {
			if(tarea.form.ubb1) {
				tarea.form.ubb1.checked=true;
			}
			else if (tarea.form.ubb2)
			{
				tarea.form.ubb2.checked=true;
			}
		}
	}
}


if(document.selection) {
	var tarea=document.getElementById('icontent') ||
		document.getElementById('rcontent') ||
		document.getElementById('econtent') ||
		document.getElementById('abulletin');
	if (tarea)
	{
		tarea.onselect=storeRange;
		tarea.onkeyup=storeRange;
		tarea.onclick=storeRange;
	}
}


var faces = document.getElementsByClassName ?
	document.getElementsByClassName('ubbface') :
	document.getElementsByName('ubbface');
faces = toArray(faces);
for (var i = 0, len = faces.length; i < len; i++) {
	var face = faces[i];
	face.onclick = sendUbbFace;
}


///// Color Palatte start
var colors = [
	[
		'#000000', '#000000', '#000000', '#003300', '#006600', '#009900', '#00CC00', '#00FF00', '#330000', '#333300',
		'#336600', '#339900', '#33CC00', '#33FF00', '#660000', '#663300', '#666600', '#669900', '#66CC00', '#66FF00'
	],
	[
		'#333333', '#000000', '#000033', '#003333', '#006633', '#009933', '#00CC33', '#00FF33', '#330033', '#333333',
		'#336633', '#339933', '#33CC33', '#33FF33', '#660033', '#663333', '#666633', '#669933', '#66CC33', '#66FF33'
	],
	[
		'#666666', '#000000', '#000066', '#003366', '#006666', '#009966', '#00CC66', '#00FF66', '#330066', '#333366',
		'#336666', '#339966', '#33CC66', '#33FF66', '#660066', '#663366', '#666666', '#669966', '#66CC66', '#66FF66'
	],
	[
		'#999999', '#000000', '#000099', '#003399', '#006699', '#009999', '#00CC99', '#00FF99', '#330099', '#333399',
		'#336699', '#339999', '#33CC99', '#33FF99', '#660099', '#663399', '#666699', '#669999', '#66CC99', '#66FF99'
	],
	[
		'#CCCCCC', '#000000', '#0000CC', '#0033CC', '#0066CC', '#0099CC', '#00CCCC', '#00FFCC', '#3300CC', '#3333CC',
		'#3366CC', '#3399CC', '#33CCCC', '#33FFCC', '#6600CC', '#6633CC', '#6666CC', '#6699CC', '#66CCCC', '#66FFCC'
	],
	[
		'#FFFFFF', '#000000', '#0000FF', '#0033FF', '#0066FF', '#0099FF', '#00CCFF', '#00FFFF', '#3300FF', '#3333FF',
		'#3366FF', '#3399FF', '#33CCFF', '#33FFFF', '#6600FF', '#6633FF', '#6666FF', '#6699FF', '#66CCFF', '#66FFFF'
	],
	[
		'#FF0000', '#000000', '#990000', '#993300', '#996600', '#999900', '#99CC00', '#99FF00', '#CC0000', '#CC3300',
		'#CC6600', '#CC9900', '#CCCC00', '#CCFF00', '#FF0000', '#FF3300', '#FF6600', '#FF9900', '#FFCC00', '#FFFF00'
	],
	[
		'#00FF00', '#000000', '#990033', '#993333', '#996633', '#999933', '#99CC33', '#99FF33', '#CC0033', '#CC3333',
		'#CC6633', '#CC9933', '#CCCC33', '#CCFF33', '#FF0033', '#FF3333', '#FF6633', '#FF9933', '#FFCC33', '#FFFF33'
	],
	[
		'#0000FF', '#000000', '#990066', '#993366', '#996666', '#999966', '#99CC66', '#99FF66', '#CC0066', '#CC3366',
		'#CC6666', '#CC9966', '#CCCC66', '#CCFF66', '#FF0066', '#FF3366', '#FF6666', '#FF9966', '#FFCC66', '#FFFF66'
	],
	[
		'#FFFF00', '#000000', '#990099', '#993399', '#996699', '#999999', '#99CC99', '#99FF99', '#CC0099', '#CC3399',
		'#CC6699', '#CC9999', '#CCCC99', '#CCFF99', '#FF0099', '#FF3399', '#FF6699', '#FF9999', '#FFCC99', '#FFFF99'
	],
	[
		'#00FFFF', '#000000', '#9900CC', '#9933CC', '#9966CC', '#9999CC', '#99CCCC', '#99FFCC', '#CC00CC', '#CC33CC',
		'#CC66CC', '#CC99CC', '#CCCCCC', '#CCFFCC', '#FF00CC', '#FF33CC', '#FF66CC', '#FF99CC', '#FFCCCC', '#FFFFCC'
	],
	[
		'#FF00FF', '#000000', '#9900FF', '#9933FF', '#9966FF', '#9999FF', '#99CCFF', '#99FFFF', '#CC00FF', '#CC33FF',
		'#CC66FF', '#CC99FF', '#CCCCFF', '#CCFFFF', '#FF00FF', '#FF33FF', '#FF66FF', '#FF99FF', '#FFCCFF', '#FFFFFF'
	]
];

var divPreview = document.getElementById('colorsPreview');
var divNumber = document.getElementById('colorsNumber');
function hoverColorSquare() {
	if (divPreview) {
		divPreview.style.backgroundColor = this.style.backgroundColor;
	}
	if (divNumber && divNumber.firstChild) {
		divNumber.firstChild.nodeValue = this.title;
	}
}

function toggleColorPalette(sender,e)
{
	if(sender && e)
	{
		var divContainer=document.getElementById('colorsContainer');
		if(divContainer)
		{

			if(divContainer.style.display!=='block')
			{
				divContainer.style.left=(e.clientX + document.documentElement.scrollLeft + document.body.scrollLeft+1)+'px';
				divContainer.style.top=(e.clientY + document.documentElement.scrollTop + document.body.scrollTop+1)+'px';
				divContainer.style.display='block';
			}
			else
			{
				divContainer.style.display='none';
			}
		}
	}
}

function pickColor()
{
	if(this && this.title)
	{
		sendUbbCode('color=' +this.title);
		this.parentNode.parentNode.parentNode.style.display='none';
	}
}

var palette=document.getElementById('colorsPalette');
if(palette)
{
	var fragment=document.createDocumentFragment();

	for(var i=0;i<colors.length;i++)
	{
		var divRow=document.createElement('div');
		divRow.className='colorsRow';

		for(var j=0;j<colors[i].length;j++)
		{
			var divCell=document.createElement('a');
			divCell.className='colorsCell';
			divCell.style.backgroundColor=colors[i][j];
			divCell.title=colors[i][j];
			divCell.onclick=pickColor;
			divCell.onmouseover=hoverColorSquare;

			divRow.appendChild(divCell);
		}

		fragment.appendChild(divRow);
	}

	palette.appendChild(fragment);
}
