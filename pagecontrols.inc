<script type="text/javascript">
//A JavaScript code for smooth-page-scrolling
//Made by MJ PC Lab 2005, All rights reserved.
//http://mjpclab.net/

function blur()
{
	if(this && this.blur) this.blur();
}

function turn_light()
{
	if(this && this.style)
	{
		this.style.borderColor='<%=TableBorderColor%>';
		this.style.backgroundColor='<%=BlockBGC%>';
	}
}

function turn_normal()
{
	if(this && this.style)
	{
		this.style.borderColor='';
		this.style.backgroundColor='';
	}
}

function refresh_pagenum(increment)
{
	if(document && document.getElementsByName)
	{
		e=document.getElementsByName('pagenum');
		if(e && e.length && e.length>0 && e[0] && e[0].firstChild && e[0].firstChild.data)
		{
			var old_start=Number(e[0].firstChild.data.replace(/^\[(\d+)\]$/,'$1'));
			var new_start=old_start+increment;
			
			if(new_start<1) new_start=1;
			if(new_start+PageListCount-1>PagesCount) new_start=PagesCount-PageListCount+1;
			if(new_start<1) new_start=1;
			
			var delta=new_start-old_start;
			for(var i=0;i<e.length && delta!=0;i++)
			{
				var old_current=Number(e[i].firstChild.data.replace(/^\[(\d+)\]$/,'$1'));
				var new_current=old_current+delta;

				e[i].href=e[i].href.replace(/page=\d+/i,'page='+new_current.toString(10))
				e[i].firstChild.data='[' + new_current.toString(10) + ']';
				
				if(new_current==CurrentPage)
					e[i].className='pagenum_curr';
				else
					e[i].className='pagenum_normal';
				
			}
			
		}
	}
	if(setTimeout) ghandle=setTimeout('refresh_pagenum(' +increment.toString(10)+ ')',250);
}

var ghandle;
var PagesCount=<%=PagesCount%>;
var PageListCount=<%=AdvPageListCount%>;
var CurrentPage=<%=CurPage%>;

//mouse effects
if (document && document.getElementsByName)
{
	var e=document.getElementsByName('page_control');
	if(e && e.length && e.length>0)
	{
		for(var i=0;i<e.length;i++)
		{
			e[i].onfocus=blur;
			if(e[i].firstChild)
			{
				e[i].firstChild.onmouseover=turn_light;
				e[i].firstChild.onmouseout=turn_normal;
			}
		}
	}
	
	e=document.getElementsByName('js_page_control');
	if(e && e.length && e.length>0)
	{
		for(var i=0;i<e.length;i++)
		{
			e[i].onmouseover=turn_light;
			e[i].onmouseout=turn_normal;
			if(e[i].style) e[i].style.visibility='';
		}
	}
}
</script>