<script type="text/javascript">
//<![CDATA[

function turn_light() {
	if (this && this.style && typeof(this.style.backgroundColor).toLowerCase()=='string')
		this.style.backgroundColor='<%=BlockBGC%>';
}

function turn_normal() {
	if (this && this.style && this.style.backgroundColor)
		this.style.backgroundColor='<%=TableContentBGC%>';
}

function hook_rowlight(r)
{
	//for HTML DOM
	if (r && r.rows && r.rows.length)
	{
		for (var i=0;i<r.rows.length;i++)
		{
			r.rows[i].onmouseover=turn_light;
			r.rows[i].onmouseout=turn_normal;
		}
	}
	
	//for XHTML DOM
	if (r && r.childNodes && r.childNodes.length && r.childNodes.length>1)
	{
		for (var i=0;i<r.childNodes.length;i++)
		{
			r.childNodes[i].onmouseover=turn_light;
			r.childNodes[i].onmouseout=turn_normal;
		}
	}
}

if (document && document.getElementById)
{
	hook_rowlight(document.getElementById('ubblist'));
	hook_rowlight(document.getElementById('titlelist'));
}

//]]>
</script>