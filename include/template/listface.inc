<script type="text/javascript">
	function selhead(ihead_id)
	{
		var iRadio=document.getElementById('ihead_' + ihead_id.toString());
		if(iRadio) iRadio.checked=true;
		
		if(form1 && form1.afaceid)
		{
			form1.afaceid.value=ihead_id;
			form1.afaceurl.value='';
		}
	}
	
	function loadHeadRest(hideNode)
	{
		var minHeadId=<%=FrequentFaceCount%>;
		var maxHeadId=<%=FaceCount%>;
		var picsPerRow=<%=PicturesPerRow%>;
		var facePath='asset/face/';
		
		for(var i=minHeadId; i<=maxHeadId+(picsPerRow - (maxHeadId+1) % picsPerRow); i++)
		{
			var cell=document.getElementById('cell_' + i.toString());
			var radio=document.getElementById('ihead_' + i.toString());
			
			if(cell) cell.style.display='';
			if(radio) radio.style.display='';
		}
		
		for(var i=minHeadId; i<=maxHeadId; i++)
		{
			var img=document.getElementById('face_' + i.toString());
			if(img)
			{
				img.src=facePath + i.toString() + '.gif';
				img.style.display='';
			}
		}

		if(hideNode && hideNode.style) {
			hideNode.style.display='none';
		}
	}
</script>

<table class="grid">
	<%
	dim lines
	lines=(FaceCount+1) \ PicturesPerRow
	if (FaceCount+1) mod PicturesPerRow <>0 then lines=lines+1
	
	dim selindex
	if isnumeric(Request.QueryString("follow")) And Request.QueryString("follow")<>"" then
		selindex=0
	elseif isnumeric(defaultindex) and defaultindex<>"" then
		if clng(defaultindex)>=0 then selindex=clng(defaultindex)
	else
		selindex=1
	end if

	
	dim ihead_id,hidestyle,imagesrc
	function FormatHeadHtml(byref str)
		FormatHeadHtml=replace(replace(replace(str,"{0}",ihead_id),"{1}",hidestyle),"{2}",imagesrc)
	end function
	
	dim i,j
	for i=0 to lines-1
		Response.Write "<tr>"
		for j=0 to PicturesPerRow-1
		
			ihead_id=i*PicturesPerRow+j
			
			if ihead_id+1 <= FrequentFaceCount Or ihead_id=0 then
				hidestyle=" style="""""
				imagesrc= " src=""asset/face/" & ihead_id & ".gif"""
			else
				hidestyle=" style=""display:none;"""
				imagesrc=" src=""about:blank"""
			end if
			
			Response.Write FormatHeadHtml("<td id=""cell_{0}"" {1}>")
			
			if ihead_id <= FaceCount then
				Response.Write FormatHeadHtml("<input type=""radio"" name=""ihead"" id=""ihead_{0}"" value=""{0}"" {1}")
				if cbyte(i*PicturesPerRow+j)=selindex then Response.Write " checked=""checked"""
				Response.Write " />"

				if ihead_id=0 then
					Response.Write FormatHeadHtml("<label onmousedown=""selhead({0})"">(无头像)</label>")
				else
					Response.Write FormatHeadHtml("<img id=""face_{0}"" alt="""" class=""listface"" {1} {2} onmousedown=""selhead({0})"" />")
				end if
			end if
			Response.Write "</td>"
		next
		Response.Write "</tr>"
	next
	%>
	<tr>
		<td colspan="<%=PicturesPerRow%>">
			<a href="javascript:void 0;" onclick="loadHeadRest(this.parentNode.parentNode);">更多头像...</a>
		</td>
	</tr>
</table>