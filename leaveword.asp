<!-- #include file="config.asp" -->

<%
Response.Expires = -1
Response.AddHeader "Pragma","no-cache"
Response.AddHeader "cache-control","no-cache, must-revalidate"

if isbanip(Request.ServerVariables("REMOTE_ADDR"))=true or isbanip(Request.ServerVariables("HTTP_X_FORWARDED_FOR"))=true then
	Response.Redirect "err.asp?number=1"
	Response.End
elseif StatusOpen=false then
	Response.Redirect "err.asp?number=2"
	Response.End
elseif StatusWrite=false then
	Response.Redirect "err.asp?number=3"
	Response.End
end if
call addstat("leaveword")
if WriteVcodeCount>0 then session("vcode")=getvcode(WriteVcodeCount)

function getstatus(isopen)
	if isopen then
		getstatus="<span style=""font-weight:bold; color:#008000;"">��</span>"
	else
		getstatus="<span style=""font-weight:bold; color:#FF0000;"">��</span>"
	end if
end function
%>

<!-- #include file="inc_dtd.asp" -->
<html>
<head>
	<!-- #include file="inc_metatag.asp" -->
	<title><%=HomeName%> ���Ա� ǩд����</title>
	<link rel="stylesheet" type="text/css" href="style.css"/>
	<!-- #include file="style.asp" -->
	<!-- #include file="getclientinfo.inc" -->
	<script type="text/javascript">
		function submitcheck(cobject)
		{
			if (cobject.ivcode && cobject.ivcode.value=='') {alert('��������֤�롣'); if(tab){tab.selectPage(0); cobject.ivcode.focus();} return false;}
			if (cobject.iname.value=='') {alert('������ƺ���'); if(tab){tab.selectPage(0); cobject.iname.focus();} return false;}
			if (cobject.ititle.value=='') {alert('��������⡣'); if(tab){tab.selectPage(0); cobject.ititle.focus();} return false;}
			if (cobject.chk_encryptwhisper && cobject.chk_encryptwhisper.checked && cobject.iwhisperpwd && cobject.iwhisperpwd.value=='') {alert('���������Ļ����롣'); if(tab){tab.selectPage(0); cobject.iwhisperpwd.focus();} return false;}
			if (cobject.imailreplyinform && cobject.imailreplyinform.checked && cobject.imail.value=='') {alert('�������ʼ���ַ�Ա�ظ�ʱ֪ͨ������ȥ���ظ�֪ͨѡ�'); if(tab){tab.selectPage(1); cobject.imail.focus();} return false;}
			cobject.submit1.disabled=true;
			return (true);
		}

		function chkoption(frm)
		{
			var objlbl;
			if (frm && frm.chk_whisper && frm.chk_whisper.checked && frm.chk_encryptwhisper) {		//���Ļ�yes
				frm.chk_encryptwhisper.disabled=false;
				if (document && document.getElementById && document.getElementById('lbl_encryptwhisper')) {
					objlbl=document.getElementById('lbl_encryptwhisper');
					if(objlbl && (typeof(objlbl.disabled)).toLowerString!='undefined') objlbl.disabled=false;
				}

				objlbl=null;
				if (document && document.getElementById && document.getElementById('lbl_whisperpwd')) objlbl=document.getElementById('lbl_whisperpwd');
				if (frm.chk_encryptwhisper.checked) {			//�������Ļ�yes
					if(objlbl && (typeof(objlbl.disabled)).toLowerString!='undefined') objlbl.disabled=false;
					frm.iwhisperpwd.disabled=false;
				} else {			//�������Ļ�no
					if(objlbl && (typeof(objlbl.disabled)).toLowerString!='undefined') objlbl.disabled=true;
					frm.iwhisperpwd.disabled=true;
				}
			} else if (frm && frm.chk_whisper && !frm.chk_whisper.checked && frm.chk_encryptwhisper) {		//���Ļ�no
				frm.chk_encryptwhisper.disabled=true;
				if (document && document.getElementById && document.getElementById('lbl_encryptwhisper')) {objlbl=document.getElementById('lbl_encryptwhisper'); if(objlbl &&(typeof(objlbl.disabled)).toLowerString!='undefined') objlbl.disabled=true;}
				if (document && document.getElementById && document.getElementById('lbl_whisperpwd')) {objlbl=document.getElementById('lbl_whisperpwd'); if(objlbl && (typeof(objlbl.disabled)).toLowerString!='undefined') objlbl.disabled=true;}
				frm.iwhisperpwd.disabled=true;
			}
		}

		var bkcontent='';
		function checklength(txtobj,max)
		{
			if (txtobj.value.length>max) {txtobj.value=bkcontent;alert('�Ѵ�����������ƣ�');} else bkcontent=txtobj.value;
		}

		function icontent_keydown(e)
		{
			if(!e)e=window.event;
			if(e && !e.target)e.target=e.srcElement;
			if(e && e.ctrlKey && e.keyCode==13)e.target.form.submit1.click();
		}

		<!-- #include file="xmlhttp.inc" -->
		function previewRequest()
		{
			if(!window.xmlHttp) window.xmlHttp=createXmlHttp();
			var divPreview=document.getElementById('divPreview');
			var iContent=document.getElementById('icontent');
			
			if(xmlHttp && divPreview && iContent)
			{
				clearChildren(divPreview);
				divPreview.style['textAlign']='center';
				divPreview.appendChild(document.createTextNode('��������Ԥ�������Ժ򡭡�'));
				
				xmlHttp.abort();
				xmlHttp.onreadystatechange=previewArrived;
				xmlHttp.open('POST','leaveword_preview.asp'+window.location.search);
				xmlHttp.setRequestHeader('Content-Type','application/x-www-form-urlencoded');
				xmlHttp.send('icontent=' + escape(iContent.value));
			}
		}
		
		function previewArrived()
		{
			if(xmlHttp.readyState==4 && xmlHttp.status==200)
			{
				var divPreview=document.getElementById('divPreview');
				if(xmlHttp && divPreview)
				{
					clearChildren(divPreview);
					divPreview.style['textAlign']='';
					
					divPreview.innerHTML=xmlHttp.responseText;
					xmlHttp.abort();
				}
			}
		}
</script>
</head>

<body<%=bodylimit%> onload="<%=framecheck%>if(tab && tab.selectedIndex==0 && form1 && form1.ivcode && form1.ivcode.value==''){form1.ivcode.focus();}">

<div id="outerborder" class="outerborder">
	<%
	if ShowTitle then
		if StatusGuestReply and isnumeric(request("follow")) and request("follow")<>"" then
			show_book_title 3,"�ظ�����"
		else
			show_book_title 3,"ǩд����"
		end if
	end if
	%>

<table border="1" cellpadding="2" class="generalwindow">
	<tr>
		<td class="wordstitle"><marquee direction="left" scrollamount="2" scrolldelay="100">��ӭ�����ԡ�</marquee></td>
	</tr>
	<tr>
		<td style="width:100%; background-color:<%=TableContentBGC%>;">
			<form method="post" action="write.asp" onsubmit="return submitcheck(this)" name="form1">
				<div id="divContainer">
					<input type="hidden" name="follow" value="<%=request("follow")%>"/>
					<input type="hidden" name="return" value="<%=request("return")%>"/>
					<input type="hidden" name="qstr" value="<%=Server.HtmlEncode(request.QueryString)%>"/>
				</div>
				
				<div id="divRequired" style="padding:5px;">
					<span style="font-weight:bold;">������Ŀ��</span><br/><br/>
					
					<%if WriteVcodeCount>0 then%>��֤��<span style="color:#FF0000">*</span><input type="text" name="ivcode" size="<%=LeaveVcodeWidth%>"/> <img src="show_vcode.asp?type=write" class="vcode" onclick="this.src=this.src" /><br/><br/><%end if%>
					
					�ơ���<span style="color:#FF0000">*</span><input type="text" name="iname" size="<%=LeaveTextWidth%>" maxlength="20" value="<%=server.htmlEncode(FormOrCookie("iname"))%>" /><br/><br/>
					
					�ꡡ��<span style="color:#FF0000">*</span><input type="text" name="ititle" size="<%=LeaveTextWidth%>" maxlength="30" value="<%=server.htmlEncode(FormOrSession(InstanceName & "_ititle"))%><%if Request.Form("ititle")="" and isnumeric(Request("follow")) and Request("follow")<>"" then response.write "Re:"%>"/><br/><br/>
					
					���ݣ� <%=getstatus(HTMLSupport)%>HTML��ǡ�<%=getstatus(UBBSupport)%>UBB���<%if HTMLSupport=false and UBBSupport=false and AllowNewLine=true then Response.Write "��" & getstatus(true) & "������"%>
					
					<br/><textarea name="icontent" id="icontent" cols="<%=LeaveContentWidth%>" rows="<%=LeaveContentHeight%>" onkeydown="icontent_keydown(arguments[0]);"<%if WordsLimit>0 then Response.Write " onpropertychange=""checklength(this,"&WordsLimit&");"""%>><%=server.htmlEncode(FormOrSession(InstanceName & "_icontent"))%></textarea>
					
					<!-- #include file="ubbtoolbar.inc" -->
					<%if UBBSupport then ShowUbbToolBar(false)%>
					<br/>

					<%if StatusWhisper=true then%>
						<img src="image/icon_whisper.gif" class="imgicon" />��<input type="checkbox" name="chk_whisper" value="1" id="chk_whisper" onclick="chkoption(this.form)"<%=cked(Request.Form("chk_whisper")="1")%> /><label id="lbl_whisper" for="chk_whisper">���Ļ�</label><%if StatusEncryptWhisper=true then%>��<input type="checkbox" name="chk_encryptwhisper" value="1" id="chk_encryptwhisper" onclick="chkoption(this.form);if(this.checked)this.form.iwhisperpwd.select();"<%=cked(Request.Form("chk_encryptwhisper")="1")%><%=dised(Request.Form("chk_whisper")<>"1")%> /><label id="lbl_encryptwhisper" for="chk_encryptwhisper"<%if Request.Form("chk_whisper")<>"1" then Response.Write " disabled=""disabled"""%>>�������Ļ�</label><br/>
									<img border="0" src="image/icon_key.gif" class="imgicon" />��<label id="lbl_whisperpwd"<%if Request.Form("chk_whisper")<>"1" or Request.Form("chk_encryptwhisper")<>"1" then Response.Write " disabled=""disabled"""%>>����</label> <input type="password" name="iwhisperpwd" id="iwhisperpwd" size="<%=LeaveTextWidth%>" maxlength="16" title="Ϊ���Ļ���������󣬱����ṩ������ܲ鿴�ظ���Ҳ���Բ鿴ԭ�����ԡ�" value="<%=server.HTMLEncode(Request.Form("iwhisperpwd"))%>"<%if Request.Form("chk_whisper")<>"1" or Request.Form("chk_encryptwhisper")<>"1" then Response.Write " disabled=""disabled"""%> /><br/>
						<%end if%>
						<br/>
					<%end if%>
				</div>
				
				<div id="divContact" style="padding:5px;">
					<span style="font-weight:bold;">��ϵ��ʽ��</span><br/><br/>
					<img src="image/icon_mail.gif" class="imgicon" />���ʼ� <input type="text" name="imail" size="<%=LeaveTextWidth%>" maxlength="50" value="<%=server.htmlEncode(FormOrCookie("imail"))%>"/><%if MailReplyInform=true then%><br/><img src="image/icon_mail.gif" class="imgicon" />��<input type="checkbox" name="imailreplyinform" id="imailreplyinform" value="1"<%=cked(Request.Form("imailreplyinform")="1")%> /><label for="imailreplyinform">�����ظ������ʼ�֪ͨ��</label><%end if%><br/><br/>
					<img src="image/icon_qq.gif" class="imgicon" />��QQ�� <input type="text" name="iqq" size="<%=LeaveTextWidth%>" maxlength="16" value="<%=server.htmlEncode(FormOrCookie("iqq"))%>"/><br/><br/>
					<img src="image/icon_msn.gif" class="imgicon" />��MSN. <input type="text" name="imsn" size="<%=LeaveTextWidth%>" maxlength="50" value="<%=server.htmlEncode(FormOrCookie("imsn"))%>"/><br/><br/>
					<img src="image/icon_homepage.gif" class="imgicon" />����ҳ <input type="text" name="ihomepage" size="<%=LeaveTextWidth%>" maxlength="127" value="<%=server.htmlEncode(FormOrCookie("ihomepage"))%>"/><br/><br/>
					<input type="checkbox" name="hidecontact" id="hidecontact" value="1"<%=cked(request.form("hidecontact")="1")%> /><label for="hidecontact">��ϵ��ʽ�������ɼ�</label><br/><br/>
				</div>
				
				<%if StatusShowHead then%>
				<div id="divFace" style="padding:5px;">
					<span style="font-weight:bold;">ͷ��</span><br/>
								
					<table cellpadding="0" style="width:100%; border-width:0px; border-collapse:collapse;">
					<tr style="height:16px;">
						<td style="width:100%; text-align:center; color:<%=TableContentColor%>;">��ѡ��ͷ��</td>
					</tr>
					<tr>
						<td style="width:100%; text-align:center;">
							<%dim defaultindex
							defaultindex=FormOrCookie("ihead")%>
							<!-- #include file="listface.inc" -->
						</td>
					</tr>
					</table>
				</div>
				<%end if%>
				
				<div id="divUbbhelp" style="padding:5px; text-align:center;">
					<div style="font-weight:bold; text-align:left;">UBB������</div><br/>
					<!-- #include file="inc_ubbhelp.asp" -->
				</div>
				
				<script type="text/javascript" src="tabcontrol.js"></script>
				<script type="text/javascript">
					tab=new TabControl('divContainer');
					
					tab.addPage('divRequired','������Ŀ');
					tab.addPage('divContact','��ϵ��ʽ');
					tab.addPage('divFace','ͷ��');
					tab.addPage('divUbbhelp','UBB����');
					
					tab.setOuterContainerCssClass('tabOuterContainer');
					tab.setTitleCssClass('tabTitle');
					tab.setTitleOverCssClass('tabTitleOver');
					tab.setTitleSelectedCssClass('tabTitleSelected');
					tab.setPageContainerCssClass('tabPageContainer');
					tab.setPageCssClass('tabPage');
				</script>

				<p align="center"><input type="submit" value="��������" name="submit1" id="submit1" />��<input type="reset" value="��д����" />��<input type="button" value="Ԥ������" onclick="previewRequest(this.value);" /></p>
				
				<table celpadding="10" cellspacing="0" style="border-width:0px; border-style:none; border-collapse:collapse; table-layout:fixed; overflow:auto;"><tr><td>
					<div id="divPreview" style="display:block; overflow:auto;"></div>
				</td></tr></table>
			</form>
		</td>
	</tr>
</table>
</div>

<!-- #include file="bottom.asp" -->
</body>
</html>