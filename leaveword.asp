<!-- #include file="config.asp" -->

<%
Response.Expires = -1
Response.AddHeader "Pragma","no-cache"
Response.AddHeader "cache-control","no-cache, must-revalidate"

if checkIsBannedIP then
	Response.Redirect "err.asp?number=1"
	Response.End
elseif StatusOpen=false then
	Response.Redirect "err.asp?number=2"
	Response.End
elseif StatusWrite=false then
	Response.Redirect "err.asp?number=3"
	Response.End
end if
if StatusStatistics then call addstat("leaveword")
if WriteVcodeCount>0 then session("vcode")=getvcode(WriteVcodeCount)

function getstatus(isopen)
	if isopen then
		getstatus="<span style=""font-weight:bold; color:#008000;"">√</span>"
	else
		getstatus="<span style=""font-weight:bold; color:#FF0000;"">×</span>"
	end if
end function
%>

<!-- #include file="inc_dtd.asp" -->
<html>
<head>
	<!-- #include file="inc_metatag.asp" -->
	<title><%=HomeName%> 留言本 签写留言</title>
	<link rel="stylesheet" type="text/css" href="css/style.css"/>
	<!-- #include file="css/style.asp" -->
	<script type="text/javascript">
		function submitcheck(cobject)
		{
			if (cobject.ivcode && cobject.ivcode.value==='') {alert('请输入验证码。'); if(tab){tab.selectPage(0); cobject.ivcode.focus();} return false;}
			if (cobject.iname.value==='') {alert('请输入称呼。'); if(tab){tab.selectPage(0); cobject.iname.focus();} return false;}
			if (cobject.ititle.value==='') {alert('请输入标题。'); if(tab){tab.selectPage(0); cobject.ititle.focus();} return false;}
			if (cobject.chk_encryptwhisper && cobject.chk_encryptwhisper.checked && cobject.iwhisperpwd && cobject.iwhisperpwd.value==='') {alert('请输入悄悄话密码。'); if(tab){tab.selectPage(0); cobject.iwhisperpwd.focus();} return false;}
			if (cobject.imailreplyinform && cobject.imailreplyinform.checked && cobject.imail.value==='') {alert('请输入邮件地址以便回复时通知，或者去掉回复通知选项。'); if(tab){tab.selectPage(1); cobject.imail.focus();} return false;}
			cobject.submit1.disabled=true;
			return (true);
		}

		function chkoption(frm)
		{
			var objlbl;
			if (frm && frm.chk_whisper && frm.chk_whisper.checked && frm.chk_encryptwhisper) {		//悄悄话yes
				frm.chk_encryptwhisper.disabled=false;
				if (document && document.getElementById && document.getElementById('lbl_encryptwhisper')) {
					objlbl=document.getElementById('lbl_encryptwhisper');
					if(objlbl && (typeof(objlbl.disabled)).toLowerString!='undefined') objlbl.disabled=false;
				}

				objlbl=null;
				if (document && document.getElementById && document.getElementById('lbl_whisperpwd')) objlbl=document.getElementById('lbl_whisperpwd');
				if (frm.chk_encryptwhisper.checked) {			//加密悄悄话yes
					if(objlbl && (typeof(objlbl.disabled)).toLowerString!='undefined') objlbl.disabled=false;
					frm.iwhisperpwd.disabled=false;
				} else {			//加密悄悄话no
					if(objlbl && (typeof(objlbl.disabled)).toLowerString!='undefined') objlbl.disabled=true;
					frm.iwhisperpwd.disabled=true;
				}
			} else if (frm && frm.chk_whisper && !frm.chk_whisper.checked && frm.chk_encryptwhisper) {		//悄悄话no
				frm.chk_encryptwhisper.disabled=true;
				if (document && document.getElementById && document.getElementById('lbl_encryptwhisper')) {objlbl=document.getElementById('lbl_encryptwhisper'); if(objlbl &&(typeof(objlbl.disabled)).toLowerString!='undefined') objlbl.disabled=true;}
				if (document && document.getElementById && document.getElementById('lbl_whisperpwd')) {objlbl=document.getElementById('lbl_whisperpwd'); if(objlbl && (typeof(objlbl.disabled)).toLowerString!='undefined') objlbl.disabled=true;}
				frm.iwhisperpwd.disabled=true;
			}
		}

		var bkcontent='';
		function checklength(txtobj,max)
		{
			if (txtobj.value.length>max) {txtobj.value=bkcontent;alert('已达最大字数限制！');} else bkcontent=txtobj.value;
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
				divPreview.appendChild(document.createTextNode('正在生成预览，请稍候……'));
				
				xmlHttp.abort();
				xmlHttp.onreadystatechange=previewArrived;
				xmlHttp.open('POST','leaveword_preview.asp'+window.location.search);
				xmlHttp.setRequestHeader('Content-Type','application/x-www-form-urlencoded');
				xmlHttp.send('icontent=' + escape(iContent.value));
			}
		}
		
		function previewArrived()
		{
			if(xmlHttp.readyState===4 && xmlHttp.status===200)
			{
				var divPreview=document.getElementById('divPreview');
				if(xmlHttp && divPreview)
				{
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
			show_book_title 3,"回复留言"
		else
			show_book_title 3,"签写留言"
		end if
	end if
	%>

	<div class="region">
		<h3 class="title">欢迎您留言</h3>
		<div class="content">
			<form method="post" action="write.asp" onsubmit="return submitcheck(this)" name="form1">
				<input type="hidden" name="follow" value="<%=request("follow")%>"/>
				<input type="hidden" name="return" value="<%=request("return")%>"/>
				<input type="hidden" name="qstr" value="<%=Server.HtmlEncode(request.QueryString)%>"/>

				<div id="tabContainer">
				</div>

				<div id="divRequired">
					<h4>必填项目：</h4>

					<%if WriteVcodeCount>0 then%>
					<div class="field">
						<span class="label">验证码<span class="required">*</span></span>
						<span class="value"><input type="text" name="ivcode" size="<%=LeaveVcodeWidth%>" autocomplete="off"/><img class="captcha" src="show_vcode.asp?type=write"/></span>
					</div>
					<%end if%>
					<div class="field">
						<span class="label">称呼<span class="required">*</span></span>
						<span class="value"><input type="text" name="iname" size="<%=LeaveTextWidth%>" maxlength="20" value="<%=server.htmlEncode(FormOrCookie("iname"))%>" /></span>
					</div>
					<div class="field">
						<span class="label">标题<span class="required">*</span></span>
						<span class="value"><input type="text" name="ititle" size="<%=LeaveTextWidth%>" maxlength="30" value="<%=server.htmlEncode(FormOrSession(InstanceName & "_ititle"))%><%if Request.Form("ititle")="" and isnumeric(Request("follow")) and Request("follow")<>"" then response.write "Re:"%>"/></span>
					</div>
					<div class="field">
						<div class="row">内容： <%=getstatus(HTMLSupport)%>HTML标记　<%=getstatus(UBBSupport)%>UBB标记<%if HTMLSupport=false and UBBSupport=false and AllowNewLine=true then Response.Write "　" & getstatus(true) & "允许换行"%></div>
						<div class="row"><textarea name="icontent" id="icontent" cols="<%=LeaveContentWidth%>" rows="<%=LeaveContentHeight%>" onkeydown="icontent_keydown(arguments[0]);"<%if WordsLimit>0 then Response.Write " onpropertychange=""checklength(this,"&WordsLimit&");"""%>><%=server.htmlEncode(FormOrSession(InstanceName & "_icontent"))%></textarea></div>
						<!-- #include file="ubbtoolbar.inc" -->
						<%if UBBSupport then ShowUbbToolBar(false)%>
					</div>
					<%if StatusWhisper=true then%>
					<div class="field">
						<div class="row">
							<img src="image/icon_whisper.gif" class="imgicon" />　
							<input type="checkbox" name="chk_whisper" value="1" id="chk_whisper" onclick="chkoption(this.form)"<%=cked(Request.Form("chk_whisper")="1")%> /><label id="lbl_whisper" for="chk_whisper">悄悄话</label>
							<%if StatusEncryptWhisper=true then%>　<input type="checkbox" name="chk_encryptwhisper" value="1" id="chk_encryptwhisper" onclick="chkoption(this.form);if(this.checked)this.form.iwhisperpwd.select();"<%=cked(Request.Form("chk_encryptwhisper")="1")%><%=dised(Request.Form("chk_whisper")<>"1")%> /><label id="lbl_encryptwhisper" for="chk_encryptwhisper"<%=dised(Request.Form("chk_whisper")<>"1")%>>加密悄悄话</label><%end if%>
						</div>
						<%if StatusEncryptWhisper=true then%>
						<div class="row"><img border="0" src="image/icon_key.gif" class="imgicon" />　<label id="lbl_whisperpwd"<%if Request.Form("chk_whisper")<>"1" or Request.Form("chk_encryptwhisper")<>"1" then Response.Write " disabled=""disabled"""%>>密码</label> <input type="password" name="iwhisperpwd" id="iwhisperpwd" size="<%=LeaveTextWidth%>" maxlength="16" title="为悄悄话设置密码后，必须提供密码才能查看回复，也可以查看原先留言。" value="<%=server.HTMLEncode(Request.Form("iwhisperpwd"))%>"<%if Request.Form("chk_whisper")<>"1" or Request.Form("chk_encryptwhisper")<>"1" then Response.Write " disabled=""disabled"""%> /></div>
						<%end if%>
					</div>
					<%end if%>
				</div>

				<div id="divContact">
					<h4>联系方式：</h4>
					<div class="field">
						<span class="label"><img src="image/icon_mail.gif" class="imgicon" />邮件</span>
						<span class="value"><input type="text" name="imail" size="<%=LeaveTextWidth%>" maxlength="50" value="<%=server.htmlEncode(FormOrCookie("imail"))%>"/><%if MailReplyInform=true then%><br/><input type="checkbox" name="imailreplyinform" id="imailreplyinform" value="1"<%=cked(Request.Form("imailreplyinform")="1")%> /><label for="imailreplyinform">版主回复后用邮件通知我</label><%end if%></span>
					</div>
					<div class="field">
						<span class="label"><img src="image/icon_qq.gif" class="imgicon" />QQ号</span>
						<span class="value"><input type="text" name="iqq" size="<%=LeaveTextWidth%>" maxlength="16" value="<%=server.htmlEncode(FormOrCookie("iqq"))%>"/></span>
					</div>
					<div class="field">
						<span class="label"><img src="image/icon_skype.gif" class="imgicon" />Skype</span>
						<span class="value"><input type="text" name="imsn" size="<%=LeaveTextWidth%>" maxlength="50" value="<%=server.htmlEncode(FormOrCookie("imsn"))%>"/></span>
					</div>
					<div class="field">
						<span class="label"><img src="image/icon_homepage.gif" class="imgicon" />主页</span>
						<span class="value"><input type="text" name="ihomepage" size="<%=LeaveTextWidth%>" maxlength="127" value="<%=server.htmlEncode(FormOrCookie("ihomepage"))%>"/></span>
					</div>
					<div class="field">
						<input type="checkbox" name="hidecontact" id="hidecontact" value="1"<%=cked(request.form("hidecontact")="1")%> /><label for="hidecontact">联系方式仅版主可见</label>
					</div>
				</div>

				<%if StatusShowHead then%>
				<div id="divFace">
					<h4>头像：</h4>
					<%defaultindex=FormOrCookie("ihead")%>
                    <!-- #include file="listface.inc" -->
				</div>
				<%end if%>

				<div id="divUbbhelp">
					<h4>UBB帮助：</h4>
					<!-- #include file="inc_ubbhelp.asp" -->
				</div>

				<script type="text/javascript" src="js/tabcontrol.js"></script>
				<script type="text/javascript">
					tab=new TabControl('tabContainer');

					tab.setOuterContainerCssClass('tab-outer-container');
					tab.setTitleContainerCssClass('tab-title-container');
					tab.setTitleCssClass('tab-title');
					tab.setTitleSelectedCssClass('tab-title-selected');
					tab.setPageContainerCssClass('tab-page-container');
					tab.setPageCssClass('tab-page');

					tab.addPage('divRequired','必填项目');
					tab.addPage('divContact','联系方式');
					tab.addPage('divFace','头像');
					tab.addPage('divUbbhelp','UBB帮助');
				</script>

				<p align="center"><input type="submit" value="发表留言" name="submit1" id="submit1" />　<input type="reset" value="重写留言" />　<input type="button" value="预览留言" onclick="previewRequest(this.value);" /></p>

				<div id="divPreview"></div>
			</form>
		</div>
	</div>
</div>

<!-- #include file="bottom.asp" -->
<%if StatusStatistics and session("gotclientinfo")<>true then%><script type="text/javascript" src="getclientinfo.asp" defer="defer" async="async"></script><%end if%>
</body>
</html>