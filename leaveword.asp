<!-- #include file="include/template/page_instruction.inc" -->
<!-- #include file="config/system.asp" -->
<!-- #include file="config/database.asp" -->
<!-- #include file="include/sql/init.asp" -->
<!-- #include file="include/sql/common.asp" -->
<!-- #include file="include/utility/database.asp" -->
<!-- #include file="include/utility/ip.asp" -->
<!-- #include file="include/utility/string.asp" -->
<!-- #include file="include/utility/backend.asp" -->
<!-- #include file="include/utility/frontend.asp" -->
<!-- #include file="include/utility/book.asp" -->
<!-- #include file="loadconfig.asp" -->
<!-- #include file="error.asp" -->
<%
Response.Expires = -1
Response.AddHeader "Pragma","no-cache"
Response.AddHeader "cache-control","no-cache, must-revalidate"

if checkIsBannedIP() then
	Call ErrorPage(1)
	Response.End
elseif Not StatusOpen then
	Call ErrorPage(2)
	Response.End
elseif Not StatusWrite then
	Call ErrorPage(3)
	Response.End
end if
if StatusStatistics then call addstat("leaveword")
if WriteVcodeCount>0 then Session(InstanceName & "_vcode_write")=getvcode(WriteVcodeCount)

function getstatus(isopen)
	if isopen then
		getstatus="<span style=""font-weight:bold; color:#008000;"">√</span>"
	else
		getstatus="<span style=""font-weight:bold; color:#FF0000;"">×</span>"
	end if
end function
%>

<!-- #include file="include/template/dtd.inc" -->
<html>
<head>
	<!-- #include file="include/template/metatag.inc" -->
	<title><%=HomeName%> 留言本 签写留言</title>
	<!-- #include file="inc_stylesheet.asp" -->

	<script type="text/javascript">
	function submitcheck(frm)
	{
		if (frm.ivcode && frm.ivcode.value.length===0) {alert('请输入验证码。'); if(tab){tab.selectPage(0); frm.ivcode.focus();} return false;}
		if (frm.iname.value.length===0) {alert('请输入称呼。'); if(tab){tab.selectPage(0); frm.iname.focus();} return false;}
		if (frm.ititle.value.length===0) {alert('请输入标题。'); if(tab){tab.selectPage(0); frm.ititle.focus();} return false;}
		if (frm.chk_encryptwhisper && frm.chk_encryptwhisper.checked && frm.iwhisperpwd && frm.iwhisperpwd.value.length===0) {alert('请输入悄悄话密码。'); if(tab){tab.selectPage(0); frm.iwhisperpwd.focus();} return false;}
		if (frm.imailreplyinform && frm.imailreplyinform.checked && frm.imail.value.length===0) {alert('请输入邮件地址以便回复时通知，或者去掉回复通知选项。'); if(tab){tab.selectPage(1); frm.imail.focus();} return false;}
		frm.submit1.disabled=true;
		return true;
	}

	function updateWhisperField(frm)
	{
		var chkWhisper=frm.chk_whisper;

		var chkEncryptWhisper=frm.chk_encryptwhisper;
		var lblEncryptWhisper=document.getElementById('lbl_encryptwhisper');

		var iWhisperPwd=frm.iwhisperpwd;

		chkEncryptWhisper.disabled=lblEncryptWhisper.disabled=!chkWhisper.checked;
		if(!chkWhisper.checked) {
			chkEncryptWhisper.checked=false;
			iWhisperPwd.value='';
		}

		iWhisperPwd.disabled=!chkEncryptWhisper.checked;
		if(!chkEncryptWhisper.checked) {
			iWhisperPwd.value='';
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

	<!-- #include file="asset/js/xmlhttp.js" -->
	function previewRequest()
	{
		if(!window.xmlHttp) window.xmlHttp=createXmlHttp();
		var divPreview=document.getElementById('divPreview');
		var iContent=document.getElementById('icontent');

		if(xmlHttp && divPreview && iContent)
		{
			setPureText(divPreview, '正在生成预览，请稍候……');

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
			InitHeaderData("回复留言")
		else
			InitHeaderData("签写留言")
		end if
		%><!-- #include file="include/template/header.inc" --><%
	end if
	%>

	<div id="mainborder" class="mainborder">
	<div class="region">
		<h3 class="title">欢迎您留言</h3>
		<div class="content">
			<form method="post" action="write.asp" onsubmit="return submitcheck(this)" name="form1">
				<input type="hidden" name="follow" value="<%=request("follow")%>"/>
				<input type="hidden" name="return" value="<%=request("return")%>"/>
				<input type="hidden" name="qstr" value="<%=HtmlEncode(request.QueryString)%>"/>

				<div id="tabContainer"></div>

				<div id="divRequired">
					<h4>必填项目：</h4>

					<%if WriteVcodeCount>0 then%>
					<div class="field">
						<span class="label">验证码<span class="required">*</span></span>
						<span class="value"><input type="text" name="ivcode" autocomplete="off"/><img id="captcha" class="captcha" src="show_vcode.asp?type=write&t=0"/></span>
					</div>
					<%end if%>
					<div class="field">
						<span class="label">称呼<span class="required">*</span></span>
						<span class="value"><input type="text" name="iname" class="longtext" maxlength="20" value="<%=HtmlEncode(FormOrCookie("iname"))%>" /></span>
					</div>
					<div class="field">
						<span class="label">标题<span class="required">*</span></span>
						<span class="value"><input type="text" name="ititle" class="longtext" maxlength="30" value="<%=HtmlEncode(FormOrSession(InstanceName & "_ititle"))%><%if Request.Form("ititle")="" and isnumeric(Request("follow")) and Request("follow")<>"" then response.write "Re:"%>"/></span>
					</div>
					<div class="field">
						<div class="row">内容： <%=getstatus(HTMLSupport)%>HTML标记　<%=getstatus(UBBSupport)%>UBB标记<%if Not HTMLSupport and UBBSupport=false and AllowNewLine then Response.Write "　" & getstatus(true) & "允许换行"%></div>
						<div class="row"><textarea name="icontent" id="icontent" rows="<%=LeaveContentHeight%>" onkeydown="icontent_keydown(arguments[0]);"<%if WordsLimit>0 then Response.Write " onpropertychange=""checklength(this,"&WordsLimit&");"""%>><%=HtmlEncode(FormOrSession(InstanceName & "_icontent"))%></textarea></div>
						<!-- #include file="include/template/ubbtoolbar.inc" -->
						<%if UBBSupport then ShowUbbToolBar(false)%>
					</div>
					<%if StatusWhisper then%>
					<div class="field">
						<div class="row">
							<img src="asset/image/icon_whisper.gif" class="imgicon" />　
							<input type="checkbox" name="chk_whisper" value="1" id="chk_whisper" onclick="updateWhisperField(this.form)"<%=cked(Request.Form("chk_whisper")="1")%> /><label id="lbl_whisper" for="chk_whisper">悄悄话</label>
							<%if StatusEncryptWhisper then%>　<input type="checkbox" name="chk_encryptwhisper" value="1" id="chk_encryptwhisper" onclick="updateWhisperField(this.form);if(this.checked)this.form.iwhisperpwd.select();"<%=cked(Request.Form("chk_encryptwhisper")="1")%><%=dised(Request.Form("chk_whisper")<>"1")%> /><label id="lbl_encryptwhisper" for="chk_encryptwhisper"<%=dised(Request.Form("chk_whisper")<>"1")%>>加密悄悄话</label><%end if%>
						</div>
						<%if StatusEncryptWhisper then%>
						<div class="row"><img border="0" src="asset/image/icon_key.gif" class="imgicon" />　<label id="lbl_whisperpwd"<%if Request.Form("chk_whisper")<>"1" or Request.Form("chk_encryptwhisper")<>"1" then Response.Write " disabled=""disabled"""%>>密码</label> <input type="password" name="iwhisperpwd" id="iwhisperpwd" maxlength="16" title="为悄悄话设置密码后，必须提供密码才能查看回复，也可以查看原先留言。" value="<%=HtmlEncode(Request.Form("iwhisperpwd"))%>"<%if Request.Form("chk_whisper")<>"1" or Request.Form("chk_encryptwhisper")<>"1" then Response.Write " disabled=""disabled"""%> /></div>
						<%end if%>
					</div>
					<%end if%>
				</div>

				<div id="divContact">
					<h4>联系方式：</h4>
					<div class="field">
						<span class="label"><img src="asset/image/icon_mail.gif" class="imgicon" />邮件</span>
						<span class="value"><input type="text" name="imail" class="longtext" maxlength="50" value="<%=HtmlEncode(FormOrCookie("imail"))%>"/><%if MailReplyInform then%><br/><input type="checkbox" name="imailreplyinform" id="imailreplyinform" value="1"<%=cked(Request.Form("imailreplyinform")="1")%> /><label for="imailreplyinform">版主回复后用邮件通知我</label><%end if%></span>
					</div>
					<div class="field">
						<span class="label"><img src="asset/image/icon_qq.gif" class="imgicon" />QQ号</span>
						<span class="value"><input type="text" name="iqq" class="longtext" maxlength="16" value="<%=HtmlEncode(FormOrCookie("iqq"))%>"/></span>
					</div>
					<div class="field">
						<span class="label"><img src="asset/image/icon_skype.gif" class="imgicon" />Skype</span>
						<span class="value"><input type="text" name="imsn" class="longtext" maxlength="50" value="<%=HtmlEncode(FormOrCookie("imsn"))%>"/></span>
					</div>
					<div class="field">
						<span class="label"><img src="asset/image/icon_homepage.gif" class="imgicon" />主页</span>
						<span class="value"><input type="text" name="ihomepage" class="longtext" maxlength="127" value="<%=HtmlEncode(FormOrCookie("ihomepage"))%>"/></span>
					</div>
					<div class="field">
						<input type="checkbox" name="hidecontact" id="hidecontact" value="1"<%=cked(request.form("hidecontact")="1")%> /><label for="hidecontact">联系方式仅版主可见</label>
					</div>
				</div>

				<%if StatusShowHead then%>
				<div id="divFace">
					<h4>头像：</h4>
					<%defaultindex=FormOrCookie("ihead")%>
                    <!-- #include file="include/template/listface.inc" -->
				</div>
				<%end if%>

				<div id="divUbbhelp">
					<h4>UBB帮助：</h4>
					<!-- #include file="include/template/ubbhelp.inc" -->
				</div>

				<script type="text/javascript" src="asset/js/tabcontrol.js"></script>
				<script type="text/javascript">
					var tab=new TabControl('tabContainer');

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

	<!-- #include file="include/template/footer.inc" -->
</div>

<script type="text/javascript">
	<!-- #include file="asset/js/refresh-captcha.js" -->
</script>
<!-- #include file="include/template/getclientinfo.inc" -->
</body>
</html>