<%Sub ShowUbbToolBar(byval allflags)%>
<%if ShowUbbTool=true then%>
	<div class="ubbtoolbar">
		<div class="row">
			<%if UbbFlag_fontstyle Or allflags then%><a name="ubbbutton" class="ubbbutton ubb_bold" onclick="sendUbbCode &&sendUbbCode('b')"><img alt="粗体" src="image/icon_ubb_bold.gif"/></a><%end if%>
			<%if UbbFlag_fontstyle Or allflags then%><a name="ubbbutton" class="ubbbutton ubb_italic" onclick="sendUbbCode && sendUbbCode('i')"><img alt="斜体" src="image/icon_ubb_italic.gif"/></a><%end if%>
			<%if UbbFlag_fontcolor Or allflags then%><a name="ubbbutton" class="ubbbutton ubb_color" onclick="toggleColorPalette && toggleColorPalette(this,arguments[0]||window.event);"><img alt="文字颜色" src="image/icon_ubb_color.gif"/></a><%end if%>
			<%if UbbFlag_image Or allflags then%><a name="ubbbutton" class="ubbbutton ubb_image" onclick="sendUbbCode && sendUbbCode('img')"><img alt="图片" src="image/icon_ubb_img.gif"/></a><%end if%>
			<%if UbbFlag_player Or allflags then%><a name="ubbbutton" class="ubbbutton ubb_flash" onclick="sendUbbCode && sendUbbCode('flash')"><img alt="Flash" src="image/icon_ubb_flash.gif"/></a><%end if%>
			<%if UbbFlag_player Or allflags then%><a name="ubbbutton" class="ubbbutton ubb_mplayer" onclick="sendUbbCode && sendUbbCode('mp')"><img alt="Windows Media Player" src="image/icon_ubb_mplayer.gif"/></a><%end if%>
			<%if UbbFlag_player Or allflags then%><a name="ubbbutton" class="ubbbutton ubb_realplayer" onclick="sendUbbCode && sendUbbCode('rm')"><img alt="Real Player" src="image/icon_ubb_realplayer.gif"/></a><%end if%>
		</div>
		<%if UbbFlag_face Or allflags then%>
		<div class="row">
			<%for i=1 to SmallFaceCount%><a name="ubbface" class="ubbface face<%=i%>" title="[face<%=i%>]"><img alt="[face<%=i%>]" src="<%=SmallFacePath & i%>.gif"/></a><%next%>
		</div>
		<%end if%>
	</div>
	
	<div class="colorsContainer" id="colorsContainer">
		<div class="colorsInfo" id="colorsInfo">
			<div class="colorsPreview" id="colorsPreview"></div>
			<div class="colorsNumber" id="colorsNumber">#??????</div>
			<div class="colorsCancel" id="colorsCancel" onclick="this.parentNode.parentNode.style.display='none';">关闭</div>
		</div>
		<div class="colorsPalette" id="colorsPalette"></div>
	</div>
		
	
	<!-- #include file="ubbhook.inc" -->
<%end if%>
<%End Sub%>