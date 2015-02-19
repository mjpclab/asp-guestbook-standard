<%Sub ShowUbbToolBar(byval allflags)%>
<%if ShowUbbTool=true then%>
	<div id="ubbtoolbar" style="width:<%=UbbToolWidth%>px; height:<%=UbbToolHeight%>px; border:none; overflow:auto; text-align:left;">
		<div>
			<%if UbbFlag_fontstyle Or allflags then%><img name="ubbbutton" id="ubb_bold" alt="粗体" src="image/icon_ubb_bold.gif" class="ubbbutton" onclick="if(sendUbbCode)sendUbbCode('[b]{0}[/b]')" /><%end if%>
			<%if UbbFlag_fontstyle Or allflags then%><img name="ubbbutton" id="ubb_italic" alt="斜体" src="image/icon_ubb_italic.gif" class="ubbbutton" onclick="if(sendUbbCode)sendUbbCode('[i]{0}[/i]')" /><%end if%>
			<%if UbbFlag_fontcolor Or allflags then%><img name="ubbbutton" id="ubb_color" alt="文字颜色" src="image/icon_ubb_color.gif" class="ubbbutton" onclick="if(toggleColorPalette) toggleColorPalette(this,window.event?window.event:arguments[0]);" /><%end if%>
			<%if UbbFlag_image Or allflags then%><img name="ubbbutton" id="ubb_image" alt="图片" src="image/icon_ubb_img.gif" class="ubbbutton" onclick="if(sendUbbCode)sendUbbCode('[img]{0}[/img]')" /><%end if%>
			<%if UbbFlag_player Or allflags then%><img name="ubbbutton" id="ubb_flash" alt="Flash" src="image/icon_ubb_flash.gif" class="ubbbutton" onclick="if(sendUbbCode)sendUbbCode('[flash]{0}[/flash]')" /><%end if%>
			<%if UbbFlag_player Or allflags then%><img name="ubbbutton" id="ubb_mplayer" alt="Windows Media Player" src="image/icon_ubb_mplayer.gif" class="ubbbutton" onclick="if(sendUbbCode)sendUbbCode('[mp]{0}[/mp]')" /><%end if%>
			<%if UbbFlag_player Or allflags then%><img name="ubbbutton" id="ubb_realplayer" alt="Real Player" src="image/icon_ubb_realplayer.gif" class="ubbbutton" onclick="if(sendUbbCode)sendUbbCode('[rm]{0}[/rm]')" /><%end if%>
		</div>
		<%if UbbFlag_face Or allflags then%>
			<%for i=1 to SmallFaceCount%><img name="ubbface" id="face<%=i%>" alt="[face<%=i%>]" src="<%=SmallFacePath & i%>.gif" class="smallface" /><%next%>
		<%end if%>
	</div>
	
	<style text="text/css">
		#colorsContainer {background-color:#000000; position:absolute; display:none; visibility:hidden;}
		#colorsInfo {float:left; padding:5px;}
		
		#colorsPreview {width:40px; height:25px; border:solid 1px #FFFFFF; float:left;}
		#colorsNumber {width:50px; height:16px; border:none; color:#FFFFFF; float:left; margin:3px; margin-left:10px;}
		#colorsCancel {width:40px; float:left; border:solid 1px #FFFFFF; margin:2px; margin-left:30px; color:#FFFFFF; text-align:center; cursor:pointer;}
		
		#colorsPalette {float:left; clear:left;}
	</style>
	<div id="colorsContainer">
		<div id="colorsInfo">
			<div id="colorsPreview" style=""></div>
			<div id="colorsNumber" style="">#??????</div>
			<div id="colorsCancel" onclick="this.parentNode.parentNode.style['visibility']='hidden';">关闭</div>
		</div>
		<div id="colorsPalette"></div>
	</div>
		
	
	<!-- #include file="ubbhook.inc" -->
<%end if%>
<%End Sub%>