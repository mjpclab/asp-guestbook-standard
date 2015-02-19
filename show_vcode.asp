<!-- #include file="config.asp" -->
<%
'Verify code bitmap display module
'(C) Copyright 2004-2005 MJ PC Lab
'http://mjpclab.net/

Response.ContentType="image/bmp"
Response.Expires = -1
Response.AddHeader "Pragma","no-cache"
Response.AddHeader "cache-control","no-cache, must-revalidate"

i_vcode=cstr(session("vcode"))		'the code to display
i_padding=2
i_offsetrange=4
i_textcolor=right(FormColor,6)	'fore color RGB
i_bgcolor=right(FormBGC,6)		'bg color RGB
i_noisecolor=i_textcolor
i_noisefrequency=18

'Construct number pixels' grid array
'===============================================
alpha_width=7
alpha_height=9

dim alpha_px(9,8)

'Num 0 grid
alpha_px(0,0)=62
alpha_px(0,1)=99
alpha_px(0,2)=99	'115
alpha_px(0,3)=99	'123
alpha_px(0,4)=99	'111
alpha_px(0,5)=99	'103
alpha_px(0,6)=99	'99
alpha_px(0,7)=99	'99
alpha_px(0,8)=62

'Num 1 grid
alpha_px(1,0)=24
alpha_px(1,1)=28
alpha_px(1,2)=30
alpha_px(1,3)=24
alpha_px(1,4)=24
alpha_px(1,5)=24
alpha_px(1,6)=24
alpha_px(1,7)=24
alpha_px(1,8)=126

'Num 2 grid
alpha_px(2,0)=62
alpha_px(2,1)=99
alpha_px(2,2)=96
alpha_px(2,3)=48
alpha_px(2,4)=24
alpha_px(2,5)=12
alpha_px(2,6)=6
alpha_px(2,7)=99
alpha_px(2,8)=127

'Num 3 grid
alpha_px(3,0)=62
alpha_px(3,1)=99
alpha_px(3,2)=96
alpha_px(3,3)=96
alpha_px(3,4)=60
alpha_px(3,5)=96
alpha_px(3,6)=96
alpha_px(3,7)=99
alpha_px(3,8)=62

'Num 4 grid
alpha_px(4,0)=48
alpha_px(4,1)=56
alpha_px(4,2)=60
alpha_px(4,3)=54
alpha_px(4,4)=51
alpha_px(4,5)=127
alpha_px(4,6)=48
alpha_px(4,7)=48
alpha_px(4,8)=120

'Num 5 grid
alpha_px(5,0)=127
alpha_px(5,1)=3
alpha_px(5,2)=3
alpha_px(5,3)=63
alpha_px(5,4)=99
alpha_px(5,5)=96
alpha_px(5,6)=96
alpha_px(5,7)=99
alpha_px(5,8)=62

'Num 6 grid
alpha_px(6,0)=28
alpha_px(6,1)=6
alpha_px(6,2)=3
alpha_px(6,3)=3
alpha_px(6,4)=63
alpha_px(6,5)=99
alpha_px(6,6)=99
alpha_px(6,7)=99
alpha_px(6,8)=62

'Num 7 grid
alpha_px(7,0)=127
alpha_px(7,1)=99
alpha_px(7,2)=96
alpha_px(7,3)=48
alpha_px(7,4)=24
alpha_px(7,5)=12
alpha_px(7,6)=12
alpha_px(7,7)=12
alpha_px(7,8)=12

'Num 8 grid
alpha_px(8,0)=62
alpha_px(8,1)=99
alpha_px(8,2)=99
alpha_px(8,3)=99
alpha_px(8,4)=62
alpha_px(8,5)=99
alpha_px(8,6)=99
alpha_px(8,7)=99
alpha_px(8,8)=62

'Num 9 grid
alpha_px(9,0)=62
alpha_px(9,1)=99
alpha_px(9,2)=99
alpha_px(9,3)=99
alpha_px(9,4)=126
alpha_px(9,5)=96
alpha_px(9,6)=96
alpha_px(9,7)=48
alpha_px(9,8)=30
'===============================================


'analyze options
'===============================================
if i_textcolor="" then i_textcolor="000000"
if i_bgcolor="" then i_bgcolor="FFFFFF"
if i_noisecolor="" then i_noisecolor="000000"

i_textcolor_red=cbyte("&H" & mid(i_textcolor,1,2))
i_textcolor_green=cbyte("&H" & mid(i_textcolor,3,2))
i_textcolor_blue=cbyte("&H" & mid(i_textcolor,5,2))

i_bgcolor_red=cbyte("&H" & mid(i_bgcolor,1,2))
i_bgcolor_green=cbyte("&H" & mid(i_bgcolor,3,2))
i_bgcolor_blue=cbyte("&H" & mid(i_bgcolor,5,2))

i_noisecolor_red=cbyte("&H" & mid(i_noisecolor,1,2))
i_noisecolor_green=cbyte("&H" & mid(i_noisecolor,3,2))
i_noisecolor_blue=cbyte("&H" & mid(i_noisecolor,5,2))

i_width=(len(i_vcode)+1)*i_padding + len(i_vcode)*alpha_width + len(i_vcode)*2*i_offsetrange
i_height=2*i_padding + alpha_height + 2*i_offsetrange
'===============================================



'Create BMP data stream & bmp header(core code)
'Copyright MJ PC Lab 2004-2005 [http://mjpclab.net]
'All rights reserved
'===============================================
dim i,i1,j,row,k,bitpos,bmpstream,tempstream,curNum,zerocount,rndtemp
dim offsetx(),offsety()
redim offsetx(len(i_vcode)-1),offsety(len(i_vcode)-1)

bmpstream=""
tempstream=""
zerocount=(4-(i_width*3 mod 4)) mod 4
randomize

for i=lbound(offsetx) to ubound(offsetx)		'Create alpha offset random distance
	offsetx(i)=fix((i_offsetrange*2+1)*rnd)
	offsety(i)=fix((i_offsetrange*2+1)*rnd)
next

'Draw bottom padding
'***********************************
for i=1 to i_padding
	tempstream=""
	for j=1 to i_width
		tempstream = tempstream & createbgcolor()
	next
	bmpstream=bmpstream & tempstream & getstr(0,zerocount)
next
'***********************************


for row=i_offsetrange*2+alpha_height-1 to 0 step -1
	tempstream=""

	for k=0 to len(i_vcode)-1
		curNum=cint(mid(i_vcode,k+1,1))
		
		'Draw left padding
		'***********************************
		for j=1 to i_padding
			tempstream = tempstream & createbgcolor()
		next
		'***********************************
		
		
		'Draw number
		'***********************************
		for j=1 to offsetx(k)		'Draw random offset X(left)
			tempstream = tempstream & createbgcolor()			
		next
		
		
		if row+1<=offsety(k) then	'Draw random offset Y(top)
			for j=1 to alpha_width
				tempstream = tempstream & createbgcolor()
			next
		elseif row>=offsety(k)+alpha_height then	'Draw random offset Y(bottom) [row+1>=offsety(k)+alpha_height+1]
			for j=1 to alpha_width
				tempstream = tempstream & createbgcolor()
			next
		elseif row-offsety(k)>ubound(alpha_px,2) then	'if subscript is too early to over
			for j=1 to alpha_width
				tempstream = tempstream & createbgcolor()
			next
		else						'Draw number pixels
			for bitpos=0 to alpha_width-1
				if clng(alpha_px(curNum,row-offsety(k)) and clng(2^bitpos)) <> 0 then
					rndtemp=rnd
					tempstream=tempstream & chrB(i_textcolor_blue) & chrB(i_textcolor_green) & chrB(i_textcolor_red)
				else
					tempstream = tempstream & createbgcolor()
				end if
			next
		end if

		for j=1 to i_offsetrange*2-offsetx(k)		'Draw random offset X(right)
			tempstream = tempstream & createbgcolor()			
		next
		'***********************************
	next


	'Draw right padding
	'***********************************
	for j=1 to i_padding
		tempstream = tempstream & createbgcolor()
	next
	'***********************************

	bmpstream=bmpstream & tempstream & getstr(0,zerocount)
next


'Draw top padding
'***********************************
for i=1 to i_padding
	tempstream=""
	for j=1 to i_width
		tempstream = tempstream & createbgcolor()
	next
	bmpstream=bmpstream & tempstream & getstr(0,zerocount)
next
'***********************************


dim header
header=""

header=chrB(&H42) & chrB(&H4D)		'位图类型 BM
header=header & getstr(&H36 + lenB(bmpstream),4)	'文件长度(header+bitstream)
header=header & chrB(0) & chrB(0) & chrB(0) & chrB(0)	'保留
header=header & getstr(&H36,4)		'实际数据偏移
header=header & getstr(&H28,4)		'位图信息头长度(win 3.x/9x/NT)
header=header & getstr(i_width,4)		'位图宽度
header=header & getstr(i_height,4)		'位图高度
header=header & getstr(1,2)		'位面数(常数)
header=header & getstr(24,2)	'每象素位数
header=header & chrB(0) & chrB(0) & chrB(0) & chrB(0)	'压缩：不压缩
header=header & getstr(lenB(bmpstream),4)	'位图数据大小,舍入为4的倍数
header=header & chrB(&H12) & chrB(&H0B) & chrB(0) & chrB(0)	'水平分辨率：象素/米
header=header & chrB(&H12) & chrB(&H0B) & chrB(0) & chrB(0)	'垂直分辨率：象素/米
header=header & chrB(0) & chrB(0) & chrB(0) & chrB(0)	'使用颜色数(0=全部)
header=header & chrB(0) & chrB(0) & chrB(0) & chrB(0)	'重要颜色数(0=全部)
'=====================================================

Response.BinaryWrite header & bmpstream



'-----------------------------------------------------
function createbgcolor
	dim rvalue
	randomize
	
	if fix(abs(rnd*i_noisefrequency))=i_noisefrequency-1 then	'Draw noise
		rvalue = chrB(i_noisecolor_blue) & chrB(i_noisecolor_green) & chrB(i_noisecolor_red)
	else
		rvalue = chrB(i_bgcolor_blue) & chrB(i_bgcolor_green) & chrB(i_bgcolor_red)
	end if
	
	createbgcolor=rvalue
end function
'-----------------------------------------------------
function getstr(num,bytecount)
if bytecount>0 then
	dim rvalue,strhex,t_i
	rvalue=""
	strhex=hex(num)
	strhex=string(bytecount*2-len(strhex),"0") & strhex
	
	for t_i = 1 to len(strhex) step 2
		rvalue=chrB(cbyte("&H" & mid(strhex,t_i,2))) & rvalue
	next
	
	getstr=rvalue
else
	getstr=""
end if
end function
%>