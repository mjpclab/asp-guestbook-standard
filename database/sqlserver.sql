--SQL Server初始化脚本，使用表名前缀gbook_
--也可以使用管理工具自动初始化数据库（可自定义表名前缀）

CREATE TABLE [dbo].[gbook_config](id INT IDENTITY(1,1) PRIMARY KEY,status INT DEFAULT 239 NOT NULL,ipconstatus TINYINT DEFAULT 0,homelogo VARCHAR(127),homename NVARCHAR(30),homeaddr VARCHAR(127),adminhtml TINYINT DEFAULT 6 NOT NULL,guesthtml TINYINT DEFAULT 6 NOT NULL,admintimeout SMALLINT DEFAULT 0 NOT NULL,showip TINYINT DEFAULT 2 NOT NULL,adminshowip TINYINT DEFAULT 4 NOT NULL,adminshoworiginalip TINYINT DEFAULT 0 NOT NULL,vcodecount TINYINT DEFAULT 68 NOT NULL,mailflag TINYINT DEFAULT 0 NOT NULL,mailreceive VARCHAR(48),mailfrom VARCHAR(48),mailsmtpserver VARCHAR(48),mailuserid NVARCHAR(48),mailuserpass NVARCHAR(48),maillevel TINYINT DEFAULT 3 NOT NULL,cssfontfamily NVARCHAR(48) DEFAULT N'宋体',cssfontsize NVARCHAR(8) DEFAULT '12px',csslineheight NVARCHAR(8) DEFAULT '150%',tablewidth NVARCHAR(5) DEFAULT '630',tableleftwidth NVARCHAR(5) DEFAULT '150',windowspace TINYINT DEFAULT 20 NOT NULL,leavecontentheight TINYINT DEFAULT 7 NOT NULL,searchtextwidth TINYINT DEFAULT 20 NOT NULL,replytextheight TINYINT DEFAULT 10 NOT NULL,itemsperpage SMALLINT DEFAULT 5 NOT NULL,titlesperpage SMALLINT DEFAULT 20 NOT NULL,picturesperrow TINYINT DEFAULT 5 NOT NULL,frequentfacecount TINYINT DEFAULT 15 NOT NULL,servertimezoneoffset SMALLINT DEFAULT 0 NOT NULL,displaytimezoneoffset SMALLINT DEFAULT 0 NOT NULL,visualflag INT DEFAULT 107 NOT NULL,advpagelistcount TINYINT DEFAULT 10 NOT NULL,ubbflag INT DEFAULT 2047 NOT NULL,pagecontrol TINYINT DEFAULT 63 NOT NULL,wordslimit INT DEFAULT 0 NOT NULL,delconfirm INT DEFAULT 9 NOT NULL,styleid INT)
CREATE TABLE [dbo].[gbook_filterconfig](filterid INT IDENTITY(1,1) PRIMARY KEY,filtersort INT DEFAULT 0,regexp NTEXT,filtermode SMALLINT DEFAULT 0,replacestr NTEXT,[memo] NVARCHAR(25))
CREATE TABLE [dbo].[gbook_floodconfig](id INT IDENTITY(1,1) PRIMARY KEY,minwait INT DEFAULT 0,searchrange INT DEFAULT 0,searchflag INT DEFAULT 32)
CREATE TABLE [dbo].[gbook_ipv4config](listid INT IDENTITY(1,1) PRIMARY KEY,cfgtype TINYINT NOT NULL,ipfrom VARCHAR(8) NOT NULL,ipto VARCHAR(8) NOT NULL)
CREATE TABLE [dbo].[gbook_ipv6config](listid INT IDENTITY(1,1) PRIMARY KEY,cfgtype TINYINT NOT NULL,ipfrom VARCHAR(32) NOT NULL,ipto VARCHAR(32) NOT NULL)
CREATE TABLE [dbo].[gbook_main](id INT IDENTITY(1,1) NOT FOR REPLICATION PRIMARY KEY,parent_id INT DEFAULT 0,root_id AS CASE WHEN parent_id<=0 THEN id ELSE parent_id END,name NVARCHAR(20),title NVARCHAR(30),email VARCHAR(50),qqid NVARCHAR(16),msnid VARCHAR(50),homepage VARCHAR(127),logdate DATETIME,lastupdated DATETIME,ipv4addr VARCHAR(15),ipv6addr VARCHAR(39),originalipv4 VARCHAR(15),originalipv6 VARCHAR(39),faceid TINYINT DEFAULT 0,guestflag INT DEFAULT 0,whisperpwd VARCHAR(32),article NTEXT,replied TINYINT DEFAULT 0)
CREATE INDEX [IX_forcount] ON [dbo].[gbook_main](parent_id ASC)
CREATE INDEX [IX_main] ON [dbo].[gbook_main](parent_id ASC,lastupdated DESC,id DESC,guestflag ASC)
CREATE INDEX [IX_root_id] ON [dbo].[gbook_main](root_id ASC)
CREATE TABLE [dbo].[gbook_reply](articleid INT DEFAULT 0 PRIMARY KEY,replydate DATETIME,htmlflag TINYINT DEFAULT 0,reinfo NTEXT)
CREATE TABLE [dbo].[gbook_stats](id INT IDENTITY(1,1) PRIMARY KEY,startdate DATETIME,[view] INT DEFAULT 0,search INT DEFAULT 0,leaveword INT DEFAULT 0,written INT DEFAULT 0,filtered INT DEFAULT 0,banned INT DEFAULT 0,login INT DEFAULT 0,loginfailed INT DEFAULT 0)
CREATE TABLE [dbo].[gbook_stats_clientinfo](id INT IDENTITY(1,1) PRIMARY KEY,os NVARCHAR(25),browser NVARCHAR(16),screenwidth INT DEFAULT 0,screenheight INT DEFAULT 0,timesect DATETIME DEFAULT '1999-12-31',sourceaddr NVARCHAR(50),fullsource NVARCHAR(255))
CREATE TABLE [dbo].[gbook_style](styleid INT IDENTITY(1,1) PRIMARY KEY,stylename NVARCHAR(5),themepath VARCHAR(31))
CREATE TABLE [dbo].[gbook_supervisor](id INT IDENTITY(1,1) PRIMARY KEY,adminpass VARCHAR(32),name NVARCHAR(20),faceid TINYINT DEFAULT 0,faceurl VARCHAR(127),email VARCHAR(50),qqid NVARCHAR(16),msnid VARCHAR(50),homepage VARCHAR(127),declareflag TINYINT DEFAULT 0,[declare] NTEXT)
--初始化配色方案：
INSERT INTO [dbo].[gbook_style](stylename,themepath) VALUES (N'(默认)','default')
INSERT INTO [dbo].[gbook_style](stylename,themepath) VALUES (N'橘红','orange')
INSERT INTO [dbo].[gbook_style](stylename,themepath) VALUES (N'湛蓝','lightblue')
INSERT INTO [dbo].[gbook_style](stylename,themepath) VALUES (N'粉红','pink')
INSERT INTO [dbo].[gbook_style](stylename,themepath) VALUES (N'翠绿','lightgreen')
INSERT INTO [dbo].[gbook_style](stylename,themepath) VALUES (N'蓝灰','grayblue')
INSERT INTO [dbo].[gbook_style](stylename,themepath) VALUES (N'蓝黑','darkblue')
INSERT INTO [dbo].[gbook_style](stylename,themepath) VALUES (N'银灰','silver')
INSERT INTO [dbo].[gbook_style](stylename,themepath) VALUES (N'黑夜','black')
INSERT INTO [dbo].[gbook_style](stylename,themepath) VALUES (N'节日气氛','festival')
--初始化参数设置：
INSERT INTO [dbo].[gbook_config](status,ipconstatus,homelogo,homename,homeaddr,adminhtml,guesthtml,admintimeout,showip,adminshowip,adminshoworiginalip,vcodecount,mailflag,mailreceive,mailfrom,mailsmtpserver,mailuserid,mailuserpass,maillevel,cssfontfamily,cssfontsize,csslineheight,tablewidth,tableleftwidth,windowspace,leavecontentheight,searchtextwidth,replytextheight,itemsperpage,titlesperpage,picturesperrow,frequentfacecount,servertimezoneoffset,displaytimezoneoffset,visualflag,advpagelistcount,ubbflag,pagecontrol,wordslimit,delconfirm,styleid)VALUES(239,0,'','MyHomePage','',6,6,20,34,132,132,68,0,'','','','','',3,N'宋体','12px','150%',630,150,20,7,20,10,5,20,5,15,0,0,107,10,2047,63,0,9,6)
INSERT INTO [dbo].[gbook_floodconfig](minwait,searchrange,searchflag) VALUES(0,0,32)
--【初始化管理员密码：8888】
INSERT INTO [dbo].[gbook_supervisor](adminpass,name,faceid,faceurl,email,qqid,msnid,homepage,declareflag,[declare]) VALUES('cf79ae6addba60ad018347359bd144d2','Admin',0,'','','','','',1,'')
INSERT INTO [dbo].[gbook_main](parent_id,name,title,logdate,lastupdated,ipv4addr,faceid,guestflag,article,replied)VALUES(0,'MJ PC Lab',N'欢迎使用申江ASP留言本标准版','2004-8-2 18:18:18','2004-8-2 18:18:18','127.0.0.1',1,1,N'欢迎使用申江ASP留言本标准版,管理员初始密码<span style="color:red;font-weight:bold;">8888</span>,请及时更改.',0)
