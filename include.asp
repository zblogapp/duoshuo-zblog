<!-- #include file="function.asp" -->
<!-- #include file="duoshuo_oauth.asp" -->
<!-- #include file="jwt.all.asp" -->

<%
Dim duoshuo_url
duoshuo_url=CBool(InStrRev(LCase(Request.ServerVariables("PATH_INFO").Item),"default.asp"))
'剩余未开发功能：
'1.文章同步到微博（需要API）
'2.未注册用户绑定功能
'3.好像没了（望天）


'注册插件
Call RegisterPlugin("duoshuo","ActivePlugin_duoshuo")
'挂口部分

';

Function ActivePlugin_duoshuo()

	'重定向评论管理
	Call Add_Action_Plugin("Action_Plugin_Admin_Begin","duoshuo.include.redirect()") 
	'重写评论框
	Call Add_Action_Plugin("Action_Plugin_TArticle_Export_CommentPost_Begin","If Level=4 Then Template_Article_CommentPost=duoshuo.show():Exit Function")
	Call Add_Action_Plugin("Action_Plugin_TArticle_Export_CMTandTB_Begin","If duoshuo.checkspider()=False Then Exit Function")
	'异步数据获取
	Call Add_Filter_Plugin("Filter_Plugin_TArticleList_Build_Template","duoshuo_include_footer")
	Call Add_Filter_Plugin("Filter_Plugin_TArticle_Build_Template","duoshuo_include_footer")
	'修正评论数
	Call Add_Filter_Plugin("Filter_Plugin_TArticle_Export_TemplateTags","duoshuo_include_cc_fix")
	'文章同步
	Call Add_Filter_Plugin("Filter_Plugin_PostArticle_Succeed","duoshuo.include.postarticle_succeed")
	'评论同步
	'Call Add_Filter_Plugin("Filter_Plugin_PostComment_Succeed","duoshuo.include.postcomment_succeed")
	
	'注册接口
	If CheckPluginState("RegPage") Then
		If duoshuo.get("duoshuo_userid")<>"undefined" Then
			Dim strDs,strAcc
			strDs=duoshuo.get("duoshuo_userid")
			strAcc=TransferHTML(FilterSQL(duoshuo.get("accesstoken")),"[html-format]")
			Call CheckParameter(strDs,"int",0)
			
			If strDs<>0 Then
				Call Add_Response_Plugin("Response_Plugin_RegPage_End","<input type=""hidden"" value="""&strDs&""" name=""duoshuo_userid""/>")
				Call Add_Response_Plugin("Response_Plugin_RegPage_End","<input type=""hidden"" value="""&strAcc&""" name=""AccessToken""/>")
			End If
		End If
		Call Add_Action_Plugin("Action_Plugin_RegSave_End","Set BlogUser=RegUser:If Duoshuo_SaveReg Then strResponse=""<script language='javascript' type='text/javascript'>alert('恭喜，注册成功。\n欢迎您成为本站一员。\n\n单击确定登陆本站。');location.href="""""&BlogHost&"zb_system/cmd.asp?act=login""""</script>""")
	End If
End Function


Function InstallPlugin_duoshuo()

	'用户激活插件之后的操作
	On Error Resume Next
	Call Duoshuo_Function
	objConn.Execute("SELECT TOP 1 ds_key FROM blog_Plugin_duoshuo")
	'判断是否有duoshuo这个库，有则err_number=0
	If Err.Number<>0 Then
		Call Duoshuo_CreateCmtDB()
		Call Duoshuo_CreateMemDB()
	End If
	Call SetBlogHint(Empty,Empty,True)
	
End Function

Sub Duoshuo_CreateCmtDB()
	If ZC_MSSQL_ENABLE Then
		objConn.Execute("CREATE TABLE [blog_Plugin_duoshuo](ds_ID int identity(1,1) not null primary key,ds_key nvarchar(128) default '',ds_cmtid int default 0)")
	Else
		objConn.Execute("CREATE TABLE [blog_Plugin_duoshuo](ds_ID AutoIncrement primary key,ds_key VARCHAR(128) default """",ds_cmtid int default 0)")
	End If
End Sub

Sub Duoshuo_CreateMemDB()
	If ZC_MSSQL_ENABLE Then
		objConn.Execute("CREATE TABLE [blog_Plugin_duoshuo_Member](ds_ID int identity(1,1) not null primary key,ds_key nvarchar(128) default '',ds_memid int default 0)")
	Else
		objConn.Execute("CREATE TABLE [blog_Plugin_duoshuo_Member](ds_ID AutoIncrement primary key,ds_key VARCHAR(128) default """",ds_memid int default 0)")
		',ds_accesstoken VARCHAR(128) default """")")
	End If
End Sub

Function UnInstallPlugin_duoshuo()

	'用户停用插件之后的操作
	
End Function


Sub Duoshuo_Function
	Dim objFunc
	Set objFunc=New TFunction
	Call GetFunction
	If FunctionMetas.GetValue("duoshuo_recentcomments")=Empty Then
		objfunc.ID=0
		objfunc.Name="多说最新评论"
		objfunc.FileName="Duoshuo_RecentComments"
		objfunc.HtmlID="Duoshuo_RecentComments"
		objfunc.Ftype="div"
		objfunc.Order=20
		objfunc.MaxLi=0
		objfunc.SidebarID=10000
		objfunc.isSystem=False
		objfunc.Source="plugin_duoshuo"
		objfunc.Content="<ul class=""ds-recent-comments"" data-num-items=""10""></ul>"
		objfunc.ViewType="html"
		objfunc.Save
	End If
	If FunctionMetas.GetValue("duoshuo_topthreads")=Empty Then
		objfunc.ID=0
		objfunc.Name="多说最热文章"
		objfunc.FileName="Duoshuo_TopThreads"
		objfunc.HtmlID="Duoshuo_TopThreads"
		objfunc.Ftype="div"
		objfunc.Order=20
		objfunc.MaxLi=0
		objfunc.SidebarID=10000
		objfunc.isSystem=False
		objfunc.Source="plugin_duoshuo"
		objfunc.ViewType="html"
		objfunc.Content="<ul class=""ds-top-threads"" data-range=""weekly"" data-num-items=""10""></ul>"
		objfunc.Save
	End If
End Sub

Function Duoshuo_SaveReg()
	Dim objDS
	Set objDS=New duoshuo_oauth
	objDS.Duoshuo_UserID=Request("duoshuo_userid")
	objDS.AccessToken=Request("accesstoken")
	objDs.ZB_UserID=BlogUser.ID
	Duoshuo_SaveReg=objDs.Post
	Set objDs=Nothing
End Function
%>