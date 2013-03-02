<!-- #include file="function.asp" -->
<!-- #include file="aspjson.asp" -->
<%
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

End Function


Function InstallPlugin_duoshuo()

	'用户激活插件之后的操作
	On Error Resume Next
	Call Duoshuo_Function
	objConn.Execute("SELECT TOP 1 ds_key FROM blog_Plugin_duoshuo")
	'判断是否有duoshuo这个库，有则err_number=0
	If Err.Number<>0 Then
		If ZC_MSSQL_ENABLE Then
			objConn.Execute("CREATE TABLE [blog_Plugin_duoshuo](ds_ID int identity(1,1) not null primary key,ds_key nvarchar(128) default """",ds_cmtid int default 0)")
		Else
			objConn.Execute("CREATE TABLE [blog_Plugin_duoshuo](ds_ID AutoIncrement primary key,ds_key VARCHAR(128) default """",ds_cmtid int default 0)")
		End If
	End If
	Call SetBlogHint(Empty,Empty,True)
	
End Function


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
%>