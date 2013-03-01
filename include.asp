<!-- #include file="function.asp" -->
<!-- #include file="aspjson.asp" -->
<%
'注册插件
Call RegisterPlugin("duoshuo","ActivePlugin_duoshuo")
'挂口部分
Function ActivePlugin_duoshuo()

	Call Add_Action_Plugin("Action_Plugin_Admin_Begin","duoshuo.include.redirect()") 
	Call Add_Action_Plugin("Action_Plugin_TArticle_Export_Begin","If Level=4 Then Disable_Export_CMTandTB=True:Disable_Export_CommentPost=True:Template_Article_CommentPost=duoshuo.show():HasCMTandTB=True")
	Call Add_Filter_Plugin("Filter_Plugin_TArticleList_Build_TemplateTags","duoshuo_include_async")
	Call Add_Filter_Plugin("Filter_Plugin_TArticle_Build_TemplateTags","duoshuo_include_async")

End Function


Function InstallPlugin_duoshuo()

	'用户激活插件之后的操作
	On Error Resume Next
	objConn.Execute("SELECT TOP 1 ds_key FROM blog_Plugin_duoshuo")
	'判断是否有duoshuo这个库，有则err_number=0
	If Err.Number<>0 Then
		If ZC_MSSQL_ENABLE Then
			objConn.Execute("CREATE TABLE [blog_Plugin_duoshuo](ds_ID int identity(1,1) not null primary key,ds_key nvarchar(128) default """",ds_cmtid int default 0)")
		Else
			objConn.Execute("CREATE TABLE [blog_Plugin_duoshuo](ds_ID AutoIncrement primary key,ds_key VARCHAR(128) default """",ds_cmtid int default 0)")
		End If
	End If
	
End Function


Function UnInstallPlugin_duoshuo()

	'用户停用插件之后的操作
	
End Function
%>