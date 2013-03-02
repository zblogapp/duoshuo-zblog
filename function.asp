<%
Sub duoshuo_Initialize()
	Set duoshuo.config=New TConfig
	duoshuo.config.Load "DuoShuo"
	If duoshuo.config.Read("ver")="" Then
		duoshuo.config.Write "ver","1.0"
		duoshuo.config.Write "duoshuo_api_hostname","api.duoshuo.com"
		duoshuo.config.Write "duoshuo_cron_sync_enabled","async"
		duoshuo.config.Write "duoshuo_cc_fix","False"
		duoshuo.config.Save
	End If
End Sub
'****************************************
' duoshuo 子菜单
'****************************************
Function duoshuo_SubMenu(id)
	If id="setting" Then id=1
	Dim aryName,aryPath,aryFloat,aryInNewWindow,aryS,i
	aryName=Array("首页","设置","导出","更多")
	aryPath=Array("main.asp","main.asp?act=setting","export.asp",IIf(duoshuo.config.Read("short_name")="","http://www","http://"&duoshuo.config.Read("short_name"))&".duoshuo.com")
	aryFloat=Array("m-left","m-left","m-left","m-right")
	aryS=Array(Not(duoshuo.config.Read("short_name")="" Or duoshuo.get("submenu")="false"),True,True,True)
	aryInNewWindow=Array(False,False,False,True)
	For i=0 To Ubound(aryName)
		duoshuo_SubMenu=duoshuo_SubMenu & IIf(aryS(i),MakeSubMenu(aryName(i),aryPath(i),aryFloat(i)&IIf(i=id," m-now",""),aryInNewWindow(i)),"")
	Next
End Function
'****************************************
' 加入异步
'****************************************
Sub duoshuo_include_async()
	If duoshuo.config.Read("duoshuo_cron_sync_enabled")="async" Then
		duoshuo.include.footdata=duoshuo.include.footdata&"<script language=""javascript"" type=""text/javascript"" src="""&BlogHost&"zb_users/plugin/duoshuo/noresponse.asp?act=api_async&"&Rnd&"""></script>"
	End If
End Sub


'****************************************
' 修正评论数
'****************************************
Function duoshuo_include_cc_fix(ByRef aryTemplateTagsName, ByRef aryTemplateTagsValue)
	duoshuo_Initialize()
	If duoshuo.config.Read("duoshuo_cc_fix")="True" Then
		aryTemplateTagsValue(7)="<span id='duoshuo_comment"&aryTemplateTagsValue(1)&"'></span>"
		If duoshuo.threadkey="" Then
			duoshuo.threadkey=aryTemplateTagsValue(1) 
			duoshuo.include.footdata=duoshuo.include.footdata&"<script type='text/javascript' src='http://api.duoshuo.com/threads/counts.jsonp?short_name="& Server.URLEncode(duoshuo.config.Read("short_name")) &"&threads="&Server.URLEncode(duoshuo.threadkey)&"&callback=duoshuo_callback'></script>"  '插入页面底部的版权信息，进行批量获
		Else
			duoshuo.threadkey=duoshuo.threadkey&","&aryTemplateTagsValue(1)
		End If
	End If
End Function
'****************************************
' 写入footer
'****************************************
Function duoshuo_include_footer(ByRef html)
	Stop
	duoshuo_Initialize()
	Call duoshuo_include_async()
	duoshuo.include.footdata="<script type='text/javascript'>function duoshuo_callback(data){if(data.response){for(var i in data.response){$('#duoshuo_comment'+i).html(data.response[i].comments);}}};var duoshuoQuery = {short_name:"""&duoshuo.config.Read("short_name")&"""};</script><script type=""text/javascript"" src=""http://static.duoshuo.com/embed.js""></script>"&duoshuo.include.footdata
	html=Replace(html,"<#ZC_BLOG_COPYRIGHT#>",duoshuo.include.footdata&"<#ZC_BLOG_COPYRIGHT#>")
End Function
%>

<script language="javascript" runat="server">
var duoshuo={}
duoshuo.get=function(s){return Request.QueryString(s).Item}
duoshuo.post=function(s){return Request.Form(s).Item}
duoshuo.config=function(){}
duoshuo.include={
	redirect:function(){
		if(duoshuo.get("act")=="CommentMng") Response.Redirect(BlogHost + "zb_users/plugin/duoshuo/main.asp?submenu=false")
	}	
}
duoshuo.show=function(){
	var k="";
	duoshuo_Initialize();
	k+='<!'+'-- Duoshuo Comment BEGIN -'+'->';
	k+='<div class="ds-thread" data-category="<#article/category/id#>" data-thread-key="<#article/id#>" ';
	k+='data-title="<#article/title#>" data-author-key="<#article/author/id#>" data-url=""></div>';
	k+='<!-'+'- Duoshuo Comment END -->';
	return k;
}
duoshuo.threadkey=""
duoshuo.include.footdata="";
</script>