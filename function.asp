<%
Dim bduoshuo_Initialize
bduoshuo_Initialize=False
Sub duoshuo_Initialize()
	If bduoshuo_Initialize Then Exit Sub
	Set duoshuo.config=New TConfig
	duoshuo.config.Load "DuoShuo"
	If duoshuo.config.Read("ver")="" Then
		duoshuo.config.Write "ver","1.0"
		duoshuo.config.Write "duoshuo_api_hostname","api.duoshuo.com"
		duoshuo.config.Write "duoshuo_cron_sync_enabled","async"
		duoshuo.config.Write "duoshuo_cc_fix","False"
		duoshuo.config.Write "duoshuo_comments_wrapper_intro",""
		duoshuo.config.Write "duoshuo_comments_wrapper_outro",""
		duoshuo.config.Write "duoshuo_seo_enabled","False"
		duoshuo.config.Save
	End If
	bduoshuo_Initialize=True
End Sub
'****************************************
' duoshuo 子菜单
'****************************************
Function duoshuo_SubMenu(id)
	If id="setting" Then id=2
	If id="personal" Then id=1
	If id="export" Then id=3
	Dim aryName,aryPath,aryFloat,aryInNewWindow,i
	aryName=Array("评论管理","多说设置","高级选项","数据导出","多说后台")
	aryPath=Array("main.asp","main.asp?act=personal","main.asp?act=setting","export.asp",IIf(duoshuo.config.Read("short_name")="","http://www","http://"&duoshuo.config.Read("short_name"))&".duoshuo.com")
	aryFloat=Array("m-left","m-left","m-left","m-left","m-right")
	aryInNewWindow=Array(False,False,False,False,True)
	For i=0 To Ubound(aryName)
		duoshuo_SubMenu=duoshuo_SubMenu & MakeSubMenu(aryName(i),aryPath(i),aryFloat(i)&IIf(i=id," m-now",""),aryInNewWindow(i))
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
		aryTemplateTagsValue(7)="<span id='duoshuo_comment"&aryTemplateTagsValue(1)&"' duoshuo_id="""&aryTemplateTagsValue(1)&"""></span>"
		If duoshuo.threadkey="" Then
			duoshuo.threadkey=aryTemplateTagsValue(1) 
		Else
			duoshuo.threadkey=duoshuo.threadkey&","&aryTemplateTagsValue(1)
		End If
	End If
End Function
'****************************************
' 写入footer
'****************************************
Function duoshuo_include_footer(ByRef html)
	duoshuo.include.footdata=""
	duoshuo_Initialize()
	Call duoshuo_include_async()
	If duoshuo.threadkey<>"" Then 
		duoshuo.include.footdata=duoshuo.include.footdata&"<script type='text/javascript' src='http://api.duoshuo.com/threads/counts.jsonp?short_name="& Server.URLEncode(duoshuo.config.Read("short_name")) &"&threads="&Server.URLEncode(duoshuo.threadkey)&"&callback=duoshuo_callback'></script>"  '插入页面底部的版权信息，进行批量获
	End If
	duoshuo.include.footdata="<script type='text/javascript'>function duoshuo_callback(data){if(data.response){for(var i in data.response){jQuery('[duoshuo_id='+i+']').html(data.response[i].comments);}}};var duoshuoQuery = {short_name:"""&duoshuo.config.Read("short_name")&"""};</script><script type=""text/javascript"" src=""http://static.duoshuo.com/embed.js""></script>"&duoshuo.include.footdata
	'为了不和Z-Blog插件YTCMS冲突
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
		if(duoshuo.get("act")=="CommentMng") Response.Redirect(BlogHost + "zb_users/plugin/duoshuo/main.asp")
	}	
}
duoshuo.show=function(){
	var k="";
	duoshuo_Initialize();
	k+='<!'+'-- Duoshuo Comment BEGIN -'+'->';
	k+=duoshuo.config.Read("duoshuo_comments_wrapper_intro");
	k+='<div class="ds-thread" data-thread-key="<#article/id#>" ';
	k+= 'data-title="<#article/title#>" data-author-key="<#article/author/id#>" data-url="<#article/url#>"></div>';	k+=duoshuo.config.Read("duoshuo_comments_wrapper_outro");
	k+='<!-'+'- Duoshuo Comment END -->';
	return k;
}
duoshuo.threadkey=""
duoshuo.include.footdata="";
duoshuo.checkspider=function(){
	duoshuo_Initialize();
	if(duoshuo.config.Read("duoshuo_seo_enabled")!="True"){return false}
	if(ZC_POST_STATIC_MODE=="STATIC"){return false}
	var spider=/(baidu|google|bing|soso|360|Yahoo|msn|Yandex|youdao|mj12|Jike|Ahrefs|ezooms|Easou|sogou)(bot|spider|Transcoder|slurp)/i
	if(spider.test(Request.ServerVariables("HTTP_USER_AGENT").Item)){
		return true
	}
	else{
		return false
	}
}
</script>