<%
Sub duoshuo_Initialize()
	Set duoshuo.config=New TConfig
	duoshuo.config.Load "DuoShuo"
	If duoshuo.config.Read("ver")="" Then
		duoshuo.config.Write "ver","1.0"
		duoshuo.config.Save
	End If
End Sub
'****************************************
' duoshuo 子菜单
'****************************************
Function duoshuo_SubMenu(id)
	If id="setting" Then id=1
	Dim aryName,aryPath,aryFloat,aryInNewWindow,aryS,i
	aryName=Array("首页","设置","更多")
	aryPath=Array("main.asp","main.asp?act=setting","http://"&duoshuo.config.Read("short_name")&".duoshuo.com")
	aryFloat=Array("m-left","m-left","m-right")
	aryS=Array(Not(duoshuo.config.Read("short_name")="" Or duoshuo.get("submenu")="false"),True,True)
	aryInNewWindow=Array(False,False,True)
	For i=0 To Ubound(aryName)
		duoshuo_SubMenu=duoshuo_SubMenu & IIf(aryS(i),MakeSubMenu(aryName(i),aryPath(i),aryFloat(i)&IIf(i=id," m-now",""),aryInNewWindow(i)),"")
	Next
End Function
'****************************************
' 加入异步
'****************************************
Function duoshuo_include_async(ByRef aryTemplateTagsName, ByRef aryTemplateTagsValue)
	duoshuo_Initialize()
	If duoshuo.config.Read("duoshuo_cron_sync_enabled")="async" Then
		Dim i,j
		j=UBound(aryTemplateTagsName)
		For i=1 to j
			If aryTemplateTagsName(i)="ZC_BLOG_COPYRIGHT" Then
				Randomize
				aryTemplateTagsValue(i)="<script language=""javascript"" type=""text/javascript"" src="""&ZC_BLOG_HOST&"zb_users/plugin/duoshuo/noresponse.asp?act=api_async"&Rnd&"""></script>" & aryTemplateTagsValue(i)
			End If
		Next
	End If
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
	k+='<scri'+'pt type="text/javascript">';
	k+='var duoshuoQuery = {"short_name":"'+duoshuo.config.Read("short_name")+'"};';
	k+='(function() {';
	k+='	var ds = document.createElement("script");';
	k+="	ds.type = 'text/javascript';ds.async = true;";
	k+="	ds.src = 'http://static.duoshuo.com/embed.js';";
	k+="	(document.getElementsByTagName('head')[0] || document.getElementsByTagName('body')[0]).appendChild(ds);";
	k+='})();';
	k+='</'+'script><!-'+'- Duoshuo Comment END -->';
	return k;
}

</script>