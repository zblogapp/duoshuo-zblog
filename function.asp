<%
Dim bduoshuo_Initialize
bduoshuo_Initialize=False
Sub duoshuo_Initialize()
	If bduoshuo_Initialize Then Exit Sub
	Set duoshuo.config=New TConfig
	duoshuo.config.Load "DuoShuo"
	If duoshuo.config.Read("ver")="" Then
		duoshuo.config.Write "ver","1.2"
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
	aryName=Array("评论管理","多说设置","高级选项","导入导出","多说后台")
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

<script language="javascript" runat="server" >
var duoshuo={}
//多说API地址
duoshuo.url={
	posts:{
		"import":"/posts/import.json", //同步评论到多说
		"create":"/posts/create.json" //发表评论
	},
	threads:{
		"count":"/threads/counts.json", //评论数转发数
		"import":"/threads/import.json" //同步文章到多说
	},
	log:{
		"list":"/log/list.json" //同步回数据库
	},
	users:{
		"import":"/users/import.json" //用户数据同步
	}
}
//常用函数
//HTTP GET请求
duoshuo.get=function(s){return Request.QueryString(s).Item}
//HTTP POST请求
duoshuo.post=function(s){return Request.Form(s).Item}
//判断是否蜘蛛
duoshuo.checkspider=function(){
	duoshuo_Initialize();
	if(duoshuo.config.Read("duoshuo_seo_enabled")!="True"){return false}
	if(ZC_POST_STATIC_MODE=="STATIC"){return true}
	var spider=/(baidu|google|bing|soso|360|Yahoo|msn|Yandex|youdao|mj12|Jike|Ahrefs|ezooms|Easou|sogou)(bot|spider|Transcoder|slurp)/i
	if(spider.test(Request.ServerVariables("HTTP_USER_AGENT").Item)){
		return true
	}
	else{
		return false
	}
}
//日期处理
duoshuo.date=function(d){
	//Microsoft JScript for ASP不支持new Date("xxxTxxx")
	this.date=d;//"2012-12-21T12:00+0800";
	this.getMonth=function(){return this.date.split("T")[0].split("-")[1]}
	this.getDay=function(){return this.date.split("T")[0].split("-")[2]}
	this.getFullYear=function(){return this.date.split("T")[0].split("-")[0]}
	this.getHours=function(){return this.date.split("T")[1].split(":")[0]}
	this.getMinutes=function(){return this.date.split("T")[1].split(":")[1]}
	this.getSeconds=function(){return this.date.split("T")[1].split(":")[2].split("+")[0]}
	
}

//挂口操作
duoshuo.include={
	redirect:function(){
		if(duoshuo.get("act")=="CommentMng") Response.Redirect(BlogHost + "zb_users/plugin/duoshuo/main.asp")
	},
	postarticle_succeed:function(obj){
		var odata=new Array(8);
		odata[0]="threads[0][thread_key]="+obj.ID;
		odata[1]="threads[0][title]="+Server.URLEncode(obj.Title);
		odata[2]="threads[0][url]="+Server.URLEncode(obj.FullUrl);
		odata[3]="threads[0][content]="
		odata[4]="threads[0][author_key]="+obj.AuthorID;
		odata[5]="threads[0][excerpt]="+Server.URLEncode(obj.HtmlIntro);
		odata[6]="threads[0][comment_status]=open";
		odata[7]="threads[0][likes]=0";
		odata[8]="threads[0][views]="+obj.ViewNums;
		var objXmlHttp=new ActiveXObject("MSXML2.ServerXMLHTTP");
		objXmlHttp.open("POST","http://" + duoshuo.config.Read("duoshuo_api_hostname") + duoshuo.url.threads['import'])
		objXmlHttp.setRequestHeader("Content-Type","application/x-www-form-urlencoded")
		objXmlHttp.send("short_name=" + Server.URLEncode(duoshuo.config.Read("short_name")) + "&secret=" + Server.URLEncode(duoshuo.config.Read("secret")) + "&" + odata.join("&"))
		objXmlHttp=null;
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

//临时变量
duoshuo.config=function(){}
duoshuo.include.footdata="";
duoshuo.threadkey="";


//API处理函数
duoshuo.api={}
duoshuo.api.create = function(meta_json,log_id) {
	var cmt=newClass("TComment"),_date = new duoshuo.date(meta_json.meta.created_at);

    cmt.Author = meta_json.meta.author_name;
    if (meta_json.meta.author_key == 1) cmt.AuthorID = 1;
    cmt.EMail = meta_json.meta.author_email;
    cmt.HomePage = meta_json.meta.author_url;
    cmt.IP = meta_json.meta.ip;
    cmt.PostTime = _date.getFullYear() + "-" + (_date.getMonth()) + "-" + _date.getDay() + " " + _date.getHours() + ":" + _date.getMinutes() + ":" + _date.getSeconds();
    cmt.Content = meta_json.meta.message;
	cmt.Agent = meta_json.meta.agent;
    cmt.log_id = meta_json.meta.thread_key;

    //统一判定，防止ShowError
    if (cmt.Author != null) {
        if ((!CheckRegExp(cmt.Author, "[username]")) || (cmt.Author.length > ZC_USERNAME_MAX)) cmt.Author = ZVA_User_Level_Name(5);
    }
    else {
        cmt.Author = ZVA_User_Level_Name(5)
    }

    if (cmt.EMail != null) {
        if (cmt.EMail.length > 0) {
            if ((!CheckRegExp(cmt.EMail, "[email]")) || cmt.EMail.length > ZC_USERNAME_MAX) cmt.EMail = "null@null.com"
        }
    }
    else {
        cmt.EMail = "null@null.com"
    }

    if (cmt.HomePage != null) {
        if (cmt.HomePage.length > 0) {
            if ((!CheckRegExp(cmt.HomePage, "[homepage]")) || cmt.HomePage.length > ZC_HOMEPAGE_MAX) cmt.HomePage = BlogHost
        }
    }
    else {
        cmt.HomePage = BlogHost
    }

    //写入数据库
    if (meta_json.meta.parent_id > 0) {
        var objRs = objConn.Execute("SELECT TOP 1 ds_cmtid FROM blog_Plugin_duoshuo WHERE ds_key='" + meta_json.meta.parent_id + "'");
        if (!objRs.EOF) cmt.ParentID = objRs("ds_cmtid").Value
        //判断是否有父节点
    }
    if (cmt.Post()) {
        objConn.Execute("INSERT INTO [blog_Plugin_duoshuo] (ds_key,ds_cmtid) VALUES('" + meta_json.meta.post_id + "'," + cmt.ID + ")");
    }
	
	cmt=null; 
	return meta_json.log_id;
	
}
duoshuo.api.approve=function(meta_json){
	if(!ZC_MSSQL_ENABLE){
		objConn.Execute("UPDATE blog_Comment INNER JOIN [blog_plugin_duoshuo] ON (((blog_plugin_duoshuo.ds_cmtid)=([blog_Comment].[comm_ID]) And (blog_plugin_duoshuo.ds_key) In("+duoshuo.join2({array:meta_json.meta,before:"'",after:"'",splittag:","})+") )) SET comm_IsCheck=FALSE");
	}
	else{
		objConn.Execute("UPDATE blog_Comment SET comm_IsCheck=0 FROM blog_comment INNER JOIN [blog_plugin_duoshuo] ON (((blog_plugin_duoshuo.ds_cmtid)=([blog_Comment].[comm_ID]) And (blog_plugin_duoshuo.ds_key) In("+duoshuo.join2({array:meta_json.meta,before:"'",after:"'",splittag:","})+") )) ");
	}
	return meta_json.log_id;
}
duoshuo.api.spam=function(meta_json){
	if(!ZC_MSSQL_ENABLE){
		objConn.Execute("UPDATE blog_Comment INNER JOIN [blog_plugin_duoshuo] ON (((blog_plugin_duoshuo.ds_cmtid)=([blog_Comment].[comm_ID]) And (blog_plugin_duoshuo.ds_key) In("+duoshuo.join2({array:meta_json.meta,before:"'",after:"'",splittag:","})+") )) SET comm_IsCheck=TRUE");
	}
	else{
		objConn.Execute("UPDATE blog_Comment SET comm_IsCheck=1 FROM blog_comment INNER JOIN [blog_plugin_duoshuo] ON (((blog_plugin_duoshuo.ds_cmtid)=([blog_Comment].[comm_ID]) And (blog_plugin_duoshuo.ds_key) In("+duoshuo.join2({array:meta_json.meta,before:"'",after:"'",splittag:","})+") )) ");
	}
	return meta_json.log_id;
}
duoshuo.api.deletepost=function(meta_json){
	objConn.Execute("DELETE blog_Comment"+(ZC_MSSQL_ENABLE?"":".*")+" FROM blog_Comment INNER JOIN [blog_plugin_duoshuo] ON  (blog_plugin_duoshuo.ds_cmtid=[blog_Comment].[comm_ID] and blog_plugin_duoshuo.ds_key in("+duoshuo.join2({array:meta_json.meta,before:"'",after:"'",splittag:","})+")) ");
	objConn.Execute("DELETE FROM blog_plugin_duoshuo WHERE ds_key in("+duoshuo.join2({array:meta_json.meta,before:"'",after:"'",splittag:","})+") ");

	return meta_json.log_id;
}
duoshuo.api.update=function(meta_json){return false }//目前还没有逻辑

duoshuo.api.sync=function(){
	var ajax=new ActiveXObject("MSXML2.ServerXMLHTTP"),url="",objRs,data=[],s=0,log_id="";	
	url="http://"+duoshuo.config.Read("duoshuo_api_hostname")+duoshuo.url.log.list+"?limit=50&short_name="+Server.URLEncode(duoshuo.config.Read("short_name"));
	url+="&secret="+Server.URLEncode(duoshuo.config.Read("secret"));
	if(duoshuo.config.Read("log_id")!=undefined){url+="&since_id="+duoshuo.config.Read("log_id");}else{duoshuo.config.Write("log_id",0)}

	ajax.open("GET",url);
	ajax.send();//发送网络请求
	var json=eval("("+ajax.responseText+")");//实例化json
	for(var i=0;i<json.response.length;i++){ 
		switch(json.response[i].action){
			case "create":
				log_id = duoshuo.api.create(json.response[i]) ;
			break;
			case "approve":
				log_id = duoshuo.api.approve(json.response[i]);
			break;
			case "spam":
			case "delete":
				log_id = duoshuo.api.spam(json.response[i]);
			break;
			case "delete-forever":
				log_id = duoshuo.api.deletepost(json.response[i]);
			break;
			case "update":
				log_id = duoshuo.api.deletepost(json.response[i]);
			break;
			default:
			break;
		}
		if(log_id){duoshuo.config.Write("log_id",log_id)}
	}
	duoshuo.config.Save();	
	if(json.response.length==50) duoshuo.api.sync();
}

//我X,Array.prototype居然会冲突，操蛋
duoshuo.join2=function(config){
	if(!config.array) config.array=[]
	if(!config.before) config.before="";
	if(!config.after) config.after="";
	var str="";
	for(var i=0;i<config.array.length;i++){
		str+=config.before+config.array[i]+config.after
		if(i<config.array.length-1){str+=config.splittag}
	}
	return str
}
</script>