<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% Option Explicit %>
<% 'On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<!-- #include file="..\..\c_option.asp" -->
<!-- #include file="..\..\..\zb_system\function\c_function.asp" -->
<!-- #include file="..\..\..\zb_system\function\c_system_lib.asp" -->
<!-- #include file="..\..\..\zb_system\function\c_system_base.asp" -->
<!-- #include file="..\..\..\zb_system\function\c_system_event.asp" -->
<!-- #include file="..\..\..\zb_system\function\c_system_manage.asp" -->
<!-- #include file="..\..\..\zb_system\function\c_system_plugin.asp" -->
<!-- #include file="..\p_config.asp" -->
<%
ShowError_Custom="If Err.Number>0 Then"&vbCrlf&"Response.Write ""({'success':'""&ZVA_ErrorMsg(id)&""'})"""&vbCrlf&"Response.End"&vbCrlf&"End If"
Dim intRight
intRight=1

Sub Duoshuo_NoResponse_Init()
	Call System_Initialize()
	'检查非法链接
	Call CheckReference("")
	'检查权限
	If BlogUser.Level>intRight Then Call ShowError(6)
	If CheckPluginState("duoshuo")=False Then Call ShowError(48)
	Call DuoShuo_Initialize
End Sub



Select Case Request.QueryString("act")
	Case "callback":Call Duoshuo_NoResponse_Init:Call CallBack
	Case "export":Call Duoshuo_NoResponse_Init:Call Export
	Case "fac":Call Duoshuo_NoResponse_Init:Call Fac
	Case "api":Call Api
	Case "api_async":intRight=4:Call api_async
	Case "save":Call Duoshuo_NoResponse_Init:Call Save
End Select


Sub CallBack()
	If Not IsEmpty(duoshuo.get("short_name")) Then
		duoshuo.config.Write "short_name",duoshuo.get("short_name")
		duoshuo.config.Write "secret",duoshuo.get("secret")
		duoshuo.config.Save
	End If
	Call SetBlogHint(True,Empty,Empty)
	Call SetBlogHint_Custom("现在，您必须导出数据到多说，否则可能会出现一些奇怪的问题。")
	Response.Write "<script>top.location.href='export.asp'</script>"
End Sub

Sub Fac()
	duoshuo.config.Delete
	Call SetBlogHint(True,Empty,Empty)
	Response.Redirect "main.asp"
End Sub

Sub Export
	Server.ScriptTimeout=1000000
	Dim intMin,intMax

	
	Response.ContentType="application/json"
	Select Case Request.Form("type")
	Case "all"
		Response.AddHeader "Content-Disposition", "attachment; filename=duoshuo_export_all.json"

		Response.Write "{""threads"":"&ArticleData("").jsString&",""posts"":"
		Response.Write QueryToJson(objConn,"SELECT comm_AuthorID As author_key,comm_ID As post_key,log_id As thread_key,comm_ParentID As parent_key"&_
						",comm_Author As author_name,comm_Email As author_email,comm_HomePage As author_url,comm_PostTime As created_at"&_
						",comm_ip As ip,comm_agent As agent,comm_Content As message FROM blog_Comment WHERE comm_IsCheck=0").jsString
		Response.Write "}"
	Case "article"
		intMin=Request.Form("articlemin")
		intMax=Request.Form("articlemax")
		Call CheckParameter(intMin,"int",1)
		Call CheckParameter(intMax,"int",1)
		Response.AddHeader "Content-Disposition", "attachment; filename=duoshuo_export_article_"&intMin&"to"&intMax&".json"
		Response.Write "{""threads"":"&ArticleData("WHERE log_ID BETWEEN "&intMin&" AND "&intMax).jsString&"}"
	Case "comment"
		intMin=Request.Form("commentmin")
		intMax=Request.Form("commentmax")
		Call CheckParameter(intMin,"int",1)
		Call CheckParameter(intMax,"int",1)
		Response.AddHeader "Content-Disposition", "attachment; filename=duoshuo_export_comment_"&intMin&"to"&intMax&".json"
		Response.Write "{""posts"":"&QueryToJson(objConn,"SELECT comm_AuthorID As author_key,comm_ID As post_key,log_id As thread_key,comm_ParentID As parent_key"&_
						",comm_Author As author_name,comm_Email As author_email,comm_HomePage As author_url,comm_PostTime As created_at"&_
						",comm_ip As ip,comm_agent As agent,comm_Content As message FROM blog_Comment WHERE comm_IsCheck=0 AND(comm_ID BETWEEN "&intMin&_
						" AND "&intMax&")").jsString&"}"
	End Select
	'Dim aryData(),rs,i
	'i=0
	'Set rs=objConn.Execute("SELECT * FROM blog_Comment")
	'Redim aryData(rs.PageSize)
	'Do Until rs.Eof
'		aryData(i)=rs("comm_id")
'		rs.MoveNext
'	Loop
'	Dim s
'	s=(new duoshuo_Duoshuo_aspjson).toJSON(aryData)
'	Response.Write s
End Sub

Function ArticleData(WHERE)
        Dim rs, jsa, col , o , k
        Set rs = objConn.Execute("SELECT [log_ID] As thread_key,[log_CateID],[log_Title] as title,[log_Intro] as excerpt,[log_Level],[log_AuthorID] as author_key,[log_PostTime],[log_ViewNums] as views,[log_Url] as url,[log_Type] FROM [blog_Article] "&WHERE)
        Set jsa = jsArray()
		jsa.Kind=1
        While Not (rs.EOF Or rs.BOF)
				Set o=New TArticle
				If o.LoadInfoByArray(Array(rs(0),"",rs(1),rs(2),rs(3),"",rs(4),rs(5),rs(6),0,rs(7),0,rs(8),False,"","",rs(9),"")) Then
	                Set jsa(Null) = jsObject()
					For Each col In rs.Fields
						If col.Name<>"url" And Left(col.Name,4)<>"log_" Then
	    	            	jsa(Null)(col.Name) = col.Value
						ElseIf col.Name = "create_at" Then
							k=CStr(col.Value)
							jsa(Null)(col.Name) = Year(k) & "-" & Right("0"&Month(k),2) & "-" & Right("0"&Day(k),2) & " " & Right("0"&Hour(k),2) & ":" & Right("0"&Minute(k),2) & ":" & Right("0"&Second(k),2)
						ElseIf col.Name = "excerpt" Then
							jsa(Null)(col.Name) = o.HtmlIntro
						ElseIf col.Name="url" Then
							jsa(Null)(col.Name) = TransferHTML(o.FullUrl,"[zc_blog_host]")
						End If
					Next
        		End If
				Set o=Nothing
		rs.MoveNext
        Wend
        Set ArticleData = jsa
End Function

Function QueryToJSON(dbc, sql)
    Dim rs, jsa, col, k
    Set rs = dbc.Execute(sql)
    Set jsa = jsArray()
	jsa.Kind=1
    While Not (rs.EOF Or rs.BOF)
		Set jsa(Null) = jsObject()
		For Each col In rs.Fields
			If col.Name = "created_at" Then
				k=CStr(col.Value)
				jsa(Null)(col.Name) = Year(k) & "-" & Right("0"&Month(k),2) & "-" & Right("0"&Day(k),2) & " " & Right("0"&Hour(k),2) & ":" & Right("0"&Minute(k),2) & ":" & Right("0"&Second(k),2)
			Else
	            jsa(Null)(col.Name) = col.Value
            End If
		Next
        rs.MoveNext
    Wend
    Set QueryToJSON = jsa
End Function


Function jsObject
	Set jsObject = new duoshuo_aspjson
	jsObject.Kind = 0
End Function

Function jsArray
	Set jsArray = new duoshuo_aspjson
	jsArray.Kind = 1
End Function




Sub Api()
	
End Sub



Sub Save()
	Dim obj
	For Each obj In Request.Form
		duoshuo.config.Write obj,Request.Form(obj)
	Next
	duoshuo.config.Save()
	Call SetBlogHint(True,Empty,Empty)
	Response.Redirect "main.asp?act=setting"
End Sub


Function comparison(ByVal str1,ByVal str2)
	str1=CDBl(str1):str2=CDBl(str2)
	If str1>str2 Then 
		comparison=CStr(str1)
	Else
		comparison=CStr(str2)
	End If
End Function
%>
<script language="javascript" runat="server" >

function Api_Async(){
	Response.ContentType="application/javascript";//配置mime头
	
	var _last=Application(ZC_BLOG_CLSID+"duoshuo_lastpub"),_now=new Date().getTime();
	if(typeof(_last)=="number"){//20分钟时间限制
		if((_now-_last)/1000>=60*0.1){ 
			_last=_now
		}
		else{
			Response.Write("({'last':'"+_last+"','now':'"+_now+"','status':'waiting'})");
			Response.End()
		}
	}
	else{
		_last=_now
	}
	Application(ZC_BLOG_CLSID+"duoshuo_lastpub")=_now
	//x分钟内不再请求
	Response.Write("({'status':'"+Api_Run().success.replace(/\r/g,"\\r").replace(/\n/g,"\\n").replace(/'/g,"\'")+"'})")
	
}
function Api_Run(){
	Duoshuo_NoResponse_Init();//加载数据库
	if(duoshuo.config.Read("duoshuo_cron_sync_enabled")!="async") return {'success':'noasync'}
	//try{
		var ajax=new ActiveXObject("MSXML2.ServerXMLHTTP"),url="",objRs,data=[],s=0,log_id="";
		
		url="http://"+duoshuo.config.Read("duoshuo_api_hostname")+"/log/list.json?short_name="+Server.URLEncode(duoshuo.config.Read("short_name"));
		url+="&secret="+Server.URLEncode(duoshuo.config.Read("secret"));
		if(duoshuo.config.Read("log_id")!=undefined){url+="&since_id="+duoshuo.config.Read("log_id");}else{duoshuo.config.Write("log_id",0)}

		ajax.open("GET",url);
		ajax.send();//发送网络请求
Response.Write(ajax.responseText)
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
					log_id = duoshuo.api.spam(json.response[i]);
				break;
				case "delete":
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
		BlogReBuild_Statistics();
		BlogReBuild_Comments();
		BlogReBuild_Functions();
		BlogReBuild_Default();
		return {'success':'success'}
	//}
	//catch(e){
	//	return {'success':e.message}
	//}
}
</script>