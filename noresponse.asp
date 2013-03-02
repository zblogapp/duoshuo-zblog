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

Sub Duoshuo_NoResponse_Init()
	Call System_Initialize()
	'检查非法链接
	Call CheckReference("")
	'检查权限
	If BlogUser.Level>1 Then Call ShowError(6)
	If CheckPluginState("duoshuo")=False Then Call ShowError(48)
	Call DuoShuo_Initialize
End Sub



Select Case Request.QueryString("act")
	Case "callback":Call Duoshuo_NoResponse_Init:Call CallBack
	Case "export":Call Duoshuo_NoResponse_Init:Call Export
	Case "fac":Call Duoshuo_NoResponse_Init:Call Fac
	Case "api":Call Api
	Case "api_async":Call api_async
	Case "save":Call Duoshuo_NoResponse_Init:Call Save
End Select


Sub CallBack()
	If Not IsEmpty(duoshuo.get("short_name")) Then
		duoshuo.config.Write "short_name",duoshuo.get("short_name")
		duoshuo.config.Write "secret",duoshuo.get("secret")
		duoshuo.config.Save
	End If
	Call SetBlogHint(True,Empty,Empty)
	Response.Write "<script>top.location.reload()</script>"
End Sub

Sub Fac()
	duoshuo.config.Delete
	Call SetBlogHint(True,Empty,Empty)
	Response.Redirect "main.asp"
End Sub

Sub Export
	
	Response.ContentType="application/json"
	Response.AddHeader "Content-Disposition", "attachment; filename=duoshuo_export.json"

	Response.Write "{""threads"":"&ArticleData().jsString&",""posts"":"
	Response.Write QueryToJson(objConn,"SELECT comm_AuthorID As author_key,comm_ID As post_key,log_id As thread_key,comm_ParentID As parent_key"&_
					",comm_Author As author_name,comm_Email As author_email,comm_HomePage As author_url,comm_PostTime As created_at"&_
					",comm_ip As ip,comm_agent As agent,comm_Content As message FROM blog_Comment WHERE comm_IsCheck=0").jsString
	Response.Write "}"
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

Function ArticleData()
        Dim rs, jsa, col , o
        Set rs = objConn.Execute("SELECT [log_ID] As thread_key,[log_CateID],[log_Title] as title,[log_Intro] as excerpt,[log_Level],[log_AuthorID] as author_key,[log_PostTime],[log_ViewNums] as views,[log_Url] as url,[log_Type] FROM [blog_Article]")
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
							jsa(Null)(col.Name) = Year(col.Value) & "-" & Month(col.Value) & "-" & Day(col.Value) & " " & Hour(col.Value) & ":" & Minute(col.Value) & ":" & Second(col.Value)
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
<script language="javascript" runat="server">

function Api_Async(){
	Response.ContentType="application/javascript";//配置mime头
	
	var _last=Application(ZC_BLOG_CLSID+"duoshuo_lastpub"),_now=new Date().getTime();
	if(typeof(_last)=="number"){//20分钟时间限制
		if((_now-_last)/1000>=60*20){ 
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
	try{
		var ajax=new ActiveXObject("MSXML2.ServerXMLHTTP"),url="",objRs,data=[],s=0;
		var _date=new Date();
		
		url="http://"+duoshuo.config.Read("duoshuo_api_hostname")+"/log/list.json?short_name="+Server.URLEncode(duoshuo.config.Read("short_name"));
		url+="&secret="+Server.URLEncode(duoshuo.config.Read("secret"));
		if(duoshuo.config.Read("log_id")!=undefined){url+="&since_id="+duoshuo.config.Read("log_id");}else{duoshuo.config.Write("log_id",0)}
		//如果不存在logid就设为0
		objRs=null;

		ajax.open("GET",url);
		ajax.send();//发送网络请求
	
		var json=eval("("+ajax.responseText+")");//实例化json
		for(var i=0;i<json.response.length;i++){
			var cmt=newClass("TComment"),tmp=json.response[i]; //实例化评论对象
			if(tmp.action=="create"){
				_date={
					"date":tmp.meta.created_at,
					"getMonth":function(){return this.date.split("T")[0].split("-")[1]},
					"getDay":function(){return this.date.split("T")[0].split("-")[2]},
					"getFullYear":function(){return this.date.split("T")[0].split("-")[0]},
					"getHours":function(){return this.date.split("T")[1].split(":")[0]},
					"getMinutes":function(){return this.date.split("T")[1].split(":")[1]},
					"getSeconds":function(){return this.date.split("T")[1].split(":")[2].split("+")[0]}
				};
				//Microsoft JScript for ASP不支持new Date("xxxTxxx")
				cmt.Author=tmp.meta.author_name;
				if(tmp.meta.author_key==1) cmt.AuthorID=1;
				cmt.EMail=tmp.meta.author_email;
				cmt.HomePage=tmp.meta.author_url;
				cmt.IP=tmp.meta.ip;
				cmt.PostTime=_date.getFullYear()+"-"+(_date.getMonth())+"-"+_date.getDay()+" "+_date.getHours()+":"+_date.getMinutes()+":"+_date.getSeconds();
				cmt.Content=tmp.meta.message;
				cmt.log_id=tmp.meta.thread_key;
				if(tmp.meta.parent_id>0){
					var objRs=objConn.Execute("SELECT TOP 1 ds_cmtid FROM blog_Plugin_duoshuo WHERE ds_key='"+tmp.meta.parent_id+"'");
					if(!objRs.EOF) cmt.ParentID=objRs("ds_cmtid").Value
					//判断是否有父节点
				} 
				if(cmt.Post()){
					objConn.Execute("INSERT INTO [blog_Plugin_duoshuo] (ds_key,ds_cmtid) VALUES('"+tmp.meta.post_id+"',"+cmt.ID+")");
					duoshuo.config.Write("log_id",tmp.log_id)
				}
				
			}
			cmt=null; 
			
		}
		duoshuo.config.Save();
		return {'success':'success'}
	}
	catch(e){
		return {'success':e.message}
	}
}
</script>