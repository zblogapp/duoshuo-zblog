<%
Class duoshuo_oauth

	Public ID
	Public Duoshuo_UserID
	Public ZB_UserID
	Public AccessToken
	Public objXmlHttp
	
	'用户绑定
	Public Function Bind()
	End Function
	
	'用户回调
	Public Function CallBack(ds_code)
		
		'objXmlHttp.Open "POST","http://" & duoshuo.config.Read("duoshuo_api_hostname") & duoshuo.url.api.callback 
		'objXmlHttp.setRequestHeader "Content-Type","application/x-www-form-urlencoded"
		'objXmlHttp.Send "code="&ds_code
		Dim strData
		
		'strData=objXmlHttp.ResponseText
		'未注册
		strData="{""remind_in"":7776000,""access_token"":""nicai"",""expires_in"":7776000,""user_id"":""1519109"",""code"":0}"
		Dim oObj
		Set oObj=duoshuo.parseJSON(strData)
		
		If LoadInfoByDsId(oObj.user_id) Then
			AccessToken=oObj.access_token
			Duoshuo_UserID=oObj.user_id
			Call Post()
			If Not Duoshuo_Login Then
				Response.Redirect "verify.asp?act=login&duoshuo_userid="&oObj.user_id&"&accesstoken="&oObj.access_token
				'登录失败处理，虽然不太可能
			End If
		Else
			Response.Redirect "verify.asp?act=login&duoshuo_userid="&oObj.user_id&"&accesstoken="&oObj.access_token
			'未注册处理
		End If
		
		
		
	End Function
	
	'数据写入
	Public Function Post()
		Call CheckParameter(ID,"int",0)
		Call CheckParameter(ZB_UserID,"int",0)
		If ID=0 Then
			objConn.Execute "INSERT INTO blog_Plugin_Duoshuo_Member (ds_key,ds_memid,ds_accesstoken) VALUES('"&FilterSQL(Duoshuo_UserID)&"',"&ZB_USERID&",'"&AccessToken&"')"
			Call LoadInfoByDsId(Duoshuo_UserID)
		Else
			objConn.Execute "UPDATE blog_Plugin_Duoshuo_Member SET ds_key='"&FilterSQL(Duoshuo_UserID)&"',ds_memid="&ZB_USERID&",ds_accesstoken='"&AccessToken&"'"
		End If
		Post=True
	End Function
	
	'API调用
	Public Function API(Method,URL,Param)
	End Function
	
	'根据多说ID读取相关信息
	Public Function LoadInfoByDsId(dsID)
		Dim objRs
		Set objRs=objConn.Execute("SELECT * FROM blog_Plugin_duoshuo_Member WHERE ds_key='"&FilterSQL(dsID)&"'")
		If Not objRs.Eof Then
			ID=objRs("ds_id")
			Duoshuo_UserID=objRs("ds_key")
			ZB_UserID=objRs("ds_memid")
			AccessToken=objRs("ds_accesstoken")
			LoadInfoByDsId=True
		End If
	End Function
	
	'根据ZBLOG用户ID读取相关信息
	Public Function LoadInfoByZBId(zbID)
		Call CheckParameter(zbID,"int",0)
		Dim objRs
		Set objRs=objConn.Execute("SELECT * FROM blog_Plugin_duoshuo_Member WHERE ds_memid="&zbID)
		If Not objRs.Eof Then
			ID=objRs("ds_id")
			Duoshuo_UserID=objRs("ds_key")
			ZB_UserID=objRs("ds_memid")
			AccessToken=objRs("ds_accesstoken")
			LoadInfoByZBId=True
		End If

	End Function
	
	Function Duoshuo_Login()
		Dim oUser
		Set oUser=New TUser 
		If oUser.LoadInfoById(ZB_UserID) Then
			Response.Write "a"
			BlogUser.LoginType="Self"
			BlogUser.Name=oUser.name
			BlogUser.PassWord=oUser.Password
			If BlogUser.Verify=True Then
				Response.Cookies("password")=BlogUser.PassWord
				If Request.Form("savedate")<>0 Then
					Response.Cookies("password").Expires = DateAdd("d", 1, now)
				End If
				Response.Cookies("password").Path = CookiesPath()
			Else
				Duoshuo_Login=False
			End If
			Response.Cookies("username")=escape(BlogUser.name)
			If Request.Form("savedate")<>0 Then
				Response.Cookies("username").Expires = DateAdd("d", 1, now)
			End If
			Response.Cookies("username").Path = CookiesPath()
			Response.Redirect BlogHost & "zb_system/cmd.asp?act=login"
			Duoshuo_Login=True
		Else
			Duoshuo_Login=False
		End If
	End Function

	
	Sub Class_Initialize()
		Set objXmlHttp=Server.CreateObject("MSXML2.ServerXMLHTTP")
	End Sub
	
	Sub Class_Terminate()
		Set objXmlHttp=Nothing
	End Sub

End Class
%>
<script language="javascript" runat="server">
duoshuo.url.api={
	callback:"/oauth2/access_token"
}
</script>