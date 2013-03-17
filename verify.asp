<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Option Explicit%>
<!-- #include file="../../c_option.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_function.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_lib.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_base.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_event.asp" -->
<!-- #include file="../../../ZB_SYSTEM/function/c_system_plugin.asp" -->
<!-- #include file="../../plugin/p_config.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="<%=ZC_BLOG_LANGUAGE%>" lang="<%=ZC_BLOG_LANGUAGE%>">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta http-equiv="Content-Language" content="<%=ZC_BLOG_LANGUAGE%>" />
<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7" />
<link rel="stylesheet" rev="stylesheet" href="../../../zb_system/css/admin.css" type="text/css" media="screen" />
<script language="JavaScript" src="../../../zb_system/script/common.js" type="text/javascript"></script>
<script language="JavaScript" src="../../../zb_system/script/md5.js" type="text/javascript"></script>
<title><%=ZC_BLOG_TITLE & ZC_MSG044 & ZC_MSG009%></title>
</head>
<body>
<%
If Not CheckPluginState("duoshuo") Then Response.End

Call System_Initialize
	
Select Case Request.QueryString("act")
Case "login"
%>
<div class="bg"></div>
<div id="wrapper">
  <div class="logo"><img src="../../../zb_system/image/admin/none.gif" title="Z-Blog<%=ZC_MSG009%>" alt="Z-Blog<%=ZC_MSG009%>"/></div>
  <div class="login"> <a href="?act=bind&duoshuo_userid=<%=Server.URLEncode(Request.QueryString("duoshuo_userid"))%>&accesstoken=<%=Server.URLEncode(Request.QueryString("accesstoken"))%>&dName=<%=Server.URLEncode(Request.QueryString("dName"))%>" title="绑定现有帐号">绑定现有帐号</a>
    <%If CheckPluginState("RegPage") Then%>
    <a href="../RegPage/Reg.asp?duoshuo_userid=<%=Server.URLEncode(Request.QueryString("duoshuo_userid"))%>&accesstoken=<%=Server.URLEncode(Request.QueryString("accesstoken"))%>&dName=<%=Server.URLEncode(Request.QueryString("dName"))%>" title="新建账户">新建账户</a>
    <%Else%>
    <a href="<%=BlogHost%>" title="返回首页"><img src="resources/return.png"/></a>
    <%End If%>
  </div>
</div>
<%
Case "bind"
	If BlogUser.Level<5 Then Duoshuo_SaveReg() :Response.Redirect "main.asp"
%>
<div class="bg"></div>
<div id="wrapper">
  <div class="logo"><img src="../../../zb_system/image/admin/none.gif" title="Z-Blog<%=ZC_MSG009%>" alt="Z-Blog<%=ZC_MSG009%>"/></div>
  <div class="login">
    <form id="frmLogin" method="post" action="">
      <dl>
        <dd>
          <label for="edtUserName"><%=ZC_MSG003%>:</label>
          <input type="text" id="edtUserName" name="edtUserName" size="20" tabindex="1" value="<%=Replace(TransferHTML(Request.QueryString("dName"),"[nohtml]"),"""","'")%>"/>
        </dd>
        <dd>
          <label for="edtPassWord"><%=ZC_MSG002%>:</label>
          <input type="password" id="edtPassWord" name="edtPassWord" size="20" tabindex="2" />
        </dd>
        <input type="hidden" name="duoshuo_userid" value="<%=Replace(TransferHTML(Request.QueryString("duoshuo_userid"),"[nohtml]"),"""","'")%>"/>
        <input type="hidden" name="AccessToken" value="<%=Replace(TransferHTML(Request.QueryString("AccessToken"),"[nohtml]"),"""","'")%>"/>
      </dl>
      <dl>
        <dd class="submit">
          <input id="btnPost" name="btnPost" type="submit" value="<%=ZC_MSG260%>" class="button" tabindex="4"/>
        </dd>
      </dl>
      <input type="hidden" name="username" id="username" value="" />
      <input type="hidden" name="password" id="password" value="" />
      <input type="hidden" name="savedate" id="savedate" value="30" />
    </form>
  </div>
</div>
<script language="JavaScript" type="text/javascript">
        
        if(GetCookie("username")){document.getElementById("edtUserName").value=unescape(GetCookie("username"))};
        $("#btnPost").click(function(){
            var strUserName=document.getElementById("edtUserName").value;
            var strPassWord=document.getElementById("edtPassWord").value;
            var strSaveDate=document.getElementById("savedate").value
            if((strUserName=="")||(strPassWord=="")){
                alert("<%=ZC_MSG010%>");
                return false;
            }
            strUserName=escape(strUserName);
            strPassWord=MD5(strPassWord);
            SetCookie("username",strUserName,strSaveDate);
            SetCookie("password",strPassWord,strSaveDate);
            document.getElementById("frmLogin").action="verify.asp?act=verify"
            document.getElementById("username").value=unescape(strUserName);
            document.getElementById("password").value=strPassWord;
            document.getElementById("savedate").value=strSaveDate;
            document.getElementById("duoshuo_userid").value="<%=TransferHTML(Request.QueryString("duoshuo_userid"),"[nohtml][<][>][""][']")%>";
			document.getElementById("accesstoken").value="<%=TransferHTML(Request.QueryString("accesstoken"),"[nohtml][<][>][""][']")%>";
        })
        
        
        </script>
</body>
</html>
<%
Case "verify"
	If Login=True Then
		Call Duoshuo_SaveReg
		Response.Write "<script>alert('绑定成功！');location.href="""&BlogHost&"zb_system/admin/admin.asp""</script>"
		Response.End
	Else 
		Response.Write "<script>alert('密码错误！');history.go(-1)</script>"
	End If
End Select
	
%>
</body>
</html>
<%



%>