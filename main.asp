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
Call System_Initialize()
'检查非法链接
Call CheckReference("")
'检查权限
If BlogUser.Level>1 Then Call ShowError(6)
If CheckPluginState("duoshuo")=False Then Call ShowError(48)
BlogTitle="多说社会化评论"
Call DuoShuo_Initialize
%>
<!--#include file="..\..\..\zb_system\admin\admin_header.asp"-->
<style type="text/css">
tr {
	height: 32px
}
#divMain2 ul li {
	margin-top: 6px;
	margin-bottom: 6px
}
.bold {
	font-weight: bold;
}
.note {
	margin-left: 10px
}
</style>
<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->
        <div id="divMain">
          <div id="ShowBlogHint">
            <%Call GetBlogHint()%>
          </div>
          <div class="divHeader"><%=BlogTitle%><%=NewWindow()%></div>
          <div class="SubMenu"><%=duoshuo_SubMenu(duoshuo.get("act"))%></div>
          <div id="divMain2"> 
            <script type="text/javascript">ActiveTopMenu("aPlugInMng");</script>
            <%
			If duoshuo.config.Read("short_name")="" Then
			%>
            <iframe id="duoshuo-remote-window" src="http://duoshuo.com/connect-site/?name=<%=Server.URLEncode(ZC_BLOG_TITLE)%>&description=<%=Server.URLEncode(ZC_BLOG_SUBTITLE)%>&url=<%=Server.URLEncode(ZC_BLOG_HOST)%>&siteurl=<%=Server.URLEncode(ZC_BLOG_HOST)%>&system_version=<%=BlogVersion%>&plugin_version=<%=Server.URLEncode(duoshuo.config.Read("ver"))%>&system=zblog&callback=<%=Server.URLEncode(BlogHost &"zb_users/plugin/duoshuo/noresponse.asp?act=callback")%>&user_key=<%=BlogUser.ID%>&user_name=<%=Server.URLEncode(BlogUser.Name)%>&admin_email=<%=Server.URLEncode(BlogUser.EMail)%>&local_api_url=<%=Server.URLEncode(BlogHost & "zb_users/plugin/duoshuo/noresponse.asp?act=api")%>" style="border:0; width:100%; height:580px;"></iframe>
            <%
			Else
			Select Case duoshuo.get("act")
			Case "personal"
			%>
            <iframe id="duoshuo-remote-window" src="http://<%=duoshuo.config.Read("short_name")%>.duoshuo.com/admin/settings/?jwt=<%=duoshuo_getjwt()%>" style="border:0; width:100%; height:580px;"></iframe>
            <%
			Case "statistics"
			%>
            <iframe id="duoshuo-remote-window" src="http://<%=duoshuo.config.Read("short_name")%>.duoshuo.com/admin/statistics/?jwt=<%=duoshuo_getjwt()%>" style="border:0; width:100%; height:580px;"></iframe>

            <%
			Case "setting"
			%>
            <form action="noresponse.asp?act=save" method="post">
              <table width="100%">
                <thead>
                  <tr>
                    <th width="30%">配置项 </th>
                    <th>选择 </th>
                  </tr>
                </thead>
                <tbody>
                  <tr>
                    <td><p><span class="bold"> · 多说API服务器</span><br/>
                        <span class="note">选择一个速度更快的服务器</span></p></td>
                    <td><ul>
                        <li>
                          <label>
                            <input type="radio" name="duoshuo_api_hostname" value="api.duoshuo.com"<%=GetChecked("duoshuo_api_hostname","api.duoshuo.com")%>>
                            api.duoshuo.com(国内主机使用)</label>
                        </li>
                        <li>
                          <label>
                            <input type="radio" name="duoshuo_api_hostname" value="api.duoshuo.org"<%=GetChecked("duoshuo_api_hostname","api.duoshuo.org")%>>
                            api.duoshuo.org(国外主机使用)</label>
                        </li>
                        <li>
                          <label>
                            <input type="radio" name="duoshuo_api_hostname" value="118.144.80.201"<%=GetChecked("duoshuo_api_hostname","118.144.80.201")%>>
                            118.144.80.201(DNS故障主机使用)</label>
                        </li>
                      </ul></td>
                  </tr>
                  <tr>
                    <td><p><span class="bold"> · 本地数据备份</span><br/>
                        <span class="note">评论同时写入本地数据库</span></p></td>
                    <td><ul>
                        <li>
                          <label>
                            <input type="radio" name="duoshuo_cron_sync_enabled" value="async"<%=GetChecked("duoshuo_cron_sync_enabled","async")%>>
                            定时写入</label>
                        </li>
                        <!--<li>
                          <label>
                            <input type="radio" name="duoshuo_cron_sync_enabled" value="sync"<%=GetChecked("duoshuo_cron_sync_enabled","sync")%>>
                            实时写入</label>
                        </li>-->
                        <li>
                          <label>
                            <input type="radio" name="duoshuo_cron_sync_enabled" value="off"<%=GetChecked("duoshuo_cron_sync_enabled","off")%>>
                            不写入</label>
                        </li>
                      </ul>
                  </tr>
                  <tr>
                    <td><p><span class="bold"> · 评论数修正</span><br/>
                        <span class="note">AJAX加载文章的评论数</span></p></td>
                    <td><input type="text" class="checkbox" name="duoshuo_cc_fix" value="<%=duoshuo.config.Read("duoshuo_cc_fix")%>" checked="checked"></td>
                  </tr>
                  <tr>
                    <td><p><span class="bold"> · 评论框前缀</span><br/>
                        <span class="note">仅在主题和评论框的div嵌套不正确的情况下使用 </span></p></td>
                    <td><input type="text" name="duoshuo_comments_wrapper_intro" value="<%=duoshuo.config.Read("duoshuo_comments_wrapper_intro")%>" style="width:50%"/></td>
                  </tr>
                  <tr>
                    <td><p><span class="bold"> · 评论框后缀</span><br/>
                        <span class="note">仅在主题和评论框的div嵌套不正确的情况下使用 </span></p></td>
                    <td><input type="text" name="duoshuo_comments_wrapper_outro" value="<%=duoshuo.config.Read("duoshuo_comments_wrapper_outro")%>" style="width:50%"/></td>
                  </tr>
                  <tr>
                    <td><p><span class="bold"> · SEO优化</span><br/>
                        <span class="note">搜索引擎爬虫访问网页时，显示静态HTML评论</span></p></td>
                    <td><input type="text" class="checkbox" name="duoshuo_seo_enabled" value="<%=duoshuo.config.Read("duoshuo_seo_enabled")%>"/></td>
                  </tr>
                  <tr>
                    <td><p><span class="bold"> · 启用多说登录</span><br/><span class="note">如果想停用，请打开侧栏管理，编辑控制面板，删除&ltdiv class="ds-login"&gt;&lt;/div&gt;即可<br/></p></td>
                    <td><p> </p>
                      <p>
                        <input name="" type="button" class="button" onClick="location.href='noresponse.asp?act=specfg&t=login'" value="立即往侧栏写入多说登录" />
                      </p></td>
                  </tr>
                  <tr>
                    <td><p><span class="bold"> · 其它</span></p></td>
                    <td><p> </p>
                      <p>
                        <input name="" type="button" class="button" onClick="if(confirm('你确定要继续吗？')){location.href='noresponse.asp?act=fac'}" value="清空插件配置" />
                      </p></td>
                  </tr>
                </tbody>
                <tfoot>
                </tfoot>
              </table>
              <p>
                <input type="submit" class="button" value="提交" />
              </p>
            </form>
            <%
			Case Else
			%>
            <iframe id="duoshuo-remote-window" src="http://<%=duoshuo.config.Read("short_name")%>.duoshuo.com/admin/?jwt=<%=duoshuo_getjwt()%>" style="width:100%; border:0;"></iframe>
            <%
			End Select
			End If
			%>
          </div>
        </div>
        <!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->
<script type="text/javascript">
ActiveLeftMenu("aCommentMng");
$(document).ready(function(){
	var iframe = $('#duoshuo-remote-window');
	resetIframeHeight = function(){
		iframe.height($(window).height() - iframe.offset().top);
	};
	resetIframeHeight();
	$(window).resize(resetIframeHeight);
});
$('#duoshuo_manage').addClass('sidebarsubmenu1');
</script>
<%Call System_Terminate()%>
<%
Function GetChecked(name,value)
	If duoshuo.config.Read(name)=value Then GetChecked=" checked=""checked"" "
End Function

Function NewWindow()
	NewWindow="<script type='text/javascript'>$(function(){var w=$('#duoshuo-remote-window');if(w.length>0){$('.divHeader').append('<span style=""font-size:12px;padding: 3px 8px;background: #f1f1f1;margin-left: 4px;color: #21759b;-webkit-border-radius: 3px;border-radius: 3px;border-width: 1px;border-style: solid;""><a href='+w.attr(""src"")+' target=""_blank"">新窗口打开</a></span>')}})</script>"
End Function
'window.open(k.attr('src'))
%>