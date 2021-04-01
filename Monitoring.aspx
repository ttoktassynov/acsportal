<%@ Page Title="Учет НКТ" Language="C#" MasterPageFile="~/Site.master" AutoEventWireup="true" CodeFile="Monitoring.aspx.cs" Inherits="Monitoring" %>
<%@ Register TagPrefix="asp" Namespace="AjaxControlToolkit" Assembly="AjaxControlToolkit"%>
<asp:Content ID="Content1" ContentPlaceHolderID="HeadContent" Runat="Server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="MainContent" Runat="Server">
    <asp:ToolkitScriptManager ID="ToolkitScriptManagerMonitoring" runat="server"></asp:ToolkitScriptManager>
   
  
    <asp:TabContainer ID="tc_monitoring" runat="server" ActiveTabIndex="4">
        
        <asp:TabPanel runat="server" HeaderText="Учет НКТ" ID="tp_nkt" TabIndex="5" Visible = "true" >
            <ContentTemplate>
            <asp:LinkButton ID="lb_narabotka_skv" runat="server" OnClick="lb_nktreports_Click" >Наработка на отказ НКТ по скважинам</asp:LinkButton><br />
            <asp:LinkButton ID="lb_portyanka_skv" runat="server" OnClick="lb_nktreports_Click" >Сводка по проведенным ПРС по скважинам</asp:LinkButton><br />
            <asp:LinkButton ID="lb_portyanka_bri" runat="server" OnClick="lb_nktreports_Click" >Сводка по проведенным ПРС по бригадам</asp:LinkButton><br />
            <asp:LinkButton ID="lb_zameri_skv" runat="server" OnClick="lb_nktreports_Click">Сводка по замерам по скважинам</asp:LinkButton><br />
            <p runat="server" id = "par_reportname_nkt" style="font-weight:bold"></p>
            <p runat="server" id = "par_error_nkt" style="color:green;font-size:large"></p>
                 <table runat="server" id = "tbl_portyanka_bri" visible="False">
                    <tr id="Tr28" runat="server">
                        <td id="Td46" runat="server">НГДУ</td>
                        <td id="Td48" runat="server">
                            <asp:DropDownList ID="ddl_ngdu" runat="server" Width = "150">
                                    <asp:ListItem Text = "НГДУ-1" Value = "1"></asp:ListItem>
                                    <asp:ListItem Text = "НГДУ-2" Value = "2"></asp:ListItem>
                                    <asp:ListItem Text = "НГДУ-3" Value = "3"></asp:ListItem>
                                    <asp:ListItem Text = "НГДУ-4" Value = "4"></asp:ListItem>
                            </asp:DropDownList>
                        </td>
                    </tr>
                    <tr id="Tr29" runat="server">
                        <td id="Td49" runat="server">Месяц</td>
                        <td id="Td50" runat="server">
                            <asp:DropDownList ID="ddl_nktreportmonth" runat="server" Width = "150"></asp:DropDownList>
                        </td>

                    </tr>
                    <tr id="Tr30" runat="server">
                        <td id="Td51" runat="server">Год</td>
                        <td id="Td52" runat="server">
                            <asp:DropDownList ID="ddl_nktreportyear" runat="server" Width="150">
                                    <asp:ListItem Text = "2011" Value = "2011"></asp:ListItem>
                                    <asp:ListItem Text = "2012" Value = "2012"></asp:ListItem>
                                    <asp:ListItem Text = "2013" Value = "2013"></asp:ListItem>
                                    <asp:ListItem Text = "2014" Value = "2014"></asp:ListItem>
                            </asp:DropDownList>
                        </td>
                    </tr>
                    <tr id="Tr31" runat="server">
                        <td id="Td53" runat="server">
                            <asp:Button ID="btn_getportyankareport" runat="server" 
                                Text="Загрузить" onclick="btn_getportyankareport_Click" />
                        </td>
                        
                    </tr>
                </table>  
                <table runat="server" id = "tbl_portyanka_skv" visible="False">
                    <tr id="Tr37" runat="server">
                        <td id="Td63" runat="server">НГДУ</td>
                        <td id="Td64" runat="server">
                            <asp:DropDownList ID="ddl_ngdu_portyanka_skv" runat="server" Width = "150" OnSelectedIndexChanged="ddl_ngdu_portyanka_skv_OnSelectedIndexChanged" AutoPostBack="True">
                                 <asp:ListItem Text = "НГДУ-1" Value = "1"></asp:ListItem>
                                    <asp:ListItem Text = "НГДУ-2" Value = "2"></asp:ListItem>
                                    <asp:ListItem Text = "НГДУ-3" Value = "3"></asp:ListItem>
                                    <asp:ListItem Text = "НГДУ-4" Value = "4"></asp:ListItem>
                            </asp:DropDownList>
                        </td>
                    </tr>
                    <tr id="Tr41" runat="server">
                        <td id="Td70" runat="server">ЦДНГ</td>
                        <td id="Td71" runat="server">
                            <asp:DropDownList ID="ddl_cdng_portyanka_skv" runat="server" Width = "150"></asp:DropDownList>
                        </td>

                    </tr>
                    <tr id="Tr42" runat="server">
                        <td id="Td72" runat="server">Фонд</td>
                        <td id="Td73" runat="server">
                            <asp:DropDownList ID="ddl_fond_portyanka_skv" runat="server" Width="150">
                                    <asp:ListItem Text = "Добывающий" Value = "Д.ф."></asp:ListItem>
                                    <asp:ListItem Text = "Нагнетательный" Value = "Н.ф."></asp:ListItem>
                            </asp:DropDownList>
                        </td>
                    </tr>
                    <tr id="Tr38" runat="server">
                        <td id="Td65" runat="server">Месяц</td>
                        <td id="Td66" runat="server">
                            <asp:DropDownList ID="ddl_month_portyanka_skv" runat="server" Width = "150"></asp:DropDownList>
                        </td>

                    </tr>
                    <tr id="Tr39" runat="server">
                        <td id="Td67" runat="server">Год</td>
                        <td id="Td68" runat="server">
                            <asp:DropDownList ID="ddl_year_portyanka_skv" runat="server" Width="150">
                                    <asp:ListItem Text = "2011" Value = "2011"></asp:ListItem>
                                    <asp:ListItem Text = "2012" Value = "2012"></asp:ListItem>
                                    <asp:ListItem Text = "2013" Value = "2013"></asp:ListItem>
                                    <asp:ListItem Text = "2014" Value = "2014"></asp:ListItem>
                            </asp:DropDownList>
                        </td>
                    </tr>
                    <tr id="Tr40" runat="server">
                        <td id="Td69" runat="server">
                            <asp:Button ID="btn_getportyankaskvreport" runat="server" 
                                Text="Загрузить" onclick="btn_getportyankaskvreport_Click" />
                        </td>
                        
                    </tr>
                </table>  
                 <table runat="server" id = "tbl_narabotka_skv" visible="False">
                    <tr id="Tr32" runat="server">
                        <td id="Td54" runat="server">НГДУ</td>
                        <td id="Td55" runat="server">
                            <asp:DropDownList ID="ddl_ngdu_nar" runat="server" Width = "150" AutoPostBack="true" OnSelectedIndexChanged="ddl_ngdu_nas_SelectedIndexChanged">
                                    <asp:ListItem Text = "НГДУ-1" Value = "1"></asp:ListItem>
                                    <asp:ListItem Text = "НГДУ-2" Value = "2"></asp:ListItem>
                                    <asp:ListItem Text = "НГДУ-3" Value = "3"></asp:ListItem>
                                    <asp:ListItem Text = "НГДУ-4" Value = "4"></asp:ListItem>
                            </asp:DropDownList>
                        </td>
                    </tr>
                    <tr id="Tr33" runat="server">
                        <td id="Td56" runat="server">ЦДНГ</td>
                        <td id="Td57" runat="server">
                            <asp:DropDownList ID="ddl_cdng_nar" runat="server" Width = "150"></asp:DropDownList>
                        </td>

                    </tr>
                    <tr id="Tr36" runat="server">
                        <td id="Td61" runat="server">Фонд</td>
                        <td id="Td62" runat="server">
                            <asp:DropDownList ID="ddl_fond_nar" runat="server" Width="150">
                                    <asp:ListItem Text = "Добывающий" Value = "Д.ф."></asp:ListItem>
                                    <asp:ListItem Text = "Нагнетательный" Value = "Н.ф."></asp:ListItem>
                            </asp:DropDownList>
                        </td>
                    </tr>
                    <tr id="Tr34" runat="server">
                        <td id="Td58" runat="server">Год</td>
                        <td id="Td59" runat="server">
                            <asp:DropDownList ID="ddl_year_nar" runat="server" Width="150">
                                    <asp:ListItem Text = "2011" Value = "2011"></asp:ListItem>
                                    <asp:ListItem Text = "2012" Value = "2012"></asp:ListItem>
                                    <asp:ListItem Text = "2013" Value = "2013"></asp:ListItem>
                                    <asp:ListItem Text = "2014" Value = "2014"></asp:ListItem>
                            </asp:DropDownList>
                        </td>
                    </tr>

                    <tr id="Tr35" runat="server">
                        <td id="Td60" runat="server">
                            <asp:Button ID="btn_get_narabotka_skv" runat="server" 
                                Text="Загрузить" onclick="btn_get_narabotka_skv_Click" />
                        </td>
                        
                    </tr>
                </table>
            </ContentTemplate>
        </asp:TabPanel>
    </asp:TabContainer>
    
    
  
</asp:Content>

