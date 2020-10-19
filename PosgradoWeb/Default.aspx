<%@ Page Title="Home Page" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="PosgradoWeb._Default" %>

<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <h3>Import / Export Database data to/from EXCEL file.</h3>
    <div>
        <table>
            <tr>
                <td>Select File : </td>
                <td>
                    <asp:FileUpload ID="FileUpload1" runat="server"/>
                </td>
                <td>
                    <asp:Button ID="btnImport" runat="server" Text="Import Data" OnClick="btnImport_Click"/>
                </td> 
            </tr>
        </table>
        <div>
            <br />
            <asp:Label ID="lblMessage" runat="server" Font-Bold="true"></asp:Label>
            <br />
            <asp:GridView ID="gvData" runat="server" AutoGenerateColumns="false">
                <EmptyDataTemplate>
                    <div style="padding:10px">
                        Data not found!
                    </div>
                </EmptyDataTemplate>
                <Columns>
                    <asp:BoundField HeaderText="Id" DataField="id"/>
                    <asp:BoundField HeaderText="Carnet de Identidad" DataField="ci"/>
                    <asp:BoundField HeaderText="Apellidos" DataField="apellidos"/>
                    <asp:BoundField HeaderText="Nombres" DataField="nombres"/>
                    <asp:BoundField HeaderText="1era Cuota" DataField="cuotaUno"/>
                    <asp:BoundField HeaderText="2do Cuota" DataField="cuotaDos"/>
                    <asp:BoundField HeaderText="3era Cuota" DataField="cuotaTres"/>
                    <asp:BoundField HeaderText="4ta Cuota" DataField="cuotaCuatro"/>
                    <asp:BoundField HeaderText="5ta Cuota" DataField="cuotaCinco"/>
                    <asp:BoundField HeaderText="6ta Cuota" DataField="cuotaSeis"/>
                    <asp:BoundField HeaderText="Curso" DataField="idCurso"/>
                </Columns>
            </asp:GridView>
        </div>
    </div>
</asp:Content>
