<%@ Page Title="" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="HomePage.aspx.cs" Inherits="ZadanieRekrutacyjneRS.HomePage" %>
<asp:Content ID="Content1" ContentPlaceHolderID="MainContent" runat="server">
    <style>
        .container {
            margin-top: 60px;
            margin-left: 60px;
        }
        .create_button {
            background-color: #008CBA;
            border-radius: 8px;
            border: 1px solid #008CBA;
            width: 150px;
            height: 50px;
            color: #fff;
            font-size: 15px;
        }
        .create_button:hover {          
            color: 	#DC143C; 
            opacity: 0.8;
        }
    </style>
    <h1>Przycisk do tworzenia pliku excela: </h1>
    <div class="container">
    <asp:Button CssClass="create_button" ID="Create" runat="server" OnClick="Create_Click" Text="Create Stylesheet" />
        </div>
</asp:Content>
