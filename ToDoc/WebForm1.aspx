<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="WebForm1.aspx.cs" Inherits="ToDoc.WebForm1" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
        <div>
            Title :
            <asp:TextBox runat="server" ID="txtTitle" required />
            <br />
            Name :
            <asp:TextBox runat="server" ID="txtName" required />
            <br />
            <asp:Button runat="server" OnClick="Unnamed_Click" Text="Ok" />
        </div>
    </form>
</body>
</html>
