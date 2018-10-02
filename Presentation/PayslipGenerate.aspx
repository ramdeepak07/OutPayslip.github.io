<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="PayslipGenerate.aspx.cs" Inherits="OutPayslip.Presentation.PayslipGenerate" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css" />
    <script src="https://ajax.googleapis.com/ajax/libs/jquery/3.3.1/jquery.min.js"></script>
    <script src="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/js/bootstrap.min.js"></script>
    <script src="~/Style"></script>
    <title>OutPayslip</title>
</head>
<body>
    <div class="jumbotron text-center">
        <h1 style="color: orangered">My OutPayslip</h1>
        <p>A Perfect Solution for Bulk Payslip Generation!</p>
    </div>
    <div class="container">
        <h1>Our team Welcomes you to the world fo bulk PaySlip generation</h1>
        <p>
            This is a startup company where we provide different solution for bussiness oriented companies.
        </p>
        <ul>
            <li>Payroll Processing</li>
            <li>Payslip Generation</li>
            <li>Employee Management</li>
        </ul>
        our site is under construction but you can make use of stable payslip generator.
    </div>
    <div class="container">
        <form id="form1" runat="server">
            <div class="progress">
                <div class="progress-bar progress-bar-danger" role="progressbar" aria-valuenow="70"
                    aria-valuemin="0" aria-valuemax="100" style="width: 70%">
                </div>
            </div>
            <asp:FileUpload ID="Fileupload" runat="server" />
            <asp:Button ID="Button1" CssClass="btn btn-warning" runat="server" Text="Generate" OnClick="Generate" />
        </form>
    </div>
</body>
</html>
