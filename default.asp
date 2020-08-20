<%@  language="vbscript" %>
<% Option Explicit %>
<!--#include file="Includes/config.asp"-->
<!--#include file="Includes/KickBoxObject.class.asp"-->
<%
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    Dim KickBox: Set KickBox = New KickBoxObject
    Dim verificationResponse

    With KickBox
        .ClientApiKey = KICKBOX_API_KEY

        If Len(Request.Form("EmailAddress")) > 0 Then
            Set verificationResponse = .VerifySingleEmail(Request.Form("EmailAddress"))
        ElseIf Len(Request.Form("EmailList")) > 0 Then
            Set verificationResponse = .VerifyBulkEmail(Request.Form("EmailList"))
        ElseIf Len(Request.Form("JobId")) > 0 Then
            Set verificationResponse = .CheckVerificationStatus(Request.Form("JobId"))
        End If
    End With
End If
%>
<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8" />
    <title>KickBox API Example</title>

    <link href="Content/bootstrap.min.css" type="text/css" rel="stylesheet" />
    <link href="Content/bootstrap-reboot.min.css" type="text/css" rel="stylesheet" />
</head>
<body>
    <div class="container-fluid">
        <h1>KickBox API Example</h1>
        <section>
            <div class="container">
                <div class="row">
                    <div class="col-md-6 col-sm-12">
                        <form name="VerificationForm" id="TwilioForm" method="post">
                            <div class="form-group">
                                <div class="col-auto">
                                    <label for="EmailAddress">Email Address</label>
                                    <input type="email" class="form-control" id="EmailAddress" name="EmailAddress" required="required">
                                </div>
                            </div>
                            <div class="form-group row">
                                <div class="col-sm-10">
                                    <button type="submit" class="btn btn-primary">Verify</button>
                                </div>
                            </div>
                        </form>
                    </div>
                    <div class="col-md-6 col-sm-12">
                        <form name="BulkVerificationForm" id="BulkVerificationForm" method="post">
                            <div class="form-group">
                                <div class="col-auto">
                                    <label for="EmailList">Email Addresses (one per line) <span class="badge badge-warning">Production Only!</span></label>
                                    <textarea class="form-control" name="EmailList" id="EmailList" rows="10"></textarea>
                                </div>
                            </div>
                            <div class="form-group row">
                                <div class="col-sm-10">
                                    <button type="submit" class="btn btn-primary">Upload</button>
                                </div>
                            </div>
                        </form>
                    </div>
                </div>
                <div class="row">
                    <div class="col-md-6 col-sm-12">
                        <form name="JobStatusForm" id="JobStatusForm" method="post">
                            <div class="form-group">
                                <div class="col-auto">
                                    <label for="JobId">Job Id <span class="badge badge-warning">Production Only!</span></label>
                                    <input type="text" class="form-control" name="JobId" id="JobId" value="804661" />
                                </div>
                            </div>
                            <div class="form-group row">
                                <div class="col-sm-10">
                                    <button type="submit" class="btn btn-primary">Check</button>
                                </div>
                            </div>
                        </form>
                    </div>
                    <div class="col-md-6 col-sm-12">
                        <% If IsObject(verificationResponse) Then %>
                            <%= verificationResponse.ToString %>
                        <% End If %>
                    </div>
                </div>
            </div>
        </section>
        <footer>
            <a href="https://coderpro.net" target="_blank"><img src="https://coderpro.net/images/logos/coderPro_logo_rounded_extra-90x90.webp" alt="coderPro.net logo" width="50" style="vertical-align:bottom;"></a> developed by <a href="https://coderpro.net" target="_blank">coderPro.net</a>
        </footer>
    </div>

    <script src="Scripts/jquery-3.5.1.min.js"></script>
    <script src="Scripts/bootstrap.min.js"></script>
</body>
</html>