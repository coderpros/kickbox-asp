<!--#include file="jsonObject.class.asp" -->
<%
Class KickBoxObject
    Private i_clientApiKey
    Dim JSON

    Public Sub Class_Initialize
        Set JSON = New JsonObject

        JSON.debug = False
    End Sub 

    Public Sub Class_Terminate
        Set JSON = Nothing
    End Sub 

    Public Property Get VerificationResponse
		VerificationResponse = i_verificationResponse
	End Property
	
	Public Property Let VerificationResponse(value)
		i_verificationResponse = value
	End Property

	Public Property Get ClientApiKey
		ClientApiKey = i_clientApiKey
	End Property
	
	Public Property Let ClientApiKey(value)
		i_clientApiKey = value
	End Property

    Public Property Get Response
		Response = i_Response
	End Property
	
	Public Property Let Response(value)
		i_Response = value
	End Property

    Public Function VerifySingleEmail(emailAddress)
        Dim d: Set d=Server.CreateObject("Scripting.Dictionary")
        Dim resp, verificationResponse

        d.Add "email", emailAddress
        d.Add "apikey", i_ClientApiKey
        
        Set Resp = CallApi("verify", "GET", d)
        Set verificationResponse = New ResponseModel

        With verificationResponse
            .AcceptAll = Resp("accept_all")
            .DidYouMean = Resp("did_you_mean")
            .Disposable = Resp("disposable")
            .Domain = Resp("domain")
            .Email = Resp("email")
            .Free = Resp("free")
            .Message = Resp("message")
            .Reason = Resp("reason")
            .Result = Resp("result")
            .Role = Resp("role")
            .SendEx = Resp("sendex")
            .Success = Resp("success")
            .User = Resp("user")
        End With    

        Set VerifySingleEmail = verificationResponse
        Set verificationResponse = Nothing
        Set Resp = Nothing
    End Function

    Public Function VerifyBulkEmail(emailList)
        Dim d: Set d=Server.CreateObject("Scripting.Dictionary")
        Dim Resp
        Dim batchVerificationResponse: Set batchVerificationResponse = New BatchResponseModel

        d.Add "apikey", i_clientApiKey
        d.Add "emailList", Request.Form("EmailList")

        Set Resp = CallApi("verify-batch", "PUT", d) 

        With batchVerificationResponse
            .Id = Resp("id")
            .Message = Resp("message")
            .Success = Resp("success")
        End With

        Set VerifyBulkEmail = batchVerificationResponse
        Set batchVerificationResponse = Nothing
        Set Resp = Nothing
    End Function

    Public Function CheckVerificationStatus(jobId)
        Dim d: Set d=Server.CreateObject("Scripting.Dictionary")
        Dim Resp
		Dim batchStatusResponse : Set batchStatusResponse = New StatusResponseModel

        d.Add "apikey", i_clientApiKey

        Set Resp = CallApi("verify-batch/" & jobId, "GET", d)

		With batchStatusResponse
			.CreatedAt = Resp("created_at")
			.DownloadUrl = Resp("download_url")
			.Duration = Resp("duration")
			.Id = Resp("id")
			.Message = Resp("message")
			.Name = Resp("name")
			.Success = Resp("success")
			.Status = Resp("status")
			
			If .Status = "completed" Then
				.Addresses = Resp("stats")("addresses")
				.Deliverable = Resp("stats")("deliverable")
				.Risky = Resp("stats")("risky")
				.SendEx = Resp("stats")("sendex")
				.Undeliverable = Resp("stats")("undeliverable")
				.Unknown = Resp("stats")("unknown")
			ElseIf .Status = "processing" Then
				.Addresses = Resp("progress")("addresses")
				.Deliverable = Resp("progress")("deliverable")
				.Risky = Resp("progress")("risky")
				.SendEx = Resp("progress")("sendex")
				.Total = Resp("progress")("total")
				.Undeliverable = Resp("progress")("undeliverable")
				.Unknown = Resp("progress")("unknown")
				.Unprocessed = Resp("progress")("unprocessed")
			End If

		End With

		Set CheckVerificationStatus = batchStatusResponse
		Set batchStatusResponse = Nothing
        Set Resp = Nothing  
    End Function

    '***************************************************
	' Helper Functions
    '***************************************************
    Private Function CallApi(method, httpMethod, inputData)
        Dim http: Set http = Server.CreateObject("WinHttp.WinHttpRequest.5.1")
        Dim url: url = KICKBOX_BASE_END_POINT & method
        Dim queryString
        Dim key

        queryString = ""

        If Not method = "verify-batch" Then
            For Each key In inputData.Keys
                queryString = queryString & "&" & key & "=" & inputData(key)
            Next
        Else
            querystring = queryString & "apikey=" & i_clientApiKey
        End If

        With http
            If Len(queryString) > 0 Then
                Call .Open(httpMethod, url & "?" & queryString, False) 
            Else
                Call .Open(httpMethod, url, False)
            End If

            If httpMethod = "POST" Then
                .SetRequestHeader "Content-Type", "application/json"
            ElseIf httpMethod = "PUT" Then
                .SetRequestHeader "Content-Type", "text/csv"
            Else
                .SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
            End If
                       
            If httpMethod = "POST" Then
                Call .Send(inputData)
            ElseIf httpMethod = "PUT" Then
                Call .Send(inputData("emailList"))
            Else
                Call .Send(queryString)
            End If
        End With

        Dim jsonOutput
        Set jsonOutput = DeserializeJson(http.ResponseText)
        
        If Left(jsonOutput("http"), 1) = "4" Then
            Call Err.Raise(12345, "JSON Error", jsonOutput("http") & "::" & jsonOutput("code") & ":" & jsonOutput("message"))
            
            Set jsonOutput = Nothing
            Set CallApi = Nothing
        Else
            Set CallApi = jsonOutput
        End If
    End Function

    Private Function DeserializeJson(jsonString)
        Dim jsonArr
        Set jsonArr = JSON.parse(jsonString)
        
        Set DeserializeJson = jsonArr
        Set jsonArr = nothing
    End Function
    '***************************************************
    '/Helper Functions
    '***************************************************
End Class

Class ResponseModel
    Private i_result, i_reason, i_role, i_free, i_disposable, i_acceptAll, i_didYouMean, i_email, i_user, i_domain, i_success, i_message, i_sendEx
    
    Public Property Get Result
	    Result = i_result
	End Property
	
	Public Property Let Result(value)
	    i_result = value
	End Property            
    
    Public Property Get Reason
	    Reason = i_reason
	End Property
	
	Public Property Let Reason(value)
	    i_reason = value
	End Property        

    Public Property Get Role
	    Role = i_role
	End Property
	
	Public Property Let Role(value)
	    i_role = value
	End Property        

    Public Property Get Free
	    Free =  i_free
	End Property
	
	Public Property Let Free(value)
	    i_free = value
	End Property        

    Public Property Get Disposable
	    Disposable = i_disposable
	End Property
	
	Public Property Let Disposable(value)
	    i_disposable = value
	End Property        

    Public Property Get AcceptAll
	    AcceptAll = i_acceptAll
	End Property
	
	Public Property Let AcceptAll(value)
	    i_acceptAll = value
	End Property        

    Public Property Get DidYouMean
	    DidYouMean = i_didYouMean
	End Property
	
	Public Property Let DidYouMean(value)
	    i_didYouMean = value
	End Property        

    Public Property Get SendEx
	    SendEx = i_sendEx
	End Property
	
	Public Property Let SendEx(value)
	    i_sendEx = value
	End Property        

    Public Property Get Email
	    Email = i_email
	End Property
	
	Public Property Let Email(value)
	    i_email = value
	End Property        

    Public Property Get User
	    User = i_user
	End Property
	
	Public Property Let User(value)
	    i_user = value
	End Property        

    Public Property Get Domain
	    Domain = i_domain
	End Property
	
	Public Property Let Domain(value)
	    i_domain = value
	End Property        

    Public Property Get  Success
	    Success = i_success
	End Property
	
	Public Property Let Success(value)
	    i_success = value
	End Property        

    Public Property Get Message
	    Message = i_message
	End Property
	
	Public Property Let Message(value)
	    i_message = value
	End Property        

    Public Function ToString()
        ToString = "Result: " & i_result & "<br />" & _
            "Reason: " & i_reason & "<br />" & _
            "Role: " & i_role & "<br />" & _
            "Free: " & i_free & "<br />" & _
            "Disposable: " & i_disposable & "<br />" & _
            "Accept All: " & i_acceptAll & "<br /> " & _
            "Did You Mean:" & i_didYouMean & "<br /> " & _
            "Normalized Email: " & i_email & "<br />" & _
            "User: " & i_user & "<br />" & _
            "Domain: " & i_domain & "<br />" & _
            "SendEx: " & i_sendEx & "<br />" & _
            "Success: " & i_success & "<br />" & _
            "Message: " & i_message & "<br />"
    End Function
End Class

Class BatchResponseModel
    Private i_id, i_result, i_message, i_success

    Public Property Get Id
	    Id =  i_id
	End Property
	
	Public Property Let Id(value)
	    i_id = value
	End Property        

    Public Property Get  Success
	    Success = i_success
	End Property
	
	Public Property Let Success(value)
	    i_success = value
	End Property        

    Public Property Get Message
	    Message = i_message
	End Property
	
	Public Property Let Message(value)
	    i_message = value
	End Property

    Public Function ToString()
        ToString = "Id: " & i_id & "<br />" & _
            "Success: " & i_success & "<br />" & _
            "Message: " & i_message & "<br />"
    End Function
End Class

Class StatusResponseModel
    Private i_id, i_name, i_createdAt, i_status, i_error, i_duration, i_success, i_message
    Private i_downloadUrl, i_deliverable, i_undeliverable, i_risky, i_unknown, i_sendex, i_addresses
	Private i_unprocessed, i_total

    Public Property Get Id
	    Id =  i_id
	End Property
	
	Public Property Let Id(value)
	    i_id = value
	End Property

        Public Property Get Name
	    Name =  i_name
	End Property
	
	Public Property Let Name(value)
	    i_name = value
	End Property      
    
    Public Property Get CreatedAt
	    CreatedAt = i_createdAt
	End Property
	
	Public Property Let CreatedAt(value)
	    i_createdAt = value
	End Property    
    
    Public Property Get Status
	    Status =  i_status
	End Property
	
	Public Property Let Status (value)
	    i_status = value
	End Property        

    Public Property Get Error
	    Error =  i_error
	End Property
	
	Public Property Let Error(value)
	    i_error = value
	End Property        

    Public Property Get Duration
	    Duration = i_duration
	End Property
	
	Public Property Let Duration(value)
	    i_duration = value
	End Property  
    
    Public Property Get Success
	    Success =  i_success
	End Property
	
	Public Property Let Success(value)
	    i_success = value
	End Property        

    Public Property Get Message
	    Message =  i_message
	End Property
	
	Public Property Let Message(value)
	    i_message = value
	End Property        

    Public Property Get DownloadUrl
	    DownloadUrl =  i_downloadUrl
	End Property
	
	Public Property Let DownloadUrl(value)
	    i_downloadUrl = value
	End Property   

	Public Property Get Deliverable
	    Deliverable =  i_deliverable
	End Property
	
	Public Property Let Deliverable(value)
	    i_deliverable = value
	End Property            

	Public Property Get Undeliverable
	    Undeliverable = i_undeliverable
	End Property
	
	Public Property Let Undeliverable(value)
	    i_undeliverable = value
	End Property        

	Public Property Get Risky
	    Risky =  i_risky
	End Property
	
	Public Property Let Risky(value)
	    i_Risky = value
	End Property        

	Public Property Get Unknown
	    Unknown =  i_unknown
	End Property
	
	Public Property Let Unknown(value)
	    i_unknown = value
	End Property        

	Public Property Get SendEx
	    SendEx =  i_sendEx
	End Property
	
	Public Property Let SendEx(value)
	    i_sendEx = value
	End Property 
	
	Public Property Get Addresses
	    Addresses =  i_addresses
	End Property
	
	Public Property Let Addresses(value)
	    i_addresses = value
	End Property

	Public Property Get Total
	    Total =  i_total
	End Property
	
	Public Property Let Total(value)
	    i_total  = value
	End Property

	Public Property Get Unprocessed
	    Unprocessed =  i_unprocessed
	End Property
	
	Public Property Let Unprocessed(value)
	    i_unprocessed = value
	End Property

	Public Function ToString
		ToString = "Id: " & i_id & "<br />" & _
			"Name: " & i_name & "<br />" & _
			"Created at: " & i_createdAt & "<br />" & _
			"Status: " & i_status & "<br />" & _
			"Error: " & i_error & "<br />" & _
			"Duration: " & i_duration & "<br>" & _
			"Success: " & i_success & "<br />" & _
			"Message " & i_message & "<br />" & _
			"Download Url: <a href='" & i_downloadUrl & "'>" & i_downloadUrl & "</a>" & "<br />" & _
			"Deliverable:" & i_deliverable & "<br />" & _
			"Undeliverable: " & i_undeliverable & "<br />" & _
			"Unprocessed: " & i_unprocessed & "<br />" & _
			"Total: " & i_total & "<br />" & _
			"Risky: " & i_risky & "<br />" & _
			"Unknown: " & i_unknown & "<br />" & _
			"SendEx: " & i_sendex & "<br />" & _
			"Addresses: " & i_addresses & "<br />"
	End Function
End Class
%>