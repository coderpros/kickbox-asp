# kickbox-asp

A [KickBox.io](https://kickbox.io) API wrapper for Classic ASP.

## How to use

- Create a free account at [Kickbox.io](https://kickbox.io).
- Sign up for an API Key.
- Update the KICKBOX_API_KEY constant in ~/Includes.config file to use your API Key.

### Verify a single email address

```vbs
Dim KickBox: Set KickBox = New KickBoxObject
Dim verificationResponse

With KickBox
    .ClientApiKey = KICKBOX_API_KEY
    Set verificationResponse = .VerifySingleEmail(string emailAddress))
End With
```

### Verify multiple email addresses

```vbs
Dim KickBox: Set KickBox = New KickBoxObject
Dim verificationResponse

With KickBox
    .ClientApiKey = KICKBOX_API_KEY
    Set verificationResponse = .VerifyBulkEmail(string emailAddressesSeparatedByNewline))
End With
```

### Check status of a bulk verification job

```vbs
Dim KickBox: Set KickBox = New KickBoxObject
Dim verificationResponse

With KickBox
    .ClientApiKey = KICKBOX_API_KEY
    Set verificationResponse = .CheckVerificationStatus(int JobId)
End With
```

*~/Default.asp is a functional example of how to use this library.*
