
#AppflowReadMessageByMessageId.ps1
# This will find a message by using Graph with Applicaiton client secret oath flow. Invoke-RestMethod is used to submit the Graph call.

# Define the necessary variables
cls

# TODO: Change to a upn
$UserUpn = ""                  # Example:  someone@contoso.com

# TODO: Change to a message ID
$messageId = ""                # Example: AC80000000024979BA068F17735mscom_mkt_prod6@mails.microsoft.com

#The AppID or ClientID  - TODO: Set the below value
$AppId = ""                    # Example: fd005f81-b3c2-4575-a212-333baab22708

#The tenant id  - TODO: Set the below value
$tenantID = ""                 # Example: dd5384f6-4d6e-4bfb-2182-5d5a22d3b5ae

#The App Secret  - TODO: Set the below value
$AppSecret = ""                # Example: ~AO8Q~TUHdzRNsHmScL-EW4284zyNf~R4izbpi

Write-Host "Start:"

#Create the request body to obtain token 
$ReqTokenBody = @{
    grant_type    = "client_credentials"
    Scope         = "https://graph.microsoft.com/.default"
    client_Id     = $AppId
    client_secret = $AppSecret
}
Write-Host "ReqTokenBody:"
$ReqTokenBody 

 

#Get Token response
Write-Host "Getting Token Response:"
$TokenResponse = Invoke-RestMethod -Method POST -Uri "https://login.microsoft.com/$tenantID/oauth2/v2.0/token" -Body $ReqTokenBody

#Write-Host "TokenResponse:"
#$TokenResponse 

$AuthHeader = @{'Authorization' = "Bearer $($TokenResponse.access_token)" }
#Write-Host "AuthHeader:"
#$AuthHeader 

Write-Host "Getting the token:"
 
# Make the request to get the token
$TokenResponse = Invoke-RestMethod -Method POST -Uri "https://login.microsoft.com/$tenantID/oauth2/v2.0/token" -Body $ReqTokenBody
#Write-Host "TokenResponse:"
#$TokenResponse 

#Write-Host "TokenResponse.access_token:"
#$TokenResponse.access_token

$AuthHeader = @{'Authorization' = "Bearer $($TokenResponse.access_token)" }
Write-Host "AuthHeader:"
$AuthHeader 


Write-Host "Doing the Graph call:"
# Make the GET request to find the email by message ID
 
$emailResponse = Invoke-RestMethod -Method Get -Uri "https://graph.microsoft.com/v1.0/users/$UserUpn/messages?$filter=internetMessageId eq '$messageId'" -Headers $AuthHeader

$emailResponse
Write-Host "Done"
 

 
