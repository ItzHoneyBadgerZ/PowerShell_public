<#
  .DESCRIPTION
    This requires an application in Entra AD with permission "Calendars.ReadWrite". Additionally, ensure the "$calendarEvents" csv includes the startDate, endDate, Location, and Subject. Obviously,
    you can modify to your specific needs, there is a lot more that can be added to the calendar invite see MS Doc: https://learn.microsoft.com/en-us/graph/api/resources/event?view=graph-rest-1.0#properties

  .CREATEDBY
    ItzHoneyBadgerz - 01/08/2024
#>

# Set your application details
$tenantId = "xxxx"
$clientId = "xxxx"
$clientSecret = "xxxx"
$calendarEvents = "C:\temp\calendarEvents.csv" #must include startDate, EndDate, Location (we used HOME), and Subject
$userEmails = "C:\temp\users.txt"

$tokenUrl = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"
$body = @{
    client_id     = $clientId
    scope         = "https://graph.microsoft.com/.default"
    client_secret = $clientSecret
    grant_type    = "client_credentials"
}
#get token. Note that this only lasts for a limited time, if this will take longer than 2 hours, utilize a refresh token. I'll add this process at a later time.
$response = Invoke-RestMethod -Uri $tokenUrl -Method POST -ContentType "application/x-www-form-urlencoded" -Body $body

$headers = @{
    Authorization = "Bearer $($response.access_token)"
    "Content-Type" = "application/json"
}

foreach($email in $userEmails){
    $graphApiUrl = "https://graph.microsoft.com/v1.0/users/$email/events"
    foreach($event in $calendarEvents){
        # Event details
        $eventSubject = $event.Subject
        $eventStart = $event.StartDate
        $eventEnd = $event.EndDate
        $eventLocation = $event.Location
        $showAs = "oof"

        $body = @{
            subject = $eventSubject
            start = @{
                dateTime = $eventStart
                timeZone = "Central Standard Time"
            }
            end = @{
                dateTime = $eventEnd
                timeZone = "Central Standard Time"
            }
            location = @{
                displayName = $eventLocation
            }
            showAs = $showAs
            categories = @("Purple Category")

        } | ConvertTo-Json

        Invoke-RestMethod -Uri $graphApiUrl -Method POST -Headers $headers -Body $body

    }
}
