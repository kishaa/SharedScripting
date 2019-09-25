$User = "<id>"
$Token = "<Token>"

$Uri = "<API>"
$base64authinfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(("{0}:{1}" -f $User, $Token)))
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$response = Invoke-RestMethod -Method Get -ContentType application/json -Uri $Uri -Headers @{Authorization=("Basic {0}" -f $base64authinfo)}
$response.value |  Export-Csv -Path "<Path for CSV>"