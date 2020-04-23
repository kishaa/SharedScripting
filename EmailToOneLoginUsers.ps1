#Get an Outlook application object
#$TimeStamp = Get-Date -Format g
 

function ExtractFileName([String] $filePath)
{
	$fileName = [System.IO.Path]::GetFileNameWithoutExtension($filePath)
	$fileExt = [System.IO.Path]::GetExtension($filePath)
	$result = [string]::format("{0}{1}", $fileName, $fileExt)
	return $result
}

function SendEmail([string] $firstName, [string] $lastName, [string] $email, [string] $pwd)
{
    $o = New-Object -com Outlook.Application
    #$imageFile = ExtractFileName -filePath C:\Users\sandeb\Documents\WindowsPowerShell\91722004.PNG
 
    $mail = $o.CreateItem(0)
    $mail.importance = 1
    $mail.subject = "One Login Credential"
    $mail.To = $email
    #$mail.CC = "sourabh.awasthi@capgemini.com ; sandipan.deb@capgemini.com"
    $mail.HTMLBody = "<html>
    <p>Hi $firstName $lastName,</p>
 
    <p>You’re invited to register to Royal Mail’s OneLogin system.</br></p>
    <p>Click <a href='https://royalmail.onelogin.com/login'>here</a> and use your e-mail address as the username and the password: $pwd</br></p>
    <p>Please keep a secure record of this password as it is auto generated and cannot be changed.</br></p>

    <p>In the coming weeks, applications will be made accessible over the internet; first of which are RMG Jira and Confluence (not Digital Labs).</br></p>
    <p>OneLogin will be used to secure the applications using 2FA (2-Factor Authentication).</br></p>
    <p>For this to work you need to register a secondary authentication device. This can be a Smartphone or SMS (UK phone numbers only).</br></p>
    <p style='color:red;'>WITHOUT EITHER METHOD YOU WILL NOT BE ABLE TO GAIN ACCESS!</br></p>

    <p>To register 2FA, once logged in, click on your name in the top right corner and select `Profile`. Select the plus symbol next to `2-Factor Authentication`.</br></p>
    <p><b>For Smartphone</b>, download the ‘OneLogin Protect’ app and scan in the QR Code on the `Add` 2-Factor Method` screen.</br></p>
    <p><b>For SMS</b>, simply select `OneLogin SMS` in the `Add 2-Factor Method` screen.</br></p>
    <p><b>Note:</b> we need to have the phone number already on file for this to work, please e-mail back if this is to be updated.</br></p>
    <br></br>
    <p>If you have any issues / queries please respond to this e-mail.</br></p>

    <p>Regards,</br><p>
    <p>RMG Enterprise Tooling</br></p> 


    </html>

    <img src='cid:$imageFile'>" 

    $mail.Send()
}

$Excel = New-Object -ComObject Excel.Application
$path = "C:\Users\sandeb\Documents\OLC.xlsx"
$workBook1 = $Excel.Workbooks.Open("$path")
$workSheet1 = $workBook1.Sheets.Item(1)
$Range = $workSheet1.UsedRange
[int]$RowValue = $Range.Rows.Count
For($Row=2; $Row -le $RowValue; $Row++)
{
    $firstName = $workSheet1.Cells.Item($Row, 1).text
    $lastName = $workSheet1.Cells.Item($Row, 2).text
    $email = $workSheet1.Cells.Item($Row, 3).text
    $pwd = $workSheet1.Cells.Item($Row, 4).text

    SendEmail -firstName $firstName -lastName $lastName -email $email -pwd $pwd
}



