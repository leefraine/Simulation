icacls "C:\inetpub\wwwroot\yourfolder\uploads" /grant "IIS AppPool\YourAppPoolName":(M)





# Create the credential object
$securePassword = ConvertTo-SecureString "PlainTextPassword" -AsPlainText -Force
$credential = New-Object System.Management.Automation.PSCredential ("username", $securePassword)

# Start the BITS transfer upload with custom headers and credentials
Start-BitsTransfer `
  -Source "C:\Path\to\file.bin" `
  -Destination "http://<IIS-SERVER>/upload.asp" `
  -TransferType Upload `
  -Credential $credential `
  -Authentication Basic `
  -CustomHeaders @("BITS_POST: 1", "User-Agent: Microsoft BITS/7.8")


