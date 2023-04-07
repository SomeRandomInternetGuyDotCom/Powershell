Function New-EncryptedCSV
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [System.Object]
        $InputCSV,
        [Parameter(Mandatory)]
        [string]
        $OutputPath,
        [Parameter(Mandatory)]
        [string]
        $SecuredFileName,
        [Parameter()]
        $CertificateName = $null
    )

If ($null -eq $CertificateName)
{
    Write-Output "Detected that you did not input a certificate, generating a certificate now..."
    $NewCertName = Read-Host "What is the friendly name for your new cert?"
    $NewCert = New-SelfSignedCertificate -DnsName $NewCertName -CertStoreLocation "Cert:\CurrentUser\My" -KeyUsage KeyEncipherment,DataEncipherment, KeyAgreement -Type DocumentEncryptionCert
    $CertificateName = $newcert.DnsNameList.Unicode
}

Do {
    $Cert = Get-ChildItem -Path Cert:\CurrentUser\My -DnsName $CertificateName -DocumentEncryptionCert

    If ($Null -eq $Cert )
    {
        $CertificateName = Read-Host "It seems we couldn't find that cert, can you please input the cert name again"
    }
}
While ($null -eq $Cert)

Write-Output "Using the following certificate to encrypt the data: "
$Cert

$Confirm = Read-Host "Is this correct? Y/N"

If ($Confirm -notlike 'y')
{
    Write-output "Exiting script as this is not the right cert.  Please rerun the script and select the proper cert name"
    Pause
    Exit
}


#Check that the output path is valid
Try{
    $PathCheck = Get-Item $OutputPath 
    If ($PathCheck.Attributes -ne "Directory")
    {
        Write-Error -Exception -Category InvalidData -ErrorAction Stop
        
    }
}
Catch {Write-Error -Message "The Path you entered is not a valid directory, please try again."
        exit
    }

#Make sure the output path is properly formatted to concatenate with the filename
If ($OutputPath.EndsWith('\') -eq $false)
{
    $OutputPath = $OutputPath + '\'
}

#Check that the input file is valid
Try{
    $InputFileCheck = Get-Item $InputCSV -ErrorAction Stop 
}
Catch {Write-Error -Message "There was an error with the input file, please verify and retry"
        Exit    
    }

#Generate the output path
$outFile = $OutputPath + $SecuredFileName + ".encrypted"

#Load the file to be encrypted 
$dataToEncrypt = Import-Csv -Path $InputCSV

$dataToEncrypt = $dataToEncrypt | ConvertTo-Csv -NoTypeInformation

Write-Output "Generating encrypted file..."

Protect-CmsMessage -To $cert.Subject -OutFile $outFile -Content $dataToEncrypt

}

Function Get-EncryptedCSV
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [String]
        $EncryptedFile
    )

#Validate that the file is working properly
try {
    $DecryptedData = Unprotect-CmsMessage -Path $EncryptedFile -ErrorAction Stop
}
catch [System.Security.Cryptography.CryptographicException]
{
    Write-Error "Looks like you don't have the proper cert installed to decrypt this file.  `nTry again after you imported the proper cert to your account."
    Exit
}
Catch [System.SystemException]
{
    Write-Error "The encrypted CSV file you entered could not be found, please try the path again."
    Exit
}

#Convert the file into CSV format for output
$DecryptedCSV = $DecryptedData | ConvertFrom-Csv 
    
Return $DecryptedCSV
}

