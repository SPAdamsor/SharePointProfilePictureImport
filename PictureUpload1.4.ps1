##The information that is  provided "as is" without warranty of any kind. We disclaim all warranties, either express or implied, including the warranties of merchantability and fitness for a particular purpose. In no event shall Microsoft Corporation or its suppliers be liable for any damages whatsoever including direct, indirect, incidental, consequential, loss of business profits or special damages, even if Microsoft Corporation or its suppliers have been advised of the possibility of such damages. Some states do not allow the exclusion or limitation of liability for consequential or incidental damages. Therefore, the foregoing limitation may not apply.
##Using Dirsync outside of SharePoint to import Profile Pictures
##Author:adamsor; https://adamsorenson.com
##Version: 1.3
##1.0 Using Profile changes script
##1.1 Improved performance.  Filtered to only include users with thumbnailphoto.  Only pulling thumbnailphoto and sAMAccountName.
##1.1 Summary added.  Logging improved.
##1.2 Fixed UploadPhoto progress bar.  Fixed location for photos with creating that folder if not already created.  Added the last line to create the thumbnails.
##1.3 Fixed performance issue when using larger pictures. Fixed issue with existing users failing since samaccountname is not included.  Fixed logging issue that added 'UR'.  No longer needing a DNLookup file.
##1.4 Added new variable to declare the Idenitifier claim.  This will work with SAML.

Add-PSSnapin "Microsoft.SharePoint.PowerShell" -ErrorAction SilentlyContinue

$Location = "C:\Dirsync\"
#First time running, just run "DirSync" then "UploadPicture $adusers"
#Update RootDSE to match your domain
$RootDSE = [ADSI]"LDAP://dc=contoso,dc=com"
$site	  = Get-SpSite http://MySiteHostURL
#This is the prefix Identifier.  This will be "contoso\" or "i:0e.t|ADFS|"
#e = UPN; 5 = Email
$domain = "domain\" 
#For windows put "sAMAccountName"; SAML should be something like "mail" or "userPrincipalName"
$IdentifierClaim = "sAMAccountName"
#This is for domains that may not resolve the short name.  We can have a domain controller, FQDN, or shortname for the domain.
$DCorDomainName = "Contoso.com"
#This will write the pictures to the folder specified in $location
$write2disk = $false
#LDAP filter that is currently set to pull in users with thumbnailphoto and not disabled users.
$LDAPFilter = "(&(objectCategory=person)(objectclass=user)(thumbnailphoto=*)(!(userAccountControl:1.2.840.113556.1.4.803:=2)))"
#Set $UseDifferentSvcAccount to true to be prompted for a different service account.  False will use the user that is running the script to connect to AD.
$UseDifferentSvcAccount = $false
#Time out for the LDAP connection.  By default this is 30 seconds but we will need to increase it in some cases
[int]$TO=2

#Do not change below this line
$cookiepath = $Location+"cookie.bin"
[xml]$DNlookup = Get-Content -Path $Location"DNLookup.xml"
$log = $Location+"out.log"
$fileloc = $Location+"DNLookup.xml"
$username = $null
$global:ADUsers = $null
#Using default paritionID
$partitionID = "0C37852B-34D0-418E-91C6-2AC25AF4BE5B"
$context  = Get-SPServiceContext($site) 
$pm       = new-object Microsoft.Office.Server.UserProfiles.UserProfileManager($context, $true) 
$PhotoFolder = $site.RootWeb.GetFolder("User Photos")

#This code calls to a Microsoft web endpoint to track how often it is used. 
#No data is sent on this call other than the application identifier
Add-Type -AssemblyName System.Net.Http
$client = New-Object -TypeName System.Net.Http.Httpclient
$cont = New-Object -TypeName System.Net.Http.StringContent("", [system.text.encoding]::UTF8, "application/json")
$tsk = $client.PostAsync("https://msapptracker.azurewebsites.net/api/Hits/2cec8200-3f12-4e13-a5c8-fef7e0b1ad09",$cont)



#Try/Catch for looking for the Profile Pictures folder which is not created OOB
Try
{
"Trying to get the Profile Pictures Folder"| out-file $log -Append -noclobber
$site.RootWeb.GetFolder("User Photos/Profile Pictures")
$PhotoFolder = $site.RootWeb.GetFolder("User Photos/Profile Pictures")

    If($PhotoFolder.Exists -eq $false)
    {
        Try
        {
        $site.RootWeb.GetFolder("User Photos").subfolders.Add("Profile Pictures")
        "Profile Pictures successfully created" | out-file $log -Append -noclobber
        $PhotoFolder = $site.RootWeb.GetFolder("User Photos/Profile Pictures")
        }
        Catch
        {
        "Unable to create the Profile Pictures.  Check log below and permissions"| out-file $log -Append -noclobber
        throw
        }
        }
    Else
    {
    "Successfully loaded the Profile Pictures"| out-file $log -Append -noclobber
    }
}
Catch
{
"Unable to get UserPhotos/Profile Pictures.  Trying to create." | out-file $log -Append -noclobber
    Try
    {
    $site.RootWeb.GetFolder("User Photos").subfolders.Add("Profile Pictures")
    "Profile Pictures successfully created" | out-file $log -Append -noclobber
    $PhotoFolder = $site.RootWeb.GetFolder("User Photos/Profile Pictures")
    }
    Catch
    {
    "Unable to create the Profile Pictures.  Check log below and permissions"| out-file $log -Append -noclobber
    throw
    }
}
$files = $PhotoFolder.Files
Add-Type -AssemblyName System.DirectoryServices.Protocols

If($UseDifferentSvcAccount -eq $true)
{
if ($cred -eq $null) { $cred=(Get-Credential).GetNetworkCredential() }
}



 function Byte2DArrayToString
{
    param([System.DirectoryServices.Protocols.DirectoryAttribute] $attr)

    $len = $attr[0].length
    $val = [string]::Empty
   
    for($i = 0; $i -lt $len; $i++)
    {
         $val += [system.text.encoding]::UTF8.GetChars($attr[0][$i])

    }
    return $val

}

   function Byte2DArrayToBinary
{
    param([System.DirectoryServices.Protocols.DirectoryAttribute] $attr)
    
    $len = $attr[0].length
    #$val = New-Object [Byte] 8
   
    for($i = 0; $i -lt $len; $i++)
    {
         $val += @([byte]::Parse($attr[0][$i]))

    }
    return $val

}


function Dirsync
{
Write-Progress -Activity "Querying AD..." -Status "Please wait."
If (Test-Path $cookiepath –PathType leaf) {[byte[]] $Cookie = Get-Content -Encoding byte –Path $cookiepath}else {$Cookie = $null}
$global:ADUsers = @()
#$LDAPConnection = New-Object System.DirectoryServices.Protocols.LDAPConnection($RootDSE.dc) 
$LDAPConnection = New-Object System.DirectoryServices.Protocols.LDAPConnection($DCorDomainName)
If($cred -ne $null)
{
$LDAPConnection.Credential=$cred
}
$Request = New-Object System.DirectoryServices.Protocols.SearchRequest($RootDSE.distinguishedName, $LDAPFilter, "Subtree", $null) 
$Request.Attributes.Add("thumbnailphoto")
$Request.Attributes.Add($IdentifierClaim)
$DirSyncRC = New-Object System.DirectoryServices.Protocols.DirSyncRequestControl($Cookie, [System.DirectoryServices.Protocols.DirectorySynchronizationOptions]::IncrementalValues, [System.Int32]::MaxValue) 
$Request.Controls.Add($DirSyncRC) | Out-Null 
$LDAPConnection.Timeout = New-TimeSpan -Minutes $TO
$Response = $LDAPConnection.SendRequest($Request)
$MoreData = $true
while ($MoreData) {
    $Response.Entries | ForEach-Object { 
        write-host $_.distinguishedName 
        $global:ADUsers += $_ 
    }
    ForEach ($Control in $Response.Controls) { 
        If ($Control.GetType().Name -eq "DirSyncResponseControl") { 
            $Cookie = $Control.Cookie 
            $MoreData = $Control.MoreData 
        } 
    } 
    $DirSyncRC.Cookie = $Cookie 
    $Response = $LDAPConnection.SendRequest($Request) 
}
Set-Content -Value $Cookie -Encoding byte –Path $cookiepath
$global:ADUsers #| export-clixml C:\dirsync\aduser.clixml
return $global:adusers
}


Function GetUsername
{
    param($aduser)
    $sam = ($aduser.DistinguishedName | dnlookup)
    #logging fix.
    If($sam.count -gt 1)
    {
        $sam=$sam[1]
        $username = $domain + $sam
        return $username
    }
    
    If($sam -ne $null)
    {    
        $username = $domain + $sam
        return $username
    }
return $false
}

Function DnLookup
{
    param([Parameter(ValueFromPipeline=$true)]$DN)
    $lookup=$null
    #DNLookup check to see if the file is created.
    If($DNlookup -eq $null)
    {
        Try
        {
            "Trying to create DNlookup.xml" | out-file $log -Append -noclobber
            $xmlpath = $Location+"DNlookup.xml"
            $xml = New-Object System.XML.XmlTextWriter($xmlpath,$null)
            $xml.Formatting = "Indented"
            $xml.Indentation = 1
            $xml.IndentChar = "`t"
            $xml.WriteStartDocument()
            $xml.WriteProcessingInstruction("xml-stylesheet", "type='text/xsl' href='style.xsl'")
            $xml.WriteStartElement("Users")
            $xml.WriteStartElement("UR")
            $xml.WriteElementString("dn",[string]$DN)

            $dsam=$aduser.Attributes[$IdentifierClaim]
            $sam=Byte2DArrayToString -attr $dsam

            $xml.WriteElementString($IdentifierClaim,[string]$sam)

            $xml.WriteEndElement()
            $xml.WriteEndElement()
            $xml.WriteEndDocument()
            $xml.Flush()
            $xml.close()
            [xml]$global:DNlookup = Get-Content -Path $Location"DNLookup.xml"
            "XML Created Successfully" | out-file $log -Append -noclobber
        }
        Catch
        {
            "Failed to create XML file" | out-file $log -Append -noclobber
            $PSItem.Exception | out-file $log -Append -noclobber
            Throw
        }
        Return $sam
    }

    $lookup=$DNlookup.Users.ur | where {$_.dn -eq $DN}

    If ($lookup -eq $null)
    {
       #$newDN=$DNLookup.CreateElement("UR")
       $olddn = @($DNlookup.users.UR)[0]
       $newDN=$olddn.clone()
       If($aduser.Attributes[$IdentifierClaim] -eq $null)
       {
            $adsi = [adsisearcher]""
            $adsi.SearchRoot.Path = $RootDSE.path
            $adsi.filter = "(distinguishedName=$dn)"
            $adsiuser = $adsi.FindOne()
            $sam = $adsiuser.Properties.$($IdentifierClaim.tolower())
       }
       Else
       {
            $dsam=$aduser.Attributes[$($IdentifierClaim.ToLower())]
            $sam=Byte2DArrayToString -attr $dsam
       }
       $newDN.dn = [string]$DN
       $newDN.$IdentifierClaim = [string]$sam
       $DNlookup.Users.AppendChild($newDN) 
       $dnlookup.Save($fileloc)
       return $sam
    }
    $sam = $lookup.$IdentifierClaim
    Return $sam
}


Function FindUserProfile
{
param([Parameter(ValueFromPipeline=$true)]$username)

$UserProfile = $pm.GetUserProfile($username)
return $userprofile
}


Function write2disk
{
param($filename,$un,$bin)

Try
{
[io.file]::WriteAllBytes($location+$filename,$bin)
}
Catch
{
"$un failed to write picture to disk but the upload probably worked"
}

}

Function UploadPicture
{
    param([Parameter(ValueFromPipeline=$true)]$adusers)
    $date = Get-Date
    [int]$i=1
    [int]$e=0
    [int]$s=0
    $c = $adusers.count
    "New upload started at $date for $($adusers.Count) users" | out-file $log -Append -noclobber

    Foreach ($ADUser in $adusers)
    {
        $decoded=@()
        [int]$p = ($i/$c)*100
        Write-Progress -Activity "Processing User Photos" -CurrentOperation $aduser.DistinguishedName -PercentComplete $p -Status "$i of $c"
        $un= GetUsername $ADUser
        If($un -eq $false)
        {
            "Could not find $($aduser.DistinguishedName)"| out-file $log -Append -noclobber
            $i++
            $e++
            Continue
        }
        try 
        {
            $UPAProfile=GetUsername $ADUser | FindUserProfile
            
        } 
        catch 
        {
            "Could not find User Profile for $un" | out-file $log -Append -noclobber
            $i++
            $e++
            Continue
        }
        If($ADuser.Attributes.thumbnailphoto -eq $false)
        {

            "$un does not have thumbnailphoto update"| out-file $log -Append -noclobber
            $i++
            
            Continue

        }

        $filename = "$($partitionID)_$($UPAProfile.recordid).jpg"
        try
        {
        $photo = $aduser.Attributes.thumbnailphoto
        #[byte[]]$bin = Byte2DArrayToBinary -attr $photo
        "Uploading $un"| out-file $log -Append -noclobber
        #Uploading with $true for overwrite.
        $files.add("User Photos/Profile Pictures/" + $filename,$photo[0],$true)
        "Upload successful for $un" | out-file $log -Append -noclobber
        $i++
        $s++
        }
        Catch
        {
        "$un did not upload" | out-file $log -Append -noclobber
        $i++
        $e++
        Continue
        }
        If($write2disk -eq $true)
        {
        write2disk $filename $un $photo[0]
        }

    }

#Summary to the logs
$fdate = Get-Date
$ddate = $fdate - $date
"Summary: Upload completed at $fdate(took $ddate).  $c users imported from AD. $e errors. $s user photos successfully uploaded." | out-file $log -Append -noclobber

}
Dirsync
UploadPicture $ADUsers
Update-SPProfilePhotoStore -MySiteHostLocation $site -CreateThumbnailsForImportedPhotos $true