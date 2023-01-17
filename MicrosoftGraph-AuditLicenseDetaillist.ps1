# Connect to Microsoft Graph
# Define AppId, secret and scope, your tenant name and endpoint URL
# Sample code from https://adamtheautomator.com/microsoft-graph-api-powershell/

$AppId = 'xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx'
$AppSecret = 'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx'
$Scope = "https://graph.microsoft.com/.default"
$TenantName = "domainname.onmicrosoft.com"
$TenantID = 'xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx'
$SecureAppSecret = ConvertTo-SecureString -AsPlainText -String $AppSecret -force
$ServicePrincipalCred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $AppId, $SecureAppSecret

Connect-AzAccount -ServicePrincipal -Tenantid $TenantID -Credential $ServicePrincipalCred | Out-Null

$Url = "https://login.microsoftonline.com/$TenantName/oauth2/v2.0/token"

# Add System.Web for urlencode
Add-Type -AssemblyName System.Web

# Create body
$Body = @{
    client_id     = $AppId
    client_secret = $AppSecret
    scope         = $Scope
    grant_type    = 'client_credentials'
}

# Splat the parameters for Invoke-Restmethod for cleaner code
$PostSplat = @{
    ContentType = 'application/x-www-form-urlencoded'
    Method      = 'POST'
    # Create string by joining bodylist with '&'
    Body        = $Body
    Uri         = $Url
}

# Request the token!
$Request = Invoke-RestMethod @PostSplat

# Actual Microsoft Graph Query part
# Create header
$Header = @{
    Authorization = "$($Request.token_type) $($Request.access_token)"
}

$UserList = get-azaduser -Select licenseassignmentstates, assignedplans -AppendSelected
Function Get-LicenseSKU {
    param($AuthHeader)
    # GET All License SKUs
    # /subscribedSkus
    $Uri = "https://graph.microsoft.com/v1.0/subscribedSkus"

    $SkuListRequest = Invoke-RestMethod -Uri $Uri -Headers $Header -Method Get -ContentType "application/json"
    $Sku = @()
    $Sku += $SkuListRequest.value
    While ($NULL -ne $SkuListRequest.'@odata.nextlink') {
        $UserList += Invoke-RestMethod -uri $SkuListRequest.'@odata.nextlink' -Headers $Header -Method Get -ContentType 'application/json'
    }
    Return $Sku
}

$SKUPile=Get-LicenseSKU -AuthHeader $Header

$listtoReturn=[System.Collections.ArrayList]@()
Foreach ($User in $UserList) {
    $DirectFound = @($user.additionalproperties.licenseAssignmentStates | Where-object { $_.keys -notcontains 'assignedByGroup' })
    if ($NULL -ne $DirectFound) {
        $Skulist = @()
        Foreach ($D in $DirectFound) {
            $SkuList += $skupile[[system.array]::indexof($Skupile.skuid,($D.skuId))].skuPartNumber
        }
    }
    $ValuetoReturn=[pscustomobject]@{Name=$User.DisplayName;UserPrincipalName=$User.UserPrincipalName;DirectSKUList=($SkuList -join ',')}
    $ListToReturn.add($ValuetoReturn)
}
$listtoReturn
$ExportFile=".\UsersWithStaticO365License-$(Get-Date -Format 'MMddyyyy-hhmmss').csv"
$ListtoReturn | Export-CSV -Path $ExportFile -NoTypeInformation -Encoding UTF8