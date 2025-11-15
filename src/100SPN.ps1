# Ensure Microsoft.Graph module is installed
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph)) {
    Install-Module Microsoft.Graph -Scope CurrentUser -Force
}
Import-Module Microsoft.Graph

# Connect with required scopes
Connect-MgGraph -Scopes "Application.ReadWrite.All","Directory.ReadWrite.All"

# Owner Object IDs (replace with your values)
$OwnerIds = @(
   #enter Owner details
)

# Output CSV path
$outCsv = "$PWD_PATH\BitsProjectApps_Secrets.csv"
New-Item -ItemType Directory -Force -Path (Split-Path $outCsv) | Out-Null

# Results collection
$results = @()

# Retry helper for Graph calls (handles throttling)
function Invoke-WithRetry {
    param(
        [scriptblock]$ScriptBlock,
        [int]$MaxRetries = 3
    )
    $retry = 0
    while ($true) {
        try {
            return & $ScriptBlock
        }
        catch {
            if ($_.Exception.Response.StatusCode -eq 429 -and $retry -lt $MaxRetries) {
                $retry++
                $delay = [math]::Min([math]::Pow(2,$retry)*1000,8000) # cap at 8s
                Write-Warning "Throttled. Retrying in $($delay/1000) seconds..."
                Start-Sleep -Milliseconds $delay
            }
            else {
                throw
            }
        }
    }
}

# Loop to create 100 applications
for ($i = 1; $i -le 100; $i++) {

    $displayName = "Bits Project - $i"
    Write-Host "Creating: $displayName"

    # Create the application (single-tenant)
    $app = Invoke-WithRetry { New-MgApplication -DisplayName $displayName -SignInAudience "AzureADMyOrg" }

    # Create service principal
    $sp = Invoke-WithRetry { New-MgServicePrincipal -AppId $app.AppId -DisplayName $displayName }

    # Add owners to application & SP (using ByRef cmdlets)
    foreach ($ownerId in $OwnerIds) {
        $body = @{ "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$ownerId" }
        try {
            New-MgApplicationOwnerByRef -ApplicationId $app.Id -BodyParameter $body
            New-MgServicePrincipalOwnerByRef -ServicePrincipalId $sp.Id -BodyParameter $body
        } catch {
            Write-Warning "Failed to assign owner $ownerId to $displayName"
        }
    }

    # Create a secret valid for 90 days (using addPassword endpoint)
    $start = (Get-Date).ToUniversalTime()
    $end = $start.AddDays(90)

    $body = @{
        passwordCredential = @{
            displayName   = "BitsProject-Secret-$i"
            startDateTime = $start
            endDateTime   = $end
        }
    } | ConvertTo-Json -Depth 3

    $pwd = Invoke-WithRetry {
        Invoke-MgGraphRequest -Method POST `
            -Uri "https://graph.microsoft.com/v1.0/applications/$($app.Id)/addPassword" `
            -Body $body -ContentType "application/json"
    }

    # Save record
    $results += [PSCustomObject]@{
        DisplayName        = $displayName
        ApplicationId      = $app.AppId
        AppObjectId        = $app.Id
        ServicePrincipalId = $sp.Id
        SecretDisplayName  = $pwd.passwordCredential.displayName
        SecretValue        = $pwd.secretText
        SecretStartUtc     = $pwd.passwordCredential.startDateTime
        SecretEndUtc       = $pwd.passwordCredential.endDateTime
    }

    # Delay to reduce throttling
    Start-Sleep -Seconds 2

    # Progress tracking
    Write-Progress -Activity "Creating Applications" -Status "Processing $displayName" -PercentComplete (($i/100)*100)
}

# Export CSV
$results | Export-Csv -Path $outCsv -NoTypeInformation -Encoding UTF8

Write-Host "Completed! CSV generated at: $outCsv"
