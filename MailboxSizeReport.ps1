Import-Module -name ExchangeOnlineManagement -Force
Connect-Exchangeonline

Function ConvertTo-Gb {
  <#
    .SYNOPSIS
        Convert mailbox size to Gb for uniform reporting.
  #>
  param(
    [Parameter(
      Mandatory = $true
    )]
    [string]$size
  )
  process {
    if ($size -ne $null) {
      $value = $size.Split(" ")

      switch($value[1]) {
        "GB" {$sizeInGb = ($value[0])}
        "MB" {$sizeInGb = ($value[0] / 1024)}
        "KB" {$sizeInGb = ($value[0] / 1024 / 1024)}
        "B"  {$sizeInGb = 0}
      }

      return [Math]::Round($sizeInGb,2,[MidPointRounding]::AwayFromZero)
    }
  }
}

$inboxes = Get-EXOMailbox -RecipientTypeDetails UserMailbox -Properties UserPrincipalName, DisplayName, ProhibitSendQuota | Select-Object DisplayName, UserPrincipalName, ProhibitSendQuota | Sort-Object DisplayName
$report = @()
foreach($inbox in $inboxes) {

    $mailboxSize = Get-EXOMailboxStatistics -Identity $inbox.UserPrincipalName | Select-Object TotalItemSize,TotalDeletedItemSize
    $username = $inbox.DisplayName
    $upn = $inbox.UserPrincipalName
    $mailboxMaxSize = $inbox.ProhibitSendQuota
    $mailboxCurrentSize = $mailboxSize.TotalItemSize.Value
    $mailboxDeletedSize = $mailboxSize.TotalDeletedItemSize.Value
    $username
    $mailTotalGB = (ConvertTo-GB -size $mailboxCurrentSize.ToString()) + (ConvertTo-GB -size $mailboxDeletedSize.ToString())
    $mailMaxGB = ConvertTo-GB -size $mailboxMaxSize

    $output = [pscustomobject]@{
          "Display Name" = $username
          "Email Address" = $UPN
          "Mailbox Size (GB)" = (ConvertTo-GB -size $mailboxCurrentSize.ToString())
          "Mailbox Deleted Size (GB)" = (ConvertTo-GB -size $mailboxDeletedSize.ToString())
          "Mailbox Total Size (GB)" = $mailTotalGB
          "Mailbox Max Size (GB)" = ConvertTo-GB -size $mailboxMaxSize
        }
        $report += $output
}
$outputfolder = "C:\Scripts"

if (!(Test-Path -path $outputfolder)) {
    New-Item -ItemType Directory -Force -Path $outputfolder | Out-Null
    
}

$report | export-csv -Path C:\Scripts\MailboxSizeReport.csv -NoTypeInformation