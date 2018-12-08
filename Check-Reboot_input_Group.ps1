param(
 [Parameter(Mandatory=$false)]
   [String] $OutCSV = $false
)
$cred = Get-Credential
$Server = Read-Host -Prompt 'Iserisci il Gruppo da verificare' | Get-ADGroupMember | select name -ExpandProperty name
#$server = Get-ADGroup -LDAPFilter "(samaccountname=hf-*)" -Server neg | Get-ADGroupMember | select name -ExpandProperty name
####################
$ThrottleLimit = 40
$jobname ="GetReboot"
####################
$risultato = @()
write-host (Get-Date) -ForegroundColor Yellow
$jobWRM = Invoke-Command -ComputerName $server -Credential $cred -ScriptBlock {
            Get-WmiObject win32_operatingsystem  | select csname, @{LABEL='LastBootUpTime' ;EXPRESSION={$_.ConverttoDateTime($_.lastbootuptime)}}
        } -JobName $jobname -ThrottleLimit $ThrottleLimit -AsJob

Get-Job -Id $jobWRM.Id | Wait-Job
$risultato = Get-Job  -Id $jobWRM.Id | Receive-Job    
Remove-Job -Id $jobWRM.Id
write-host (Get-Date) -ForegroundColor Yellow
if ($OutCSV -ne $false) {$risultato | export-csv $OutCSV -Delimiter "," -IncludeTypeInformation}
$risultato | Sort-Object -Property LastReboot

.NOTES

  Author: Angelo Malatacca
  Version: 17.12.05
  Date: 2017-12-05