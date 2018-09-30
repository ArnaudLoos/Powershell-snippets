<#
.SYNOPSIS
    .
.DESCRIPTION
    This script will find local administrators of client computers in your
    domain and will same them as CSV file in current directory.

.PARAMETER Path
    This will be the DN of the OU or searchscope. Simply copy the DN of OU
    in which you want to query for local admins. If not defined, the whole
    domain will be considered as search scope.

.PARAMETER ComputerName
    This parametr defines the computer account in which the funtion will
    run agains. If not specified, all computers will be considered as search
    scope and consequently this function will get local admins of all 
    computers. You can define multiple computers by utilizing comma (,).

.EXAMPLE
    C:\PS> Get-LocalAdminToCsv
    
    This command will get local admins of all computers in the domain.

    C:\PS> Get-LocalAdminToCsv -ComputerName PC1,PC2,PC3

    This command will get local admins of PC1,PC2 and PC3.

    C:\PS> Get-LocalAdminToCsv -Path "OU=Computers,DC=Contoso,DC=com"

.NOTES
    Author: Mahdi Tehrani
    Date  : February 18, 2017   
#>


Import-Module activedirectory
Clear-Host
function Get-LocalAdminToCsv {
    Param(
            $Path          = (Get-ADDomain).DistinguishedName,   
            $ComputerName  = (Get-ADComputer -Filter * -Server (Get-ADDomain).DNsroot -SearchBase $Path -Properties Enabled | Where-Object {$_.Enabled -eq "True"})
         )

    begin{
        [array]$Table = $null
        $Counter = 0
         }
    
    process
    {
    $Date       = Get-Date -Format MM_dd_yyyy_HH_mm_ss
    $FolderName = "LocalAdminsReport("+ $Date + ")"
    New-Item -Path ".\$FolderName" -ItemType Directory -Force | Out-Null

        foreach($Computer in $ComputerName)
        {
            try
            {
                $PC      = Get-ADComputer $Computer
                $Name    = $PC.Name
                $CountPC = @($ComputerName).count
            }

            catch
            {
                Write-Host "Cannot retrieve computer $Computer" -ForegroundColor Yellow -BackgroundColor Red
                Add-Content -Path ".\$FolderName\ErrorLog.txt" "$Name"
                continue
            }

            finally
            {
                $Counter ++
            }

            Write-Progress -Activity "Connecting PC $Counter/$CountPC " -Status "Querying ($Name)" -PercentComplete (($Counter/$CountPC) * 100)

            try
            {
                $row = $null
                $members =[ADSI]"WinNT://$Name/Administrators"
                $members = @($members.psbase.Invoke("Members"))
                $members | foreach {
                            $User = $_.GetType().InvokeMember("Name", 'GetProperty', $null, $_, $null)
                                    $row += $User
                                    $row += " ; "
                                    }
                write-host "Computer ($Name) has been queried and exported." -ForegroundColor Green -BackgroundColor black 
                
                $obj = New-Object -TypeName PSObject -Property @{
                                "Name"           = $Name
                                "LocalAdmins"    = $Row
                                                    }
                $Table += $obj
            }

            catch
            {
            Write-Host "Error accessing ($Name)" -ForegroundColor Yellow -BackgroundColor Red
            Add-Content -Path ".\$FolderName\ErrorLog.txt" "$Name"
            }

            
        }
        try
        {
            $Table  | Sort Name | Select Name,LocalAdmins | Export-Csv -path ".\$FolderName\Report.csv" -Append -NoTypeInformation
        }
        catch
        {
            Write-Warning $_
        }
    }

    end{}
   }
    