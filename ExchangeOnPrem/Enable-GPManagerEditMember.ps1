Function Write-log{
    param(
            [Parameter( Mandatory=$False)]
            [String]$Message,
            [String]$LogFile

    )

        $Time = (Get-date).ToString()
    $Message = $Time + "  $Message"
    Try{
    Write-Host $Message
    $Message | Out-File -FilePath $Logfile -Force -Append -ErrorAction stop
    #Add-Content -Value $Msg -Path $Logfile -ErrorAction stop -Force
    }
    Catch{
    $ErrLine = "ERROR: Cannot append Log Message $Message"
    [Console]::Error.WriteLine($ErrLine)
    "Exception:{0}" -f $_.Exception.Message
    Exit 2
    }


}


$Logs = New-Item -Type file -Path D:\Directory -Name GroupEditLog.txt -Force

$Groups  = Get-DistributionGroup  -Filter {ManagedBy -ne $null} -ResultSize unlimited

#$Groups  = Get-DistributionGroup UCCXAGENTS-NNITDK -Filter {ManagedBy -ne $null} -ResultSize unlimited  #Test cmdlet

Foreach($group in $Groups){
    Try{
    $GroupOwner = $Group.ManagedBy[0]

    Add-ADPermission -Identity $group.name -User $GroupOwner -AccessRights Writeproperty -Properties Member -ErrorAction stop
    Write-log -Message "$group Member property edit success" -LogFile $Logs.FullName    
    }
    catch{
    $Errline = "Unalbe to edit the $group member property"
    [Console]::Error.WriteLine($ErrLine)
    "Exception:{0}" -f $_.Exception.Message

    Write-log -Message $Errline -LogFile $Logs.FullName
    Exit 2
    }


}
