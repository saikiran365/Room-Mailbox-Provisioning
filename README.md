# Room-Mailbox-Provisioning
# Room Mailbox Provisioning in Hybrid Exchange environment

# Input file
$Rooms = Import-Csv Rooms_OnPremise.csv
$ErrorFile = ".\Error_Log.csv"

Foreach($Room in $Rooms)
  {
  Try
    {
    $RoomId = $Room.RID
    $DG = $Room.RoomFinder
    $DG
    $ID = "Res-"+$RoomId
    $ID
    $Lastname = "r-"+$RoomId

##############################################################################update Domain  
 
    $UPN = "r-"+$RoomId+"@domain.com"
    $UPN

##############################################################################update Domain
    
    $RemoteRoutingAddr = "Res-"+$RoomId+"@domain.mail.onmicrosoft.com"
    $RemoteRoutingAddr
    $DisplayName = $Room.DisplayName
    $DisplayName
    $Office = $Room.Office
    $Office
    $Phone = $Room.Phone
    $Phone
    $City = $Room.City
    $City
    $Country = $Room.Country
    $Country
    $Floor = $Room.Floor
    $Floor

##############################################################update OnPremisesOrganizationalUnit ‘domain.com/Resources'
    
    New-RemoteMailbox -Alias "$ID" -Name "$ID" -FirstName "" -LastName "$LastName" -DisplayName "$ID" -UserPrincipalName "$UPN" -OnPremisesOrganizationalUnit ‘domain.com/Resources’ -RemoteRoutingAddress "$RemoteRoutingAddr" -Room

    Write-Host "Provisioned Room " $ID -ForegroundColor Green
    
    Start-Sleep -Seconds 2
    Set-RemoteMailBox $ID -DisplayName $DisplayName
    Write-Host "Display Name Set " $DisplayName -ForegroundColor Green

    Start-Sleep -Seconds 2
    If($DG)
        {
        Add-DistributionGroupMember -Identity "$DG" -Member "$ID"
        Write-Host "Added to Distribution Group " $DG -ForegroundColor Green
        }
    else
        {
        Write-Host "Distribution Group Not provided" -ForegroundColor Red
        }
    
    Set-User $ID -Office $Office -Phone $Phone -City $City -CountryOrRegion $Country -StreetAddress $Floor

    Write-Host "Provisioning Room Completed & Details updated" $DisplayName -ForegroundColor Green
    }
    
 Catch
    {
    Write-Host $_.Exception.Message -ForegroundColor Red
    $Error = $_.Exception.Message.ToString()
    Write-Host "Room Not Provisioned "$DisplayName -ForegroundColor DarkRed
    $DisplayName+"`t"+$Error | Out-File $Error_File -Append
    Continue
    }

}
