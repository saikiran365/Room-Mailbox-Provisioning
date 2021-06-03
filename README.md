#############################################################
# Room Mailbox Provisioning in Hybrid Exchange environment

# save the input csv in same location with columns "RID,DisplayName,RoomFinder,Office,Phone,City,Country,Floor"
# RID === ID or Alias of the Room
# DisplayName == Display Name of the Room
# RoomFinder = If you have room finder group and want to add to that group give under this parameter
# Office = Office location
# Phone = Phone number of the Room
# City = City
# Country = Country
# Floor = StreetAddress or Floor & Building
# Copy below code & update domain details with yours then save as .Ps1 

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
        ############################################################ **Update Domain below**  
 
        $UPN = "r-"+$RoomId+"@domain.com"

        ############################################################**Write Host UPN of the Room Getting created**
        $UPN
   
        ############################################################**update Domain below**
    
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

        #############################################################**update OnPremisesOrganizationalUnit ‘domain.com/Resources'**
    
        New-RemoteMailbox -Alias "$ID" -Name "$ID" -FirstName "" -LastName "$LastName" -DisplayName "$ID" -UserPrincipalName "$UPN" -OnPremisesOrganizationalUnit ‘domain.com/Resources’ -RemoteRoutingAddress "$RemoteRoutingAddr" -Room

        Write-Host "Provisioned Room " $ID -ForegroundColor Green

        Start-Sleep -Seconds 2
        Set-RemoteMailBox $ID -DisplayName $DisplayName
        Write-Host "Display Name Set " $DisplayName -ForegroundColor Green

        Start-Sleep -Seconds 2
        #############################################################Adding room to room finder
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
    #######End of Try
    Catch
        {
        Write-Host $_.Exception.Message -ForegroundColor Red
        $Error = $_.Exception.Message.ToString()
        Write-Host "Room Not Provisioned "$DisplayName -ForegroundColor DarkRed
        $DisplayName+"`t"+$Error | Out-File $Error_File -Append
        Continue
        }
      }
