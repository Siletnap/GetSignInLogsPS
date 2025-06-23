$TenantId           = 'TenantId'
$ClientId           = 'ClientId'
$CertThumbprint     = 'CertThumbprint'
$csvPath = "C:\csvPath.csv"

# Install modules if not installed
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph)) {
    Install-Module Microsoft.Graph -Scope CurrentUser -Force
}

function Format-Logs{
	param(
		$currentBatch
	)
		# Format dateTime and add properties to each object 
		foreach ($row in $currentBatch) {
			$local = $row.createdDateTime.ToLocalTime()
			$localF = $local.ToString('yyyy/MM/dd HH:mm:ss')
			$row | Add-Member -MemberType NoteProperty -Name CreatedDateTimeLocalF -Value ($localF)
		  
			$browser = $row.deviceDetail.browser
			$deviceID = $row.deviceDetail.deviceID
			$displayName = $row.deviceDetail.displayName
			$os = $row.deviceDetail.operatingSystem
			$locationCR = "$($row.location.city), $($row.location.state), $($row.location.countryOrRegion)"
			$addStatsDetails = $row.status.additionalDetails
			$errCode = $row.status.errorCode
			$failureReason = $row.status.failureReason
		  
			$row | Add-Member -MemberType NoteProperty -Name Browser        	 		-Value $browser
			$row | Add-Member -MemberType NoteProperty -Name DeviceID 					-Value $deviceID
			$row | Add-Member -MemberType NoteProperty -Name DisplayName            	-Value $displayName
			$row | Add-Member -MemberType NoteProperty -Name OperatingSystem       		-Value $os
			$row | Add-Member -MemberType NoteProperty -Name LocationL		       		-Value $locationCR
			$row | Add-Member -MemberType NoteProperty -Name AdditionalStatusDetails    -Value $addStatsDetails
			$row | Add-Member -MemberType NoteProperty -Name StatusErrorCode       		-Value $errCode
			$row | Add-Member -MemberType NoteProperty -Name FailureReason       		-Value $failureReason
		}
	
		
		# Format object properties/ordering
		return $currentBatch |
					Select-Object `
						Id,
						CorrelationId,
						UserId,
						CreatedDateTimeLocalF,
						UserDisplayName,
						UserPrincipalName,
						IPAddress,
						LocationL,
						ConditionalAccessStatus,
						Browser,
						DeviceID,
						DisplayName,
						OperatingSystem,
						StatusErrorCode,
						FailureReason,
						AdditionalStatusDetails,
						AppDisplayName,
						AppId,
						ClientAppUsed,
						CreatedDateTime,
						DeviceDetail,
						IsInteractive,
						ResourceDisplayName,
						ResourceId,
						RiskDetail,
						RiskEventTypes,
						RiskEventTypesV2,
						RiskLevelAggregated,
						RiskLevelDuringSignIn,
						RiskState,
						Status,
						AppliedConditionalAccessPolicies,
						Location,
						AdditionalProperties
}

# Error check connection to Graph
Try {
    Import-Module Microsoft.Graph.Authentication, Microsoft.Graph.Reports -ErrorAction Stop
    Connect-MgGraph -NoWelcome -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertThumbprint -ErrorAction Stop
	
	# If csv !exists, create new one
	if (-not (Test-Path $csvPath)) {
		Write-Verbose "No existing CSV found; creating new file and seeding with all current sign-ins."
		
		# Grab everything you want on day one (e.g. last 30, or -All) -Top 30 #or# -All & sort old-->new
		$initialBatch = Get-MgAuditLogSignIn -Top 300 |
						Sort-Object createdDateTime
		
		Format-Logs -currentBatch $initialBatch |
		  # Write to csv (initialisation)
		  Export-Csv -Path $csvPath -NoTypeInformation
		
		Write-Host "CSV bootstrapped with $($initialBatch.Count) rows."
		return
	} else {
		
		# If csv exists v
		
		# Loads entire csv vvv may impact performance with larger files
		# consider $headerLine = Get-Content -Path $csvPath -First 1 ;;; $columnNames = $headerLine -split ',' ;;; $lastLine = Get-Content -Path $csvPath -Tail 1 ;;; $lastObject = $lastLine | ConvertFrom-Csv -Header $columnNames
		$existingLogs = Import-Csv -Path $csvPath

		# Grab latest csv row for its unique id
		$lastObject  = $existingLogs | Select-Object -Last 1
		# Write-Output $lastObject.id
		
		# Grab latest logs (300? ~9000 per month) and sort old-->new
		$newLogs = Get-MgAuditLogSignIn -Top 300  |
				Sort-Object createdDateTime
				
		# Remove logs matching id of latest from csv or older
		while ($newLogs.Count -gt 1 -and $newLogs[0].Id -ne $lastObject.Id) {
			# Drop the first element by reassigning a sub-array
			$newLogs = $newLogs[1..($newLogs.Count - 1)]
		}

		# Finally, drop the matching element itself
		if ($newLogs.Count -gt 1 -and $newLogs[0].Id -eq $lastObject.Id) {
			$newLogs = $newLogs[1..($newLogs.Count - 1)]
		}elseif($newLogs.Count -gt 0 -and $newLogs[0].Id -eq $lastObject.Id) {
			$newLogs = 0
		}
		
		if ($newLogs) {
			# Format dateTime and add properties to each object
			Format-Logs -currentBatch $newLogs |
				# Write to csv (append existing)
				Export-Csv -Path $csvPath -NoTypeInformation -Append

			Write-Host "Appended $($newLogs.Count) new records."
		}else {
			Write-Host "No new sign-ins since $($lastTime)"
		}
	}
}
Finally {
    Disconnect-MgGraph
}
