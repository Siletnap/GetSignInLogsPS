$TenantId           = 'TenantId'
$ClientId           = 'ClientId'
$CertThumbprint     = 'CertThumbprint'
$csvPath = "C:\csvPath.csv"

$nestedProps = @(
  'location',
  'deviceDetail',
  'status',
  'appliedConditionalAccessPolicies'
  # add more as discovered in metadata…
)


if (-not (Get-Module -ListAvailable -Name Microsoft.Graph)) {
    Install-Module Microsoft.Graph -Scope CurrentUser -Force
}


Try {
    Import-Module Microsoft.Graph.Authentication, Microsoft.Graph.Reports -ErrorAction Stop
    Connect-MgGraph -NoWelcome -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertThumbprint -ErrorAction Stop

	if (-not (Test-Path $csvPath)) {
		Write-Verbose "No existing CSV found; creating new file and seeding with all current sign-ins."
		
		# 1a) grab everything you want on day one (e.g. last 30, or -All) -Top 30 #or# -All
		$initialBatch = Get-MgAuditLogSignIn -Top 100 |
						Sort-Object createdDateTime
		
		<#
		foreach ($rowL in $initialBatch){
			$rowLocalTime = $rowL.createdDateTime.ToLocalTime()
			$rowL.createdDateTime = $rowLocalTime.ToString('yyyy/MM/dd HH:mm:ss')
			Write-Output "rowlocaltostring $($rowLocalTime.ToString('yyyy/MM/dd HH:mm:ss'))"
			Write-Output $rowL.createdDateTime
		}
		#>

		<#
		$secondaryBatch = $initialBatch |
							  Select-Object *,
								@{Name='newCreatedDateTime';Expression={
								   $_.createdDateTime.ToLocalTime.ToString('yyyy/MM/dd HH:mm:ss')
								}}
		$secondaryBatch | Export-Csv -Path $csvPath -NoTypeInformation
		#>
		
		
		foreach ($row in $initialBatch) {
			$local = $row.createdDateTime.ToLocalTime()
			$localF = $local.ToString('yyyy/MM/dd HH:mm:ss')
			# Write-Output 'l '$local
			# Write-Output "^v"
			# Write-Output 'f '$localF
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
	
		
		# 1b) write it out, including the header row
		$initialBatch |
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
				AdditionalStatusDetails,
				FailureReason,
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
				AdditionalProperties |
		  Export-Csv -Path $csvPath -NoTypeInformation
		
		
		# 1c) you can exit here if you don’t want to merge again right now
		Write-Host "CSV bootstrapped with $($initialBatch.Count) rows."
		return
	} else {
		
		# Loads entire csv vvv may impact performance with larger files
		# consider $headerLine = Get-Content -Path $csvPath -First 1 ;;; $columnNames = $headerLine -split ',' ;;; $lastLine = Get-Content -Path $csvPath -Tail 1 ;;; $lastObject = $lastLine | ConvertFrom-Csv -Header $columnNames
		$existingLogs = Import-Csv -Path $csvPath
		#Write-Output $existingLogs
		# Grab only sign-ins newer than our last record
		$lastObject  = $existingLogs | Select-Object -Last 1
		Write-Output $lastObject.id
		#$lastTimeString  = $lastObject.createdDateTime
		#Write-Output $lastObject.id
		#$lastTime = [datetime]$lastTimeString
		
		<#
		Write-Output "lasttimestring"
		Write-Output $lastTimeString
		$lastTime = [datetime]$lastTimeString
		Write-Output "lasttime"
		Write-Output $lastTime
		#>
		
		<#
		# $utcFilter = "createdDateTime gt $($([datetime]$lastTime).ToUniversalTime().ToString('u'))"
		$newLogs = Get-MgAuditLogSignIn -Top 1
		foreach ($log in $newLogs){
			Write-Output $log.createdDateTime.ToUniversalTime().ToString('o')
			$localDateTime = $log.createdDateTime.ToLocalTime()
			Write-Output $localDateTime
			$formatted = $localDateTime.ToString('yyyy/MM/dd HH:mm:ss')
			Write-Output $formatted
			Write-Output $log.ID
		}
		#>
		
		# $utcFilter = "createdDateTime gt $($lastTime.ToUniversalTime().ToString('o'))"
		# $newFilteredLogs = Get-MgAuditLogSignIn -Filter $utcFilter -Top 30  |
		#		Sort-Object createdDateTime
				
		$newFilteredLogs = Get-MgAuditLogSignIn -Top 30  |
				Sort-Object createdDateTime
				
		Write-Output $newFilteredLogs.Count
				
		while ($newFilteredLogs.Count -gt 1 -and $newFilteredLogs[0].Id -ne $lastObject.Id) {
			# Drop the first element by reassigning a sub-array
			$newFilteredLogs = $newFilteredLogs[1..($newFilteredLogs.Count - 1)]
			Write-Output $newFilteredLogs.Count
		}

		# Finally, drop the matching element itself
		if ($newFilteredLogs.Count -gt 1 -and $newFilteredLogs[0].Id -eq $lastObject.Id) {
			$newFilteredLogs = $newFilteredLogs[1..($newFilteredLogs.Count - 1)]
		}elseif($newFilteredLogs.Count -gt 0 -and $newFilteredLogs[0].Id -eq $lastObject.Id) {
			$newFilteredLogs = 0
		}
		
		Write-Output $newFilteredLogs.Count
		
		
		if ($newFilteredLogs) {
			
			foreach ($row in $newFilteredLogs) {
				$local = $row.createdDateTime.ToLocalTime()
				$localF = $local.ToString('yyyy/MM/dd HH:mm:ss')
				$row | Add-Member -MemberType NoteProperty -Name CreatedDateTimeLocalF -Value ($localF)
				
				Write-Output $row.CreatedDateTimeLocalF
				
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
			
			$newFilteredLogs |
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
					AdditionalStatusDetails,
					FailureReason,
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
					AdditionalProperties |
			  Export-Csv -Path $csvPath -NoTypeInformation -Append

			Write-Host "Appended $($newFilteredLogs.Count) new records."
		}else {
			Write-Host "No new sign-ins since $($lastTime)"
		}
		##>
	
	}
}
Finally {
    Disconnect-MgGraph
}
