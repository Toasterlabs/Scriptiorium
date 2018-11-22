function Test-DomainControllerHealth{

	# Variables (sort of)
	$TimeOut = "60"
	$getForest = [system.directoryservices.activedirectory.Forest]::GetCurrentForest()
	$DCServers = $getForest.domains | ForEach-Object {$_.DomainControllers} | ForEach-Object {$_.Name}
	
	$healthReport = @()

	# Process each domain controller
	Foreach($DC in $DCServers){
		
		$objTemp = [PSCustomObject]
		$objTemp = "" | select Servername, Connectivity, NetlogonStatus, NTDSStatus, DNSServiceStatus, NetlogonTest, ReplicationTest, ServicesTest,
		AdvertisingTest, FSMOTest
		
		# Setting name
		$objTemp.Servername = $DC

		# First we test network connectivity
		if ( Test-Connection -ComputerName $DC -Count 1 -ErrorAction SilentlyContinue ) {
			invoke-WriteLog -Message "SUCCESS: $DC Ping Test" -Runlog $runlog
			
			$objTemp.Connectivity = 'True'

			# Setting short identity
			$ShortIdentity = $Identity.Replace(('.'+$getForest.Name),'')                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                     
		        
			#region netlogon
			# Get the service status
			$serviceStatus = start-job -scriptblock {get-service -ComputerName $($args[0]) -Name "Netlogon" -ErrorAction SilentlyContinue} -ArgumentList $DC
            
			# Wait for completion
			wait-job $serviceStatus -timeout $timeout
            
			# If we're still waiting after 60 seconds the service timed out
			if($serviceStatus.state -like "Running"){
                 Invoke-writelog -Message "ALERT: $DC Netlogon Service TimeOut" -Runlog $runlog
                 stop-job $serviceStatus
				 $objTemp.NetlogonStatus = "TimeOut"
            }Else{
                $NetlogonStatus = Receive-job $serviceStatus
                
				if ($NetlogonStatus.status -eq "Running") {
 					$objTemp.NetLogonStatus = $NetlogonStatus.Status
                }Else{ 
       			    $objTemp.NetLogonStatus = $NetlogonStatus.Status
                } 
            }
			#endregion

			#region NTDS Service Status
            $serviceStatus = start-job -scriptblock {get-service -ComputerName $($args[0]) -Name "NTDS" -ErrorAction SilentlyContinue} -ArgumentList $DC
			wait-job $serviceStatus -timeout $timeout

			# If we're still waiting after 60 seconds the service timed out
			if($serviceStatus.state -like "Running"){
                 stop-job $serviceStatus
				 $objTemp.NTDSStatus = "TimeOut"
            }Else{
                $NetlogonStatus = Receive-job $serviceStatus
                
				if ($NetlogonStatus.status -eq "Running") {
 					$objTemp.NTDSStatus = $NetlogonStatus.Status
                }Else{ 
       			    $objTemp.NTDSStatus = $NetlogonStatus.Status
                } 
            }
			#endregion

			#region DNS Service Status
			$serviceStatus = start-job -scriptblock {get-service -ComputerName $($args[0]) -Name "DNS" -ErrorAction SilentlyContinue} -ArgumentList $DC
            wait-job $serviceStatus -timeout $timeout
            
			# If we're still waiting after 60 seconds the service timed out
			if($serviceStatus.state -like "Running"){
                 stop-job $serviceStatus
				 $objTemp.DNSServiceStatus = "TimeOut"
            }Else{
                $NetlogonStatus = Receive-job $serviceStatus
                
				if ($NetlogonStatus.status -eq "Running") {
 					$objTemp.DNSServiceStatus = $NetlogonStatus.Status
                }Else{ 
       			    $objTemp.DNSServiceStatus = $NetlogonStatus.Status
                } 
            }
			#endregion
			
			#region Netlogon Test
			add-type -AssemblyName microsoft.visualbasic 
            $cmp = "microsoft.visualbasic.strings" -as [type]
            $sysvol = start-job -scriptblock {dcdiag /test:netlogons /s:$($args[0])} -ArgumentList $DC
            wait-job $sysvol -timeout $timeout

			if($sysvol.state -like "Running"){
				stop-job $sysvol
				$objtemp.NetlogonTest = "TimeOut"
			}Else{
				$sysvol1 = Receive-job $sysvol
               if($cmp::instr($sysvol1, "passed test NetLogons")){
					$objtemp.NetlogonTest = "Passed"
			   }Else{
					$objtemp.NetlogonTest = "Failed"
			   }
			}
			#endregion

			#region Repl test
            add-type -AssemblyName microsoft.visualbasic 
            $cmp = "microsoft.visualbasic.strings" -as [type]

            $sysvol = start-job -scriptblock {dcdiag /test:Replications /s:$($args[0])} -ArgumentList $DC
            wait-job $sysvol -timeout $timeout

            if($sysvol.state -like "Running"){
				$objTemp.ReplicationTest = "TimeOut"
				stop-job $sysvol
               }else{
				$sysvol1 = Receive-job $sysvol
				if($cmp::instr($sysvol1, "passed test Replications")){
                  $objTemp.ReplicationTest = "Passed"
               }else{
                  $objTemp.ReplicationTest = "Failed"
               }
			}
			#endregion

			#region Services Test
			add-type -AssemblyName microsoft.visualbasic 
            $cmp = "microsoft.visualbasic.strings" -as [type]
            
			$sysvol = start-job -scriptblock {dcdiag /test:Services /s:$($args[0])} -ArgumentList $DC
            wait-job $sysvol -timeout $timeout
            
			if($sysvol.state -like "Running"){
				$objTemp.ServicesTest = "TimeOut"
				stop-job $sysvol
			}else{
				$sysvol1 = Receive-job $sysvol
				if($cmp::instr($sysvol1, "passed test Services")){
					$objTemp.ServicesTest = "Passed"
				}else{
					$objTemp.ServicesTest = "Failed"
                }
			}
			#endregion Services Test

			#region Advertising Test
			add-type -AssemblyName microsoft.visualbasic 
            $cmp = "microsoft.visualbasic.strings" -as [type]
            
			$sysvol = start-job -scriptblock {dcdiag /test:Advertising /s:$($args[0])} -ArgumentList $DC
            wait-job $sysvol -timeout $timeout
            
			if($sysvol.state -like "Running"){
				$objTemp.AdvertisingTest = "TimeOut"
				stop-job $sysvol
            }else{
				$sysvol1 = Receive-job $sysvol
				if($cmp::instr($sysvol1, "passed test Advertising")){
					$objTemp.AdvertisingTest = "Passed"
                }else{
					$objTemp.AdvertisingTest = "Failed"
                }
			}
			#endregion

			#region FSMO Test
            add-type -AssemblyName microsoft.visualbasic 
            $cmp = "microsoft.visualbasic.strings" -as [type]
            
			$sysvol = start-job -scriptblock {dcdiag /test:FSMOCheck /s:$($args[0])} -ArgumentList $DC
            wait-job $sysvol -timeout $timeout
            
			if($sysvol.state -like "Running"){
				$objTemp.FSMOTest = "TimeOut"
				stop-job $sysvol
			}else{
				$sysvol1 = Receive-job $sysvol
				if($cmp::instr($sysvol1, "passed test FsmoCheck")){
					$objTemp.FSMOTest = "Passed"
                }else{
					$objTemp.FSMOTest = "Failed"
                }
			}
            #endregion

			$healthReport += $objTemp
		}Else{
			$objTemp.Connectivity = 'False'

			$healthReport += $objTemp
		}
	}

	Return $healthReport
}
			