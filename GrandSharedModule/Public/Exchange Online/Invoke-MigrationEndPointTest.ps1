Function Invoke-MigrationEndPointTest{

	Write-Verbose "Starting Migration Endpoint testing"

	# Query if there is a migration endpoint
	
	Try{
		Write-Verbose "Retrieving endpoints"
		$Endpoints = Get-MigrationEndPoint
	}Catch{
		Write-Warning "Failed to retrieve endpoints"
	}
	
    If(!($Endpoints)){
		Write-Warning "No Endpoints to test!"
	}Else{
		
	    # Test migration endpoint
	    $validEndpoints = @()
	
		foreach($endpoint in $endpoints){
			Write-Verbose "Testing $endpoint"
			$EndpointTest = Test-MigrationServerAvailability -Endpoint $endpoint.RemoteServer
				
			if($EndpointTest.Result -eq "Failed"){
				Write-Verbose "Test of $endpoint failed"
			}

			If($EndpointTest.Result -eq "Success"){
				Write-Verbose -Message "Endpoint $($Endpoint.RemoteServer) test result has yielded a success!"
				$validEndpoints += $endpoint
		    }
		}
	}

	return $validEndpoints
}