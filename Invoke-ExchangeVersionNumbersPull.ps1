      $WebResponse = Invoke-WebRequest "https://technet.microsoft.com/en-us/library/hh135098(v=exchg.150).aspx"
    $data = $WebResponse.AllElements | where{$_.tagname -eq "table"}
    $results = @()

    foreach($i in $data){
        $Sections = (($i.innerHTML -split "<TR>"))
        
        foreach($section in $sections){          
            $ReleaseDate = ((($section -split "<P>") -split "</p>") | Select-String -Pattern "(January|February|March|April|May|June|July|August|September|October|November|December)\s\d{1,2},\s\d{4}|(January|February|March|April|May|June|July|August|September|October|November|December),\s\d{4}|(January|February|March|April|May|June|July|August|September|October|November|December)\s\d{4}|(January|February|March|April|May|June|July|August|September|October|November|December)\d{1,2},\s\d{4}") -replace "`n",""        
            $productName = ((($section -split "<P>") -split "</p>") | Select-string -Pattern "(Exchange)") -replace "`n",""  -replace '<a[^>]+href=\"(.*?)\"[^>]*>',""  -replace '</A>',''
            $VersionNumber = ((($section -split "<P>") -split "</p>") | select-string -Pattern "\d{1,2}\.\d{1,2}\.\d") -replace "`n",""

            if(!([STRING]::IsNullOrEmpty($productName))){
                $results += New-Object -Type PSObject -Property (
                    @{
                        "productName" = ($ReleaseDate)
                        "ReleaseDate" = ($productName)
                        "VersionNumber" = ($VersionNumber)
                    }
                )
            }
        }
    } 

    return $results
