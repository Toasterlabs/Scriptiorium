    $WebResponse = Invoke-WebRequest "https://technet.microsoft.com/en-us/library/hh135098(v=exchg.150).aspx"
    $data = $WebResponse.AllElements | where{$_.tagname -eq "table"}
    $results = @()

    foreach($i in $data){
        $Sections = (($i.innerHTML -split "<TR>")) 
        $sections = ($sections -replace ' data-th="&#10;         Product name&#10;        ">',"" -replace ' data-th="&#10;         Release date&#10;        ">',"" -replace ' data-th="&#10;         Build number&#10;        ">',"" -replace '<TD','' -replace '<TBODY>','' -replace '</TBODY>','' -replace '</TD>','' -replace '</TR>','' -replace '<TR Responsive="true">','' -replace '<TH scope=col>Product name </TH>','' -replace '<TH scope=col>Product name </TH>','' -replace '<TH scope=col>Release date </TH>','' -replace '<TH scope=col>Build number </TH>','' -replace '</P>','' -replace '<a[^>]+href=\"(.*?)\"[^>]*>',""  -replace '</A>','' -replace 'data-th="&#10;           Release date&#10;          ">','') -replace 'data-th="&#10;             Release date&#10;            ">',''
        foreach($section in $sections){
        
            $cleanedUp = ($Section -split "<P>" ) -replace "`n",""

            $results += New-Object -Type PSObject -Property (
                @{
                    "productName" = ($cleanedUp[1])
                    "ReleaseDate" = ($cleanedUp[2])
                    "VersionNumber" = ($cleanedUp[3])
                }
            )
        }
    }

