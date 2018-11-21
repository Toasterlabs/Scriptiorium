#Requires -Version 3
Function Get-PwnedBreach
{
    <#
            .SYNOPSIS
            Report breached sites via the https://haveibeenpwned.com API service.
 
            .DESCRIPTION
            Report breached sites via the https://haveibeenpwned.com API service.
            This function queries the https://haveibeenpwned.com API service created by Troy Hunt (@troyhunt)
            and reports breached sites.  Returned with the breached site are the details surrounding the 
            breach and source of the original dump.
            .EXAMPLE
            Get-PwnedBreach
            Retuns all breached sites.
            .INPUTS
            None
 
            .NOTES
            Author:  Mark Ukotic
            Website: http://blog.ukotic.net
            Twitter: @originaluko
            GitHub:  https://github.com/originaluko/
            .LINK
            https://github.com/originaluko/haveibeenpwned
    #>
    
    Begin
    {
        $URI = 'https://haveibeenpwned.com/api/v2/breaches'
    }
    Process
    {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        try
        {
            $Request = Invoke-RestMethod -Uri $URI
        }
         Catch [System.Net.WebException] {
            Switch ($_.Exception.Message) {
                'The remote server returned an error: (400) Bad Request.' {
                    Write-Error -Message 'Bad Request.'
                }
                'The remote server returned an error: (403) Forbidden.' {
                    Write-Error -Message 'Forbidden - no user agent has been specified in the request.'
                }
                'The remote server returned an error: (404) Not Found.' {
                    Write-Output  'Not Found - No breach results found.'
                }
                'The remote server returned an error: (429) Too Many Requests.' {
                    Write-Error -Message 'Too many requests - the rate limit has been exceeded.'
                }
            }
            Break
        }
        $Request 
    }
}