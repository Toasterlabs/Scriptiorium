# Active Directory
## Export-GPOs
- Written by Andrew Ellis
- Sourced from <https://github.com/AndrewEllis93/PowerShell-Scripts>
- This exports all of your GPOs' HTML reports, a CSV detailing all the GPO links, and a txt list of all the GPOs to the specified output directory.

## Get-ADDomain
* PARAM: DomainName (Optional)
* PARAM: Credential (Optional)
* DESCR: Gathers Active Directory Domain object

## Get-ADForest
* PARAM: ForestName (Optional)
* PARAM: Credential (Optional)
* DESCR: Gathers Active Directory Forest object

## Get-ADObject
* PARAM: DomainController
* PARAM: SearchRoot (Optional)
* PARAM: SearchScrope (Optional)
* PARAM: PropertiesToLoad (Optional)
* PARAM: Credential (Optional)
* DESCR: Returns an object that represents an Active Directory object.

## get-DfsrGUID
* PARAM: ComputerName
* DESCR: Returns the DFSR volume GUID

## Get-dfsrLastUpdateDelta
* PARAM: ComputerName
* DESCR: Retrieves the delta between the last update time and now

## Get-dfsrLastUpdateTime
* PARAM: ComputerName
* DESCR: Retrieves the last time dfsr changed (or updated. However you want to see it...)
