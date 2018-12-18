# Scripts
This folder contains a number of scripts which might, or might not, require the functions defined in this repository.

## Invoke-BuildModule
Use this script to compile a psm1 file with the required functions for your script. This is leaps and bounds faster than importing each function seperately!

## Test-DomainControllerHealth
Tests the health of every domain controller in the forest.

## Invoke-M365PSTImporter
Wrapper for the AZCopy tool to Upload PST files to Azure Blob (the free one for Exchange Online PST imports)
This one does not allow for automated import of CSV files

## Invoke-M365PSTImport
Wrapper for the AZCopy tool to Upload PST files to Azure Blob (It costs money, but allows for the import process of PST files to be automated.
TODO: Add code to allow importing PST files in to mailboxes
