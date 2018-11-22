function Get-MailboxReport{
	Param(
		[Parameter(Mandatory = $true)]
		$userPrincipalName
	)

	# Validation
	Write-Verbose "Validating Email address"
	If($EmailAddress -notmatch "^([\w-\.]+)@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.)|(([\w-]+\.)+))([a-zA-Z]{2,4}|[0-9]{1,3})(\]?)$"){
		$Validation = "Failed"
		Write-Warning -Message "Invalid format: $userPrincipalName. Expected format: e.g. John@Contoso.com"
		return
	}

	# Collecting Mailbox details
	Write-verbose 'Collecting mailbox details' 
	$mb = Get-Mailbox -identity $userPrincipalName
	
	# Statistics
	$stats = $mb | Get-MailboxStatistics | Select-Object TotalItemSize,TotalDeletedItemSize,ItemCount,LastLogonTime,LastLoggedOnUserAccount

	# Collecting archive database information
	Write-Verbose 'Collecting Archive information'
	if ($mb.ArchiveDatabase){
        $archivestats = $mb | Get-MailboxStatistics -Archive | Select-Object TotalItemSize,TotalDeletedItemSize,ItemCount
    }else{
        $archivestats = "n/a"
    }

	# Collecting inbox statistics
	Write-Verbose 'Collecting inbox statistics'
	$inboxstats = Get-MailboxFolderStatistics $mb -FolderScope Inbox | Where {$_.FolderType -eq "Inbox"}

	# Collecting sent items statistics
	Write-Verbose 'Collecting sent items statistics'
	$sentitemsstats = Get-MailboxFolderStatistics $mb -FolderScope SentItems | Where {$_.FolderType -eq "SentItems"}

	# Collecting Deleted items statistics
	Write-Verbose 'Collecting deleted items statistics'
	$deleteditemsstats = Get-MailboxFolderStatistics $mb -FolderScope DeletedItems | Where {$_.FolderType -eq "DeletedItems"}

	# Collecting Last logon time
	Write-Verbose 'Collecting last logon time'
	$lastlogon = $stats.LastLogonTime

	# Collecting user object
	Write-Verbose 'Collecting user object'
	$user = Get-User $mb

	# Collecting primary database
	Write-Verbose 'Collecting Primary database for mailbox'
	$primarydb = $mailboxdatabases | where {$_.Name -eq $mb.Database.Name}

	# Collecting Archive databse
	Write-Verbose 'Collecting Archive database for mailbox'
	$archivedb = $mailboxdatabases | where {$_.Name -eq $mb.ArchiveDatabase.Name}

	# Object to return data in
	Write-Verbose 'Creating return object'
	$ObjUser = [PSCustomObject]
	$ObjUser = "" | select DisplayName,UserPrincipalName,PrimaryEmailAddress,OU,MailboxType,Title,Department,Office,MailboxTotalSize,MailboxSize,MailboxRecoverableItemsSize,
	MailboxInboxFolderSize, MailboxSentItemsSize,MailboxDeletedItemsSize,ArchiveTotalSize,ArchiveSize,ArchiveDeletedItemsSize,MailboxItems,ArchiveItems,
	AuditEnabled,EmailAddressPolicyUpdateEnabled,HiddenFromAddressLists,DatabaseQuotaDefaults, IssueWarningQuota,ProhibitSendQuota,ProhibitReceiveQuota,
	LastMailboxLogon,LastLogonBy,PrimaryMailboxDatabase,PrimaryServer,ArchiveMailboxDatabase,ArchiveServer

	# Populating
	Write-Verbose "Populating return object"
	$ObjUser.DisplayName = $mb.DisplayName
	$ObjUser.UserPrincipalName = $user.UserPrincipalName
	$ObjUser.PrimaryEmailAddress = $mb.PrimarySMTPAddress
	$ObjUser.OU = $user.OrganizationalUnit
	$ObjUser.MailboxType = $mb.RecipientTypeDetails
	$ObjUser.Title = $user.Title
	$ObjUser.Department = $user.Department
	$ObjUser.Office = $user.Office
	$ObjUser.MailboxTotalSize = ($stats.TotalItemSize.Value.ToMB() + $stats.TotalDeletedItemSize.Value.ToMB())
	$ObjUser.MailboxSize =$stats.TotalItemSize.Value.ToMB()
	$ObjUser.MailboxRecoverableItemsSize = $stats.TotalDeletedItemSize.Value.ToMB()
	$ObjUser.MailboxInboxFolderSize = $inboxstats.FolderandSubFolderSize.ToMB()
	$ObjUser.MailboxSentItemsSize = $sentitemsstats.FolderandSubFolderSize.ToMB()
	$ObjUser.MailboxDeletedItemsSize = $deleteditemsstats.FolderandSubFolderSize.ToMB()
	
	if ($archivestats -eq "n/a"){
		$ObjUser.ArchiveTotalSize = 'N/A'
		$ObjUser.ArchiveSize = 'N/A'
		$ObjUser.ArchiveDeletedItemsSize = 'N/A'
	}Else{
		$ObjUser.ArchiveTotalSize = ($archivestats.TotalItemSize.Value.ToMB() + $archivestats.TotalDeletedItemSize.Value.ToMB())
		$ObjUser.ArchiveSize = $archivestats.TotalItemSize.Value.ToMB()
		$ObjUser.ArchiveDeletedItemsSize = $archivestats.TotalDeletedItemSize.Value.ToMB()
	}
	
	$ObjUser.MailboxItems = $stats.ItemCount
	$ObjUser.ArchiveItems = $archivestats.ItemCount
	$ObjUser.AuditEnabled = $mb.AuditEnabled
	$ObjUser.EmailAddressPolicyUpdateEnabled = $mb.EmailAddressPolicyEnabled
	$ObjUser.HiddenFromAddressLists = $mb.HiddenFromAddressListsEnabled
	$ObjUser.DatabaseQuotaDefaults = $mb.UseDatabaseQuotaDefaults
	
	if ($mb.UseDatabaseQuotaDefaults -eq $true){
		$ObjUser.IssueWarningQuota = $primarydb.IssueWarningQuota
		$ObjUser.ProhibitSendQuota = $primarydb.ProhibitSendQuota
		$ObjUser.ProhibitReceiveQuota = $primarydb.ProhibitSendReceiveQuota
	}ElseIf($mb.UseDatabaseQuotaDefaults -eq $false){
		$ObjUser.IssueWarningQuota = $mb.IssueWarningQuota
		$ObjUser.ProhibitSendQuota = $mb.ProhibitSendQuota
		$ObjUser.ProhibitReceiveQuota = $mb.ProhibitSendReceiveQuota
	}

	$ObjUser.LastMailboxLogon = $lastlogon
	$ObjUser.LastLogonBy = $stats.LastLoggedOnUserAccount
	$ObjUser.PrimaryMailboxDatabase = $mb.Database
	$ObjUser.PrimaryServer = $primarydb.MasterServerOrAvailabilityGroup
	$ObjUser.ArchiveMailboxDatabase = $mb.ArchiveDatabase
	$ObjUser.ArchiveServer = $archivedb.MasterServerOrAvailabilityGroup

	# Returning
	return $ObjUser
}
