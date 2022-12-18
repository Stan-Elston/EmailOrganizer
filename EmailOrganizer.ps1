# Create a COM object for Outlook and get the MAPI namespace
$Outlook = New-Object -ComObject Outlook.Application
$namespace = $Outlook.GetNameSpace("MAPI")

# Get the Inbox and Sent Items folders and create a folder for non-mail items if it doesn't already exist
$Inbox = $namespace.GetDefaultFolder([Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderInbox)
$SentItems = $namespace.GetDefaultFolder([Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderSentMail)


# Get all default folders in the Outlook mailbox
$folders = $namespace.Folders

while (($namespace.GetDefaultFolder([Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderInbox).Folders).count -gt 0)
{
	$Inbox_Folders = $namespace.GetDefaultFolder([Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderInbox).Folders
	foreach ($Folder in $Inbox_Folders)
	{
		while (($Folder.Items).count -gt 0)
		{
			foreach ($Item in $Folder.Items)
			{
				$Item.Move($Inbox) | Out-Null
			}
		}
		try
		{
			$Folder.Delete()
		}
		
		Catch
		{
			#Write code here
		}
		
	}
}


$nonMailFolder = $Inbox.Folders | Where-Object { $_.Name -eq "Non-Mail Items" }
if (-not $nonMailFolder)
{
	$nonMailFolder = $Inbox.Folders.Add("Non-Mail Items")
}

# Move all subfolders to the Inbox and Sent Items folders and move non-mail items to the non-mail folder
foreach ($folder in $folders)
{
	# Skip default folders
	if ($folder.DefaultItemType -eq [Microsoft.Office.Interop.Outlook.OlItemType]::olMailItem)
	{
		continue
	}
	
	while (($folder.Items).count -gt 0)
	{
		foreach ($Item in $folder.Items)
		{
			if ($Item.Class -eq [Microsoft.Office.Interop.Outlook.OlObjectClass]::olMail)
			{
				# Mark mail item as read
				$Item.UnRead = $false
				$Item.Move($folder) | Out-Null
			}
			else
			{
				$Item.Move($nonMailFolder) | Out-Null
			}
		}
	}
	if ($createdFolders -notcontains $folder)
	{
		$folder.Delete() | Out-Null
	}
}

# Create an empty dictionary to store the mail items by sender
$itemsBySender = @{ }

# Group the mail items in the Inbox and Sent Items folders by sender
foreach ($folder in @($Inbox, $SentItems))
{
	foreach ($Item in $folder.Items)
	{
		if ($Item.Class -eq [Microsoft.Office.Interop.Outlook.OlObjectClass]::olMail)
		{
			$sender = $Item.Sendername
			if (-not $itemsBySender)
			{
				$itemsBySender = @{ }
			}
			if (-not $itemsBySender.ContainsKey($sender))
			{
				$itemsBySender[$sender] = @()
			}
			$itemsBySender[$sender] += $Item
		}
	}
}


# Create a folder for each unique sender and move their mail items to the corresponding folder
foreach ($sender in $itemsBySender.Keys)
{
	# Check if a folder for the sender already exists
	$senderFolder = $Inbox.Folders | Where-Object { $_.Name -eq $sender }
	if (-not $senderFolder)
	{
		# Create a new folder for the sender
		$senderFolder = $Inbox.Folders.Add($sender)
	}
	# Move the mail items to the sender's folder
	foreach ($Item in $itemsBySender[$sender])
	{
		$Item.Move($senderFolder) | Out-Null
		$Item.UnRead = $false
	}
}

# Dispose of the Outlook application object
$Outlook.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null
Remove-Variable outlook

