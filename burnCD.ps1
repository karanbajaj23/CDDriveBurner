<# 
	.SYNOPSIS 
		Burns multiple CDs from folders present inside a specified parent folder path 

	.DESCRIPTION 

	.EXAMPLE 
		PS> .\burnCD.ps1 'C:\ParentFolder'

	.PARAMETER 
		Path The folder path containing the folders you'd like to burn to CDs 

#>
[CmdletBinding()]
param (
	[Parameter(Mandatory = $True,
			   ValueFromPipeline = $True,
			   ValueFromPipelineByPropertyName = $True,
			   HelpMessage = 'Enter Parent path')]
	[string]$ParentPath
)

begin
{
	
	Function Eject-CDTray
	{
        <#         
			.SYNOPSIS
				Ejects the local machine's CD tray.
		
			.DESCRIPTION
				This function looks for all available CD drives and ejects them.
		
			.EXAMPLE
				Eject-CDTray
		#>
		[CmdletBinding()]
		param (
		
		)
		
		begin
		{
			$sh = New-Object -ComObject "Shell.Application"
		}
		
		process
		{
			$sh.Namespace(17).Items() | Where { $_.Type -eq "CD Drive" } | foreach { $_.InvokeVerb("Eject") }
		}
		
		end
		{
			
		}
	}
	
	Function Close-CDTray
	{
        <#
			.SYNOPSIS
				Closes the local machine's CD tray if it's ejected.
			
			.EXAMPLE
				Close-CDTray
		#>
		[CmdletBinding()]
		param (
		
		)
		
		begin
		{
			$DiskMaster = New-Object -com IMAPI2.MsftDiscMaster2
			$DiscRecorder = New-Object -com IMAPI2.MsftDiscRecorder2
			$id = $DiskMaster.Item(0)
		}
		
		process
		{
			$DiscRecorder.InitializeDiscRecorder($id)
			$DiscRecorder.CloseTray()
		}
		
		end
		{
			
		}
	}
	
	Function Out-CD
	{
        <#
			.SYNOPSIS
				Burns the contents of a folder to a CD
			
			.DESCRIPTION
				This function retrieves the contents of a specified folder path and burns a CD with the specified title.
		
			.EXAMPLE
		
				PS> Out-CD -Path 'C:\Folder'
		
			.PARAMETER Path
				The folder path containing the files you'd like to burn to the CD.
		#>
		[CmdletBinding()]
		param (
			[Parameter(Mandatory = $True,
					   ValueFromPipeline = $True,
					   ValueFromPipelineByPropertyName = $True,
					   HelpMessage = 'What is the folder path you would like to burn?')]
			[string]$Path
		)
		
		begin
		{
			try
			{
				Write-Verbose 'Creating COM Objects.'
				
				$DiskMaster = New-Object -com IMAPI2.MsftDiscMaster2
				$DiscRecorder = New-Object -com IMAPI2.MsftDiscRecorder2
				$FileSystemImage = New-Object -com IMAPI2FS.MsftFileSystemImage
				$DiscFormatData = New-Object -com IMAPI2.MsftDiscFormat2Data
			}
			catch
			{
				$err = $Error[0]
				Write-Error $err
				return
			}
		}
		
		process
		{
			Write-Verbose 'Initializing Disc Recorder.'
			$id = $DiskMaster.Item(0)
			$DiscRecorder.InitializeDiscRecorder($id)
			
			Write-Verbose 'Assigning recorder.'
			$dir = $FileSystemImage.Root
			$DiscFormatData.Recorder = $DiscRecorder
			$DiscFormatData.ClientName = 'PowerShell Burner'

			Write-Verbose 'Empty medium.'
			$FileSystemImage.ChooseImageDefaults($DiscRecorder)
			$FileSystemImage.VolumeName = 'Volume'
			
			Write-Verbose "Adding directory tree ($Path)."
			$dir.AddTree($Path, $false)
			
			Write-Verbose 'Creating image.'
			$result = $FileSystemImage.CreateResultImage()
			$stream = $result.ImageStream
			
			Write-Verbose 'Burning.'
			$DiscFormatData.Write($stream)
			
			Write-Verbose 'Done.'
		}
		
		end
		{
			
		}
	}
	
	Eject-CDTray
}

process
{
    $files = Get-ChildItem $ParentPath

    for ($i=0; $i -lt $files.Count; $i++) 
    {
        $infile = $files[$i].FullName
        Read-Host 'Press Enter to burn '$infile
        Close-CDTray
		Out-Cd -Path $infile -Verbose
		Eject-CDTray
    }
}

end
{
	Close-CDTray
}
