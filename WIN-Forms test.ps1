# Load Assemblies
Add-type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

## Variables

$InitialDirectory = "C:\"

### Define Controls ###

# Main GUI appearance 
$form = New-Object System.Windows.Forms.Form
$form.Text = 'GUI to transfer files from one machine to another'
$form.Size = New-Object System.Drawing.Size(600,400)
$form.AutoSize = $true
$form.StartPosition = 'CenterScreen'

# Create Original Location 
$OriginalLocation = New-Object System.Windows.Forms.ListBox
$OriginalLocation.Location = New-Object System.Drawing.Size(20,50)
$OriginalLocation.Size = New-Object System.Drawing.Size(240,200)
$OriginalLocation.Anchor = ([System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Top)
$OriginalLocation.IntegralHeight = $False
$OriginalLocation.AllowDrop = $True


## What is this?
$OriginalLocation.controls.AddRange(@($OriginalLocation))

# Label for Original location
$LabelOL = New-Object System.Windows.Forms.Label
$LabelOL.text = "Files to be copied (Drag and Drop):"
$LabelOL.Location = New-Object System.Drawing.Size(20,30)
$LabelOL.BackColor = "Transparent"
$LabelOL.AutoSize = $true

# Create New Location --Move to the right of GUI
$NewLocation = New-Object System.Windows.Forms.ComboBox
$NewLocation.Location = New-Object System.Drawing.Size(300,50)
$NewLocation.Size = New-Object System.Drawing.Size(200,20)

# Label for New location
$LabelNL = New-Object System.Windows.Forms.Label
$LabelNL.text = "Select New Location Path Here:"
$LabelNL.Location = New-Object System.Drawing.Size(300,30)
$LabelNL.BackColor = "Transparent"
$LabelNL.AutoSize = $true

# Create the Copy Button 
$CopyButton = New-Object System.Windows.Forms.Button
$CopyButton.Location = New-Object System.Drawing.Size(20,335)
$CopyButton.Size = New-Object System.Drawing.Size(200,20)
$CopyButton.Text = "Copy Files"
$CopyButton.Add_Click({CopyButton})

# Create the Browse Button !! Add File Browser on add_click
$BrowseButton = New-Object System.Windows.Forms.Button
$BrowseButton.Location = New-Object System.Drawing.Size(20,10)
$BrowseButton.Size = New-Object System.Drawing.Size(80,20)
$BrowseButton.Text = "Browse Files"
$BrowseButton.Add_Click({
	Add-Type -AssemblyName System.windows.forms | Out-Null
	$OpenDialog = New-Object -TypeName System.Windows.Forms.OpenFileDialog
	#Initiate browse path can be set by using initialDirectory
	$OpenDialog.initialDirectory = $initialDirectory
	$OpenDialog.ShowDialog() | Out-Null
	$filePath = $OpenDialog.filename
	$filePath.ListBox.value = $OriginalLocation
	#Assigining the file choosen path to the text box
	$OriginalLocation.List = $filePath 
	Write-Host $filePath
	$OriginalLocation.Refresh()
})

### Add forms ###
$form.SuspendLayout()
$form.controls.Add($OriginalLocation)
$form.Controls.Add($LabelOL)
$form.controls.Add($NewLocation)
$form.Controls.Add($LabelNL)
$form.controls.add($CopyButton)
$form.controls.add($BrowseButton)
$form.ResumeLayout()

### Functions ###

$Locations = @("Server 1","Server 2","Server 3")

foreach($Location in $Locations){

$NewLocation.Items.Add($Location)

}
function CopyButton {

	$selection = $NewLocation.SelectedItem.Equals()

	$input = $OriginalLocation.text

	if($selection -eq "Server 1"){

		Copy-Item "$input" -Destination "C:\Test" -Recurse

	}
	elseif($selection -eq "Server 2"){

		
	}
	elseif($selection -eq "Server 3"){

		
	}
}

$OriginalLocation_DragOver = [System.Windows.Forms.DragEventHandler]{
	if ($_.Data.GetDataPresent([Windows.Forms.DataFormats]::FileDrop)) # $_ = [System.Windows.Forms.DragEventArgs]
	{
	    $_.Effect = 'Copy'
	}
	else
	{
	    $_.Effect = 'None'
	}
}

$OriginalLocation_DragDrop = [System.Windows.Forms.DragEventHandler]{
	foreach ($filename in $_.Data.GetData([Windows.Forms.DataFormats]::FileDrop)) # $_ = [System.Windows.Forms.DragEventArgs]
    {
		$OriginalLocation.Items.Add($filename)
	}
    $statusBar.Text = ("List contains $($OriginalLocation.Items.Count) items")
}

$form_FormClosed = {
	try
    {
        $OriginalLocation.remove_Click($button_Click)
		$OriginalLocation.remove_DragOver($OriginalLocation_DragOver)
		$OriginalLocation.remove_DragDrop($OriginalLocation_DragDrop)
        $OriginalLocation.remove_DragDrop($OriginalLocation_DragDrop)
		$OriginalLocation.remove_FormClosed($Form_Cleanup_FormClosed)
	}
	catch [Exception]
    { }
}

$OriginalLocation.Add_DragOver($OriginalLocation_DragOver)
$OriginalLocation.Add_DragDrop($OriginalLocation_DragDrop)
$form.Add_FormClosed($form_FormClosed)

$form.ShowDialog()
