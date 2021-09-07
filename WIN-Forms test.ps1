# Load Assemblies
Add-type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

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

# Label for Original location
$LabelOL = New-Object System.Windows.Forms.Label
$LabelOL.text = "Files to be copied (Drag and Drop):"
$LabelOL.Location = New-Object System.Drawing.Size(20,30)
$LabelOL.BackColor = "Transparent"
$LabelOL.AutoSize = $true

# Create New Location --Move to the right of GUI
$NewLocation = New-Object System.Windows.Forms.TextBox
$NewLocation.Location = New-Object System.Drawing.Size(300,50)
$NewLocation.Size = New-Object System.Drawing.Size(200,20)

# Label for New location
$LabelNL = New-Object System.Windows.Forms.Label
$LabelNL.text = "Select New Location Path Here:"
$LabelNL.Location = New-Object System.Drawing.Size(300,30)
$LabelNL.BackColor = "Transparent"
$LabelNL.AutoSize = $true

# Create the Copy Button !! CHANGE LOCATION AND SIZE !!
$CopyButton = New-Object System.Windows.Forms.Button
$CopyButton.Location = New-Object System.Drawing.Size(20,335)
$CopyButton.Size = New-Object System.Drawing.Size(200,20)
$CopyButton.Text = "Copy Files"
$CopyButton.Add_Click({Copy-Button})

### Add forms ###
$form.SuspendLayout()
$form.controls.Add($OriginalLocation)
$form.Controls.Add($LabelOL)
$form.controls.Add($NewLocation)
$form.Controls.Add($LabelNL)
$form.controls.add($CopyButton)
$form.ResumeLayout()

### Functions ###

function Copy-Button {
    Copy-Item "$OriginalLocation" -Destination "$NewLocation"
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
