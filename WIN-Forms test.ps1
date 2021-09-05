# Load Assemblies
Add-type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Functions
function Copy-Button {
    Copy-Item "$InputBox" -Destination "$Outputbox"
}

# Main GUI appearance 
$form = New-Object System.Windows.Forms.Form
$form.Text = 'GUI to transfer files from one machine to another'
$form.Size = New-Object System.Drawing.Size(600,400)
$form.AutoSize = $true
$form.StartPosition = 'CenterScreen'

# Create the Copy Button !! CHANGE LOCATION AND SIZE !!
$CopyButton = New-Object System.Windows.Forms.Button
$CopyButton.Location = New-Object System.Drawing.Size(20,90)
$CopyButton.Size = New-Object System.Drawing.Size(200,20)
$CopyButton.Text = "Copy Files"
$CopyButton.Add_Click({Copy-Button})

$form.controls.add($CopyButton)

$form.ShowDialog()

