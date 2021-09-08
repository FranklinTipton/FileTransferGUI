#Form class
$Forms = 'system.Windows.Forms.Form'
#Button class
$Button = 'system.Windows.Forms.Button'
#Getting the font details for applying the fields
$FontStyle = 'Microsoft Sans Serif,10,style=Bold'
$SysDrawPoint = 'System.Drawing.Point'
# setting InitialDirectory which open as intially during the click of browse button
$InitialDirectory = "C:\New folder"
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()
 
#The New-Object cmdlet creates an instance of a .NET Framework or COM object.
#Specifies the fully qualified name of the .NET Framework class. You cannot specify both the TypeName parameter and the ComObject parameter.
 
$TrackerFileUpload = New-Object -TypeName $Forms
$TrackerFileUpload.ClientSize = '400,114'
$TrackerFileUpload.text = 'dotnet-helpers.com - Please upload Daily Tracker'
$TrackerFileUpload.TopMost = $false
 
$txtFileName = New-Object -TypeName system.Windows.Forms.TextBox
$txtFileName.width = 280
$txtFileName.height = 20
$txtFileName.location = New-Object -TypeName $SysDrawPoint -ArgumentList (107,39)
$txtFileName.Font = $FontStyle
 
$btnFileBrowser = New-Object -TypeName $Button
$btnFileBrowser.BackColor = '#1a80b6'
$btnFileBrowser.text = 'Browse'
$btnFileBrowser.width = 96
$btnFileBrowser.height = 30
$btnFileBrowser.location = New-Object -TypeName $SysDrawPoint -ArgumentList (200,78)
$btnFileBrowser.Font = $FontStyle
$btnFileBrowser.ForeColor = '#ffffff'
 
$btnUpload = New-Object -TypeName $Button
$btnUpload.BackColor = '#00FF7F'
$btnUpload.text = 'Upload'
$btnUpload.width = 96
$btnUpload.height = 30
$btnUpload.location = New-Object -TypeName $SysDrawPoint -ArgumentList (291,78)
$btnUpload.Font = $FontStyle
$btnUpload.ForeColor = '#000000'
 
$lblFileName = New-Object -TypeName system.Windows.Forms.Label
$lblFileName.text = 'Choose File'
$lblFileName.AutoSize = $true
$lblFileName.width = 25
$lblFileName.height = 10
$lblFileName.location = New-Object -TypeName $SysDrawPoint -ArgumentList (20,40)
$lblFileName.Font = $FontStyle
 
#Adding the textbox,buttons to the forms for displaying
$TrackerFileUpload.controls.AddRange(@($txtFileName, $btnFileBrowser, $lblFileName, $btnUpload))
 
#Browse button click event
$btnFileBrowser.Add_Click({
Add-Type -AssemblyName System.windows.forms | Out-Null
#Creating an object for OpenFileDialog to a Form
$OpenDialog = New-Object -TypeName System.Windows.Forms.OpenFileDialog
#Initiat browse path can be set by using initialDirectory
$OpenDialog.initialDirectory = $initialDirectory
#Set filter only to upload Excel file
$OpenDialog.filter = '.xlsx (*(.xlsx)) | *.xlsx | *.txt'
$OpenDialog.ShowDialog() | Out-Null
$filePath = $OpenDialog.filename
#Assigining the file choosen path to the text box
$txtFileName.Text = $filePath 
$TrackerFileUpload.Refresh()
})

#Upload button click eventy
$btnUpload.Add_Click({
#Set the destination path
$destPath = 'C:\EMS\Destination'
#$txtFileName.Text = ""
Copy-Item -Path $txtFileName.Text -Destination $destPath
})
 
$null = $TrackerFileUpload.ShowDialog()
$txtFileName.Text = ""