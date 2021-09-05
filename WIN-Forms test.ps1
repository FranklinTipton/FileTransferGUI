Add-type -assembly System.Windows.Forms

# Creating Variables 
$form = New-Object System.Windows.Forms.Form


# Main GUI appearance
$form.Text = 'GUI to transfer files from one machine to another'
$form.Width = 500
$form.Height = 350
$form.AutoSize = $true



$form.ShowDialog()

