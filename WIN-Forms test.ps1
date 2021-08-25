Add-type -assembly System.Windows.Forms

$main_form = New-Object System.Windows.Forms.Form

$main_form.Text = 'GUI to transfer files from one machine to another'
$main_form.Width = 600
$main_form.Height = 400

# $main_form.AutoSize = $true

$main_form.ShowDialog()

