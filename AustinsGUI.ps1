[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 
Add-Type -AssemblyName presentationframework, presentationcore
$wpf = @{ }
$wpfHelp = @{ }
$inputXML = @"
<Window x:Name="MainWindow1" x:Class="Bookmark_Manager_GUI.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:Bookmark_Manager_GUI"
        mc:Ignorable="d"
        Title="Bookmark Manager" Height="450" Width="800">
    <Grid>
        <Grid.ColumnDefinitions>
            <ColumnDefinition/>
        </Grid.ColumnDefinitions>
        <TextBox x:Name="JSONInput" HorizontalAlignment="Left" Height="50" Margin="10,22,0,0" TextWrapping="Wrap" Text="Paste your JSON located in the 'Managed Bookmarks' GPO" VerticalAlignment="Top" Width="640" AcceptsReturn="True" VerticalScrollBarVisibility="Visible" />
        <TextBox x:Name="JSONOutput" HorizontalAlignment="Left" Height="50" Margin="10,77,0,0" TextWrapping="Wrap" Text="Once you import the CSV, the properly formatted JSON text will show up here" VerticalAlignment="Top" Width="640"  AcceptsReturn="True" VerticalScrollBarVisibility="Visible"/>
        <Button x:Name="JSONSubmit" Content="Convert JSON to CSV" HorizontalAlignment="Left" Margin="658,22,0,0" VerticalAlignment="Top" Width="123" Height="50"/>
        <Button x:Name="convertCSVtoJSON" Content="Convert CSV to JSON" HorizontalAlignment="Left" Margin="658,77,0,0" VerticalAlignment="Top" Width="123" Height="23"/>
        <Button x:Name="jsonCopy" Content="Copy to Clipboard" HorizontalAlignment="Left" Margin="658,105,0,0" VerticalAlignment="Top" Width="123" Height="22"/>
        <DataGrid x:Name="dataGrid1" IsReadOnly="True" HorizontalAlignment="Left" Height="250" Margin="10,136,0,0" VerticalAlignment="Top" Width="771" AutoGenerateColumns="False" >
            <DataGrid.Columns>
                <DataGridTextColumn Header="Name" Binding="{Binding 'Name'}" MinWidth="288" />
                <DataGridTextColumn Header="URL" Binding="{Binding 'URL'}" MinWidth="288" />
                <DataGridTextColumn Header="Folder" Binding="{Binding 'Folder'}" MinWidth="170" />
            </DataGrid.Columns>
        </DataGrid>
        <Button x:Name="buttonClose" Content="Close" HorizontalAlignment="Left" Margin="706,391,0,0" VerticalAlignment="Top" Width="75"/>
        <Button x:Name="buttonHelp" Content="Help" HorizontalAlignment="Left" Margin="626,391,0,0" VerticalAlignment="Top" Width="75"/>
        <Label x:Name="JSONErrorOutput" Content="" HorizontalAlignment="Left" Margin="10,0,0,0" VerticalAlignment="Top" Height="22" Width="640" Foreground="Red" FontSize="10"/>
    </Grid>
</Window>

"@
$inputHelpXML = @"
<Window x:Name="helpWindow" x:Class="Bookmark_Manager_GUI.Window1"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:Bookmark_Manager_GUI"
        mc:Ignorable="d"
        Title="Help" Height="450" Width="800">
    <Grid>
        <Label Content="Instructions" HorizontalAlignment="Left" Margin="10,29,0,0" VerticalAlignment="Top" Width="774" Height="36" FontSize="24" HorizontalContentAlignment="Center"/>
        <Button x:Name="buttonCloseHelp" Content="Close" HorizontalAlignment="Center" Margin="10,391,0,0" VerticalAlignment="Top" Width="75"/>
        <Label Content="Step 1: Open 'Group Policy Management' to the appropriate group policies for managing Google Chrome and Microsoft Edge Bookmarks " HorizontalAlignment="Left" Margin="10,85,0,0" VerticalAlignment="Top" Width="774"/>
        <Label Content="Step 2: Navigate to User Configuration &gt; Policies &gt; Administrative Templates &gt; (Google Chrome or Microsoft Edge) -&gt; &#xD;&#xA;Managed Bookmarks/Favorites " HorizontalAlignment="Left" Margin="10,112,0,0" VerticalAlignment="Top" Width="774" Height="48"/>
        <Label Content="Step 5: Edit the CSV file in excel following the same pattern as presented " HorizontalAlignment="Left" Margin="10,211,0,0" VerticalAlignment="Top" Width="774"/>
        <Label Content="Step 4: Click 'Process JSON to CSV' and the file will be automatically saved in a temporary location and opened in Excel " HorizontalAlignment="Left" Margin="10,185,0,0" VerticalAlignment="Top" Width="774"/>
        <Label Content="Step 3: Copy the JSON text from the input field and copy it to the JSON input field in the PowerShell form. " HorizontalAlignment="Left" Margin="10,159,0,0" VerticalAlignment="Top" Width="774"/>
        <Label Content="Step 6: Once finished, save the CSV file in the same location. If you want to save it somewhere else, remember where and open it in the next step" HorizontalAlignment="Left" Margin="10,237,0,0" VerticalAlignment="Top" Width="774"/>
        <Label Content="Step 7: Click 'Convert CSV to JSON' in the PowerShell Form and open the CSV file you edited (The default save location and file name will be &#xD;&#xA; automatically entered for your convenience) " HorizontalAlignment="Left" Margin="10,263,0,0" VerticalAlignment="Top" Width="774"/>
        <Label Content="Step 8: Once the CSV processess you will be presented with the properly formatted text in the output field and a datagrid showing what data&#xD;&#xA; was picked up from the CSV (*Note: If the datagrid is not accurate or does not show anything, then there is an error with the CSV format) &#xA;" HorizontalAlignment="Left" Margin="10,300,0,0" VerticalAlignment="Top" Width="774" Height="42"/>
        <Label Content="Step 9: Click 'Copy to Clipboard' to copy the output JSON text and paste it into the appropriate Chrome/Edge Group Policy Objects" HorizontalAlignment="Left" Margin="10,336,-0.4,0" VerticalAlignment="Top" Width="784"/>

    </Grid>
</Window>

"@
$inputXMLClean = $inputXML -replace 'mc:Ignorable="d"','' -replace "x:N",'N' -replace 'x:Class=".*?"','' -replace 'd:DesignHeight="\d*?"','' -replace 'd:DesignWidth="\d*?"',''
$inputHelpXMLClean = $inputHelpXML -replace 'mc:Ignorable="d"','' -replace "x:N",'N' -replace 'x:Class=".*?"','' -replace 'd:DesignHeight="\d*?"','' -replace 'd:DesignWidth="\d*?"',''



[xml]$xaml = $inputXMLClean
[xml]$xamlHelp = $inputHelpXMLClean

$reader = New-Object System.Xml.XmlNodeReader $xaml
$tempform = [Windows.Markup.XamlReader]::Load($reader)
$namedNodes = $xaml.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]")
$namedNodes | ForEach-Object {$wpf.Add($_.Name, $tempform.FindName($_.Name))}

$readerHelp = New-Object System.Xml.XmlNodeReader $xamlHelp
$tempformHelp = [Windows.Markup.XamlReader]::Load($readerHelp)
$namedNodesHelp = $xamlHelp.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]")
$namedNodesHelp | ForEach-Object {$wpfHelp.Add($_.Name, $tempformHelp.FindName($_.Name))}


function parseJSONtoCSV($JSONInput){

    try{
    $jsonConverted = $wpf.JSONInput.Text | ConvertFrom-Json
    $wpf.JSONErrorOutput.Content = ""
    }catch{
    $wpf.JSONErrorOutput.Content = "You did not enter a properly formatted JSON string. Please ensure you copy and pasted the entire string from the GPO and try again"
    return
    }

    try{
    #$SaveFileDialog = New-Object Windows.Forms.SaveFileDialog   
    #$SaveFileDialog.initialDirectory = [System.IO.Directory]::GetCurrentDirectory()   
    #$SaveFileDialog.title = "Save JSON as a CSV (For Editing)"   
    #$SaveFileDialog.filter = 'CSV (*.csv)|*.csv|All Files|*.*'
    #$SaveFileDialog.ShowHelp = $True   
    #$SaveFileDialog.ShowDialog()
    $global:pathToOutputFile = "$env:TEMP\tempBookmarks.csv"   #OLD$SaveFileDialog.FileName

    $firstEntry = [PSCustomObject]@{ Column1 = "Enter your bookmarks here"; Column2 = ""; Column3 = "" }
    $firstEntry | Export-Csv $pathToOutputFile -NoTypeInformation
    }catch{return}
    ForEach ($entry in $jsonConverted){
        if($entry.toplevel_name -ne $null){
            $csvOutput = New-Object PSObject -Property @{ Column1 = "Main Folder Name"; Column2 = $entry.toplevel_name; Column3 = "" } | Export-CSV $pathToOutputFile -Append
            $csvOutput = New-Object PSObject -Property @{ Column1 = "Name"; Column2 = "URL"; Column3 = "Folder (Blank if not in a folder)" } | Export-CSV $pathToOutputFile -Append
            }
        elseif($entry.children -ne $null){
            $entry.children | ForEach-Object{
            $name = $_.name
            $url = $_.url
            $csvOutput = New-Object PSObject -Property @{ Column1 = $name; Column2 = $url; Column3 = $entry.name } | Export-CSV $pathToOutputFile -Append
  
            }
        }else{
            $csvOutput = New-Object PSObject -Property @{ Column1 = $entry.name; Column2 = $entry.url; Column3 = "" } | Export-CSV $pathToOutputFile -Append
        }
    }

    start excel $pathToOutputFile
}

function parseJSONToGrid($csvInput){
        $wpf.dataGrid1.Items.Clear()
        $script:CsvData.Clear()
        $script:CsvData = New-Object System.Collections.ArrayList 
        $json = $csvInput | ConvertFrom-Json

        #Arrays
        $myArray = @()
        $topLevelArray = @()
        $index = 0

#region Process json to find out how many objects there are (i.e. top level, folders, links)
        $json | ForEach-Object{
            $topLevelArray += $index
            $index++
            }

        write-host "There are $index objects in the JSON" -ForegroundColor Yellow
        $index--
        $topLevelArray = $topLevelArray | Where-Object { $_ -ne 0}
        #endregion

#region Folder Identification and Recursion
        $folderAmount = 0
        $folderLocationArray = @()
        #resetIndex
        $index = 0

        $json | ForEach-Object{
           #if json has a value for children, this means that it is a folder
           if($json[$index].children){
           $folderAmount++
           $folderLocationArray += $index
           #Remove folder index from topLevelArray so that it is not added to the links at the end
           $topLevelArray = $topLevelArray | Where-Object { $_ -ne $index}
   
            } 
           $index++
        }

        write-host "There are $folderAmount folders.... Beginning folder recursion" -ForegroundColor Yellow

        #resetIndex
        $index = 0
        $folderLocationArray | ForEach-Object {
            $currentFolder = $_


            #For each child filter through the folder for links
            $json[$_].children | ForEach-Object{
                write-host $json[$currentFolder].children[$index].name -ForegroundColor Magenta
                write-host $json[$currentFolder].children[$index].url -ForegroundColor Magenta


            #Add filtered urls to an object
            $myObject = [PSCustomObject]@{
                Name = $json[$currentFolder].children[$index].name
                URL = $json[$currentFolder].children[$index].url
                Folder = $json[$currentFolder].name


            }
            #Add to Array
            $myArray += $myObject
            $wpf.dataGrid1.AddChild($myObject)
            $index++
                }
            $index = 0

        }

        #endregion


 #region Final Links.... Located Outside of Folders
        #Create object to signify final links, outside of folder
        $topLevelArray | ForEach-Object {
            $currentIndex = $_
            #Final Objects for Links
            $myObject = [PSCustomObject]@{
                Name = $json[$currentIndex].name
                URL = $json[$currentIndex].url
                Folder = ''

            }
            #Add to Array
            $myArray += $myObject
            $wpf.dataGrid1.AddChild($myObject)
        }
        #endregion
        #$script:CsvData.AddRange(($myArray))
        
        }

function parseCSVtoJSON{
    #Generate filebrowser object for selecting CSV to import
    $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
        InitialDirectory = $env:TEMP #OLD[Environment]::GetFolderPath('Desktop') 
        FileName = "tempBookmarks.csv"
        Filter = 'CSV (*.csv)|*.csv'
    }
    $FileBrowser.ShowDialog()
    $csv = import-csv $FileBrowser.FileName

    #Begin generatting a final JSON string
    $finalJSON = "["

    #Empty ending string used in region LINT the rest to filter ends and avoid unnecessary formatting errors caused by switching between folders and not folders
    $endingString = ""
    #region FILTER CSV
    $index = 1
    $filteredCSV = @()
    ForEach($object in $csv){
        $filteredCSV += $csv[$index]
        $index++
    }
    $csv = $filteredCSV
    $index = 2
    $filteredCSV = @()
    $filteredCSV += $csv[0]
    ForEach($object in $csv){
        $filteredCSV += $csv[$index]
        $index++
    }

    #endregion

    #region LINT Top Level
    $index = 0
    $filteredCSV | ForEach-Object{
        if($filteredCSV[$index].Column1 -eq "Main Folder Name"){
            $topLevelname = $filteredCSV[$index].Column2
            $finalJSON += "{`"toplevel_name`": `"$topLevelName`"}"
        }
        $index++
    }

    #endregion

    #region LINT the rest...
    $index = 1

    #current folder variable to be used for indentifying a change in folders
    $currentFolder = "None"

    #Variables for identifying when to place ] and not to for separating folders and individual links
    $firstFolder = $true
    $firstFinal = $true

    #If/Else hell for filtering through objects
    $filteredCSV.column1 | ForEach-Object{
        if($filteredCSV[$index].Column3 -eq '' -and $firstFinal -eq $true){
            $name = $filteredCSV[$index].Column1
            $url = $filteredCSV[$index].Column2
            $newStr = "]}, {`"name`": `"$name`", `"url`": `"$url`"}"
            $endingString += $newStr
            $firstFinal = $false
        }elseif($filteredCSV[$index].Column3 -eq '' -and $firstFinal -eq $false){
            $name = $filteredCSV[$index].Column1
            $url = $filteredCSV[$index].Column2
            $newStr = ", {`"name`": `"$name`", `"url`": `"$url`"}"
            $endingString += $newStr
        }elseif($filteredCSV[$index].Column3 -ne $currentFolder -and $firstFolder -eq $true){
            $folderName = $filteredCSV[$index].Column3
            $currentFolder = $folderName
            $name = $filteredCSV[$index].Column1
            $url = $filteredCSV[$index].Column2
            $newStr = ", {`"name`": `"$folderName`", `"children`":[{`"name`": `"$name`", `"url`": `"$url`"}"
            $finalJSON += $newStr
            $firstFolder = $false
        }elseif($filteredCSV[$index].Column3 -ne $currentFolder -and $firstFolder -eq $false){

            if($filteredCSV[$index].Column1 -eq $null -or $filteredCSV[$index].Column2 -eq $null){
            }else{
            $finalJSON += "]}"
            $folderName = $filteredCSV[$index].Column3
            $currentFolder = $folderName
            $name = $filteredCSV[$index].Column1
            $url = $filteredCSV[$index].Column2
            $newStr = ", {`"name`": `"$folderName`", `"children`":[{`"name`": `"$name`", `"url`": `"$url`"}"
            $finalJSON += $newStr
            }
        }elseif($filteredCSV[$index].Column3 -eq $currentFolder){
            $name = $filteredCSV[$index].Column1
            $url = $filteredCSV[$index].Column2
            $newStr = ", {`"name`": `"$name`", `"url`": `"$url`"}"
            $finalJSON += $newStr
        }
        $index++
    }

    $finalJSON += $endingString
    $finalJSON +="]"

    $wpf.JSONOutput.Text = $finalJSON
    $finalJSON | out-file "C:\Users\aubeaty\OneDrive - Microsoft\Documents\Tasks\Admin GUI\test2.json"
    parseJSONToGrid $finalJSON
       #endregion
    }
 
#region Create Icon from base64
# Create a streaming image by streaming the base64 string to a bitmap streamsource
$base64 = "AAABAAEAAAAAAAEAIAAFBQAAFgAAAIlQTkcNChoKAAAADUlIRFIAAAEAAAABAAgGAAAAXHKoZgAABMxJREFUeNrt2s9qVGccx+Hvb5LUKgkWatcVuhC8jNKdm1Bw0Vvp2osp7cZ96RXoPq5qulaoaNBKMufXRaZ0Y5IzEGHezPPAwEDeHMj75zOczEkAAAAAAAAAAIDh1ZxB7378LukkVZv/l7RFvdC0TFI5ePrnJ3/85I99c3QDTvSi9pLu/Pz9myuH7866aCdJ7yf1YPU7m3fMHPwrtkVOkxwl/eHiUZUkd5I8TLJnVkdc6FpWctRVJ3PGzwtAVZJ6kKqnSe7aGEMG4FWqHiV5cdGgaTpNUt8m06+dfGOdx1vnSr3N4ovDzvTs+gLw/9i7SQ7M85A+Jtm5+q6vdzp9kMT9wIA6Sad35qZ7sfa1GXlvfI6xjLvOawUAuGEEAAQAEABAAAABAAQAEABAAAABAAQAEABAAAABAAQAEABAAAABAAQAEABAAAABAAQAEABAAAABAAQAEABAAAABAAQAEABAAAABAAQAEABAAAABAAEwBSAAgAAAAgAIACAAgAAAAgAIACAAgAAAAgAIACAAgAAAYwegTBRsbwBEALY4AL16AVsYAEAAAAEABAAQAEAAAAEABAAQAEAAAAEABAAQAEAAAAEABAAQAEAAAAEABAAQAEAAAAEABAAQAEAAAAEABAAQAEAAAAEABAAEABAAQAAAAQAEABAAQAAAAQAEABAAQAAAAQAEABAAQAAAAQAEABAAQACAIQJQqxdjWmftfDCMvc6z13p3jQufJnmV5GOSNs+jbYp+neRsxtKdJXm9ej+ZurHWuZJ3q7N6jQGYlklylKpHSXbM85DOspyOL/9sqCR5WVkcZr0PBzanAcskf11vAFJJ+kOSF+4CRtVzl+7j+Toz7DoDAAAAAADM+2Jo/5c3qSSnnXTHN4GD+u/xvg8/ffXJn/fvB6t3nv8Z+kgvbiU9pX74+8rR6zzscSfJ/ZTHRAfdFmdJXub8e/4LderLSt1P2oNAY5qSHCf1fs7gWYu8erTg4dT5rZN9czze+a/k9aJy2Jc+5DOlUvfTy6dJ34unSga0OEn6cTI9v7YArOwluRcBGFKfv2asd++uDv/XZm3Ilb69OqvzcrH+HmLgBnyOsWyWtf6B434etpgAgAAAAgAIACAAgAAAAgAIACAAgAAAAgAIACAAgAAAAgAIACAAgAAAAgAIACAAgAAAAgAIACAAgAAAAgAIACAAgAAAAgAIACAAgAAAAgAIACAAIACmAAQAEABAAAABAAQAEABAAAABAAQAEABAAAABAAQAuDEBKJMFWxsA5x+2OADdJgu2NgCAAAACAAgAIACAAAACAAgAIACAAAACAAgAIACAAAACAAgAIACAAAACAAgAIACAAAACAAgAIACAAAACAAgAIACAAAACAAgACAAgAIAAAAIACAAgAIAAAAIACAAgAIAAAAIACAAgAIAAAAIACAAgAIAAAAIAbHgAypQNrk3BDbfWGd2de8WqLKvydrWBbKPxNsVJJdPl26OSZEoWJ0nfPn/PUOtc9S6p5dwOzArA7iKp5CiLHHZnZyO3tyRdNUOnlRwnyT8XjVrcSpLjpB8n2TNtQy71MrV3lJp3IGYFoDtJ5aQ7zzbynDn81zSPU5J6n0zPTcawAcj54XcoAAAAAAAAAABga/wL7cGxAuIgQ3kAAAAASUVORK5CYII="
$icon = New-Object System.Windows.Media.Imaging.BitmapImage
$icon.BeginInit()
$icon.StreamSource = [System.IO.MemoryStream][System.Convert]::FromBase64String($base64)
$icon.EndInit()

# Freeze() prevents memory leaks.
$icon.Freeze()

#endregion



 #region Add functionality to GUI elements
$wpf.JSONSubmit.add_Click({
        parseJSONtoCSV($wpf.JSONInput.Text)
	})
$wpf.JSONInput.add_GotFocus({
        $wpf.JSONInput.Text = ""
    })
$wpf.convertCSVtoJSON.add_Click({
        parseCSVtoJSON
	})
$wpf.jsonCopy.add_Click({
        $wpf.JSONOutput.Text | Set-Clipboard
    })
$wpf.buttonClose.add_Click({
        $wpf.MainWindow1.close()
    })
$wpf.buttonHelp.add_Click({
        $wpfHelp.helpWindow.showDialog()
    })
$wpfHelp.buttonCloseHelp.add_Click({
        $wpfHelp.helpWindow.hide()
    })
$wpf.MainWindow1.Icon = $icon
$wpfHelp.helpWindow.Icon = $icon
#endregion
 #MAIN
 $wpf.MainWindow1.ShowDialog()
