# Load Assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Create Form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Zip Folder Creator"
$form.Width = 500
$form.Height = 350
$form.StartPosition = "CenterScreen"

# Add ListBox for Verbose Output
$verboseOutput = New-Object System.Windows.Forms.ListBox
$verboseOutput.Width = 350
$verboseOutput.Height = 100
$verboseOutput.Location = New-Object System.Drawing.Point(50, 150)
$form.Controls.Add($verboseOutput)

# Add Button to Select Folder
$button1 = New-Object System.Windows.Forms.Button
$button1.Text = "Select Folder"
$button1.Location = New-Object System.Drawing.Point(50, 50)
$button1.Width = 150

# Event Handler for Button Click
$button1.Add_Click({
    # Select Folder to Compress
    $folder_choice = New-Object System.Windows.Forms.FolderBrowserDialog
    $folder_choice.Description = "Select a folder to Zip"
    $folder_choice.ShowDialog()
    $selected_folder = $folder_choice.SelectedPath
    [System.Windows.Forms.MessageBox]::Show("Selected folder is: $($selected_folder)")

    # Select Destination Folder
    $destination = New-Object System.Windows.Forms.FolderBrowserDialog
    $destination.Description = "Select Destination Folder"
    $destination.ShowDialog()
    $destination_folder = $destination.SelectedPath + "\Backup.zip"
    [System.Windows.Forms.MessageBox]::Show("Destination folder is: $($destination_folder)")

    # Perform Compression and Update Verbose Output
    $verboseOutput.Items.Add("Starting compression...")
    try {
        Compress-Archive -Path $selected_folder -DestinationPath $destination_folder -CompressionLevel Optimal -Verbose | ForEach-Object {
            $verboseOutput.Items.Add($_)  # Add each verbose message to the ListBox
        }
        $verboseOutput.Items.Add("Compression completed successfully!")
        [System.Windows.Forms.MessageBox]::Show("Compression is finished!")
    } catch {
        $verboseOutput.Items.Add("Error: $_")
        [System.Windows.Forms.MessageBox]::Show("An error occurred during compression.")
    }
})

# Add Button to Form
$form.Controls.Add($button1)

# Add Label with Your Name at the Bottom
$label = New-Object System.Windows.Forms.Label
$label.Text = "Developed by: Chinthakindi Srivatsava"
$label.AutoSize = $true
$label.Location = New-Object System.Drawing.Point(50, 300)  # Position at bottom right
$form.Controls.Add($label)


# Display Form
$form.ShowDialog()