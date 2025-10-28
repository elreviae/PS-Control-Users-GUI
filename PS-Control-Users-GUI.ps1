<#
Script PowerShell avec interface Windows Forms 
pour contrôle des utilisateurs connectés sur un serveur (Exemple : Serveur RDS)

Maxime DES TOUCHES - 2025 | https://github.com/elreviae ------------
#>

# Load required assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Create the main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "User Sessions Monitor"
$form.Size = New-Object System.Drawing.Size(810, 500)
$form.StartPosition = "CenterScreen"
$form.MinimumSize = New-Object System.Drawing.Size(600, 400)
$form.AutoSize = $true
$form.Font = New-Object System.Drawing.Font("Segoe UI", 10)

# Create the server label
$serverLabel = New-Object System.Windows.Forms.Label
$serverLabel.Location = New-Object System.Drawing.Point(10, 10)
$serverLabel.Size = New-Object System.Drawing.Size(120, 30)
$serverLabel.Text = "Server name or IP :"
$serverLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$form.Controls.Add($serverLabel)

# Create the server input text box
$serverTextBox = New-Object System.Windows.Forms.TextBox
$serverTextBox.Location = New-Object System.Drawing.Point(140, 10)
$serverTextBox.Size = New-Object System.Drawing.Size(200, 30)
$serverTextBox.Text = ""  # Empty by default for local machine
$form.Controls.Add($serverTextBox)

# Create the Refresh button
$refreshButton = New-Object System.Windows.Forms.Button
$refreshButton.Location = New-Object System.Drawing.Point(420, 10)
$refreshButton.Size = New-Object System.Drawing.Size(120, 30)
$refreshButton.Text = "Refresh Users"
$refreshButton.BackColor = [System.Drawing.Color]::LightGray
$refreshButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$form.Controls.Add($refreshButton)

# Create the Close button
$closeButton = New-Object System.Windows.Forms.Button
$closeButton.Location = New-Object System.Drawing.Point(640, 10)
$closeButton.Size = New-Object System.Drawing.Size(120, 30)
$closeButton.Text = "Close"
$closeButton.BackColor = [System.Drawing.Color]::LightGray
$closeButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$form.Controls.Add($closeButton)

# Create the DataGridView for tabular display
$dataGridView = New-Object System.Windows.Forms.DataGridView
$dataGridView.Location = New-Object System.Drawing.Point(10, 50)
$dataGridView.Size = New-Object System.Drawing.Size(780, 400)
$dataGridView.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
$dataGridView.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::AutoSize
$dataGridView.AllowUserToAddRows = $false
$dataGridView.ReadOnly = $true
$dataGridView.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
$dataGridView.BorderStyle = [System.Windows.Forms.BorderStyle]::None
$dataGridView.BackgroundColor = [System.Drawing.Color]::White
$dataGridView.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::LightGray
$form.Controls.Add($dataGridView)

# Define columns for DataGridView with sorting enabled
$columns = @("USERNAME", "SESSIONNAME", "ID", "STATE", "IDLE TIME", "LOGON TIME")
foreach ($col in $columns) {
    $column = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $column.Name = $col
    $column.HeaderText = $col
    $column.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::Automatic
    $dataGridView.Columns.Add($column)
}

# Function to parse quser output and populate DataGridView
function Update-QuserData {
    param ($grid, $server)
    try {
        # Clear existing rows
        $grid.Rows.Clear()

        # Execute quser with optional server parameter
        $quserArgs = if ($server) { "/server:$server" } else { "" }
        $quserOutput = & quser $quserArgs 2>&1
        if ($quserOutput -is [System.Management.Automation.ErrorRecord]) {
            [System.Windows.Forms.MessageBox]::Show("Error executing quser: $($quserOutput.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }

        # Parse quser output (skip header line)
        $lines = $quserOutput | Select-Object -Skip 1
        foreach ($line in $lines) {
            # Split line into fields, handling spaces and variable column widths
            $fields = $line -split '\s+', 6 | ForEach-Object { $_.Trim() }
            if ($fields.Count -ge 6) {
                # Adjust fields based on quser output format
                if ($fields[0].StartsWith(">")) {
                    $fields[0] = $fields[0].Substring(1) # Remove '>' for active session
                }
                $grid.Rows.Add($fields[0], $fields[1], $fields[2], $fields[3], $fields[4], $fields[5])
            }
        }
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Error processing quser output: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}

# Timer for auto-refresh (every 30 seconds)
$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 30000  # 30 seconds in milliseconds
$timer.Add_Tick({
    Update-QuserData -grid $dataGridView -server $serverTextBox.Text
})
$timer.Start()

# Refresh button click event
$refreshButton.Add_Click({
    Update-QuserData -grid $dataGridView -server $serverTextBox.Text
})

# Close button click event
$closeButton.Add_Click({
    $timer.Stop()  # Stop the timer to prevent further refreshes
    $form.Close()  # Close the form
})

# Handle form resize to adjust DataGridView size
$form.Add_Resize({
    $dataGridView.Width = $form.ClientSize.Width - 20
    $dataGridView.Height = $form.ClientSize.Height - 60
})

# Initial population of data
Update-QuserData -grid $dataGridView -server $serverTextBox.Text

# Show the form
$form.ShowDialog() | Out-Null