# https://github.com/DeBoogie

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

function Get-ConfigSettings {
    $configPath = Join-Path $PSScriptRoot "config.json"
    $config = Get-Content $configPath -Raw | ConvertFrom-Json
    
    if (-not $config.DC -or $config.DC.Count -eq 0 -or -not ($config.DC | Where-Object { $_ -is [string] -and $_.Length -gt 0 })) {
        Write-LogMessage "Configuration error: DC array is not defined or empty in the config.json file" -level "ERROR"
        throw "Configuration error: At least one Domain Controller must be defined in the DC array"
    }
    
    return $config
}

$form = New-Object System.Windows.Forms.Form
$form.Text = "AD Tool"
$form.Size = New-Object System.Drawing.Size(800, 600)
$form.StartPosition = "CenterScreen"
$form.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right

$menuStrip = New-Object System.Windows.Forms.MenuStrip
$fileMenu = New-Object System.Windows.Forms.ToolStripMenuItem
$fileMenu.Text = "File"
$toolsMenu = New-Object System.Windows.Forms.ToolStripMenuItem
$toolsMenu.Text = "Tools"
$tableMenu = New-Object System.Windows.Forms.ToolStripMenuItem
$tableMenu.Text = "Table"
$exportMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$exportMenuItem.Text = "Export to Excel"
$exportCsvMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$exportCsvMenuItem.Text = "Export to CSV"
$exitMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$exitMenuItem.Text = "Exit"
$batchSearchMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$batchSearchMenuItem.Text = "Batch Search"
$resetResultsMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$resetResultsMenuItem.Text = "Clear Table"

$fileMenu.DropDownItems.Add($exitMenuItem)
$toolsMenu.DropDownItems.Add($batchSearchMenuItem)
$toolsMenu.DropDownItems.Add($exportMenuItem)
$toolsMenu.DropDownItems.Add($exportCsvMenuItem)
$tableMenu.DropDownItems.Add($resetResultsMenuItem)

$menuStrip.Items.Add($fileMenu)
$menuStrip.Items.Add($toolsMenu)
$menuStrip.Items.Add($tableMenu)
$form.MainMenuStrip = $menuStrip
$form.Controls.Add($menuStrip)

$userLabel = New-Object System.Windows.Forms.Label
$userLabel.Location = New-Object System.Drawing.Point(410, 70)
$userLabel.Size = New-Object System.Drawing.Size(360, 20)
$userLabel.Text = "Properties"
$userLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($userLabel)

$listBox = New-Object System.Windows.Forms.CheckedListBox
$listBox.Location = New-Object System.Drawing.Point(410, 100)
$listBox.Size = New-Object System.Drawing.Size(360, 455)
$listBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($listBox)

$dcLabel = New-Object System.Windows.Forms.Label
$dcLabel.Location = New-Object System.Drawing.Point(10, 40)
$dcLabel.Size = New-Object System.Drawing.Size(100, 20)
$dcLabel.Text = "Domain Controller:"
$dcLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
$form.Controls.Add($dcLabel)

$dcComboBox = New-Object System.Windows.Forms.ComboBox
$dcComboBox.Location = New-Object System.Drawing.Point(110, 40)
$dcComboBox.Size = New-Object System.Drawing.Size(280, 20)
$dcComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$dcComboBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$config = Get-ConfigSettings
$dcComboBox.Items.AddRange($config.DC)
$dcComboBox.SelectedIndex = 0
$form.Controls.Add($dcComboBox)

$searchLabel = New-Object System.Windows.Forms.Label
$searchLabel.Location = New-Object System.Drawing.Point(10, 70)
$searchLabel.Size = New-Object System.Drawing.Size(360, 20)
$searchLabel.Text = "Search AD User (SAM or UPN)"
$searchLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
$form.Controls.Add($searchLabel)

$searchBox = New-Object System.Windows.Forms.TextBox
$searchBox.Location = New-Object System.Drawing.Point(10, 100)
$searchBox.Size = New-Object System.Drawing.Size(280, 20)
$searchBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
$form.Controls.Add($searchBox)

$searchButton = New-Object System.Windows.Forms.Button
$searchButton.Location = New-Object System.Drawing.Point(300, 100)
$searchButton.Size = New-Object System.Drawing.Size(90, 20)
$searchButton.Text = "Search"
$searchButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
$form.Controls.Add($searchButton)

$resultTable = New-Object System.Windows.Forms.DataGridView
$resultTable.Location = New-Object System.Drawing.Point(10, 170)
$resultTable.Size = New-Object System.Drawing.Size(380, 380)
$resultTable.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$resultTable.AllowUserToDeleteRows = $true
$resultTable.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
$resultTable.AutoSizeRowsMode = [System.Windows.Forms.DataGridViewAutoSizeRowsMode]::AllCells
$resultTable.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::AutoSize
$form.Controls.Add($resultTable)

$contextMenu = New-Object System.Windows.Forms.ContextMenuStrip
$deleteMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$deleteMenuItem.Text = "Delete Row"
$deleteMenuItem.Add_Click({
        if ($resultTable.SelectedRows.Count -gt 0) {
            if ([System.Windows.Forms.MessageBox]::Show("Are you sure you want to delete this row?", "Confirm Delete", [System.Windows.Forms.MessageBoxButtons]::YesNo) -eq [System.Windows.Forms.DialogResult]::Yes) {
                $selectedIndex = $resultTable.SelectedRows[0].Index
                $resultTable.Rows.RemoveAt($selectedIndex)
                $script:searchResults = @($script:searchResults | Select-Object -First $selectedIndex) + @($script:searchResults | Select-Object -Skip ($selectedIndex + 1))
                Write-LogMessage "Deleted row at index $selectedIndex" -level "INFO"
            }
        }
    })
$contextMenu.Items.Add($deleteMenuItem)
$resultTable.ContextMenuStrip = $contextMenu
$script:searchResults = @()

function Write-LogMessage {
    param (
        [string]$message,
        [string]$level = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logPath = Join-Path $PSScriptRoot "ADTool.log"
    $logEntry = "[$timestamp] [$level] $message"
    Add-Content -Path $logPath -Value $logEntry
    Write-Host $logEntry
}

function Get-ADUserProperties {
    param (
        [string]$username
    )
    if ([string]::IsNullOrWhiteSpace($username)) {
        Write-LogMessage "Username cannot be empty" -level "ERROR"
        return 0
    }
    try {
        Write-LogMessage "Searching for user: $username" -level "INFO"
        $checkedIndices = $listBox.CheckedIndices
        $checkedItems = New-Object System.Collections.ArrayList
        foreach ($index in $checkedIndices) {
            $checkedItems.Add($listBox.Items[$index])
        }
        
        $searchType = if ($username -match '@') { 'UserPrincipalName' } else { 'SamAccountName' }

        $selectedDC = $dcComboBox.SelectedItem
        $user = Get-ADUser -Filter "$searchType -eq '$username'" -Properties * -Server $selectedDC
        if ($null -ne $user) {
            $listBox.Items.Clear()
            $propertyList = @()
            $user | Get-Member -MemberType Property | ForEach-Object {
                $propertyName = $_.Name
                $propertyValue = $user.$($propertyName)
                $propertyList += "$propertyName`: $propertyValue"
            }
            $listBox.Items.AddRange($propertyList)

            foreach ($item in $checkedItems) {
                $propertyName = $item.Split("`: ")[0]
                for ($i = 0; $i -lt $listBox.Items.Count; $i++) {
                    if ($listBox.Items[$i].StartsWith("$propertyName`:")) {
                        $listBox.SetItemChecked($i, $true)
                        break
                    }
                }
            }

            $propertyList.Clear()
            [System.GC]::Collect()
            Write-LogMessage "User found: $username" -level "SUCCESS"
            $script:searchResults += @($user)
            return $user
        }
        else {
            Write-LogMessage "User not found: $username" -level "ERROR"
            return 0
        }
    }
    catch {
        Write-LogMessage "Error searching for user $username : $($_.Exception.Message)" -level "ERROR"
        Write-LogMessage "Stack Trace: $($_.ScriptStackTrace)" -level "DEBUG"
        return 0
    }
}

$searchButton.Add_Click({
        $username = $searchBox.Text
        Get-ADUserProperties -username $username
        UpdateResultTable
    })

$batchSearchMenuItem.Add_Click({
        $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $openFileDialog.Filter = "Text Files (*.txt)|*.txt|CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
    
        if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {

            $usernames = Get-Content $openFileDialog.FileName
            $total = $usernames.Count

            $failed = @()
        
            $progressForm = New-Object System.Windows.Forms.Form
            $progressForm.Text = "Batch Search Progress"

            $progressForm.Size = New-Object System.Drawing.Size(400, 300)
            $progressForm.StartPosition = "CenterScreen"
            $progressForm.FormBorderStyle = "FixedDialog"
            $progressForm.MaximizeBox = $false
        
            $progressLabel = New-Object System.Windows.Forms.Label
            $progressLabel.Location = New-Object System.Drawing.Point(10, 20)

            $progressLabel.Size = New-Object System.Drawing.Size(360, 40)
            $progressForm.Controls.Add($progressLabel)
        
            $failedListBox = New-Object System.Windows.Forms.ListBox
            $failedListBox.Location = New-Object System.Drawing.Point(10, 70)
            $failedListBox.Size = New-Object System.Drawing.Size(360, 150)
            $progressForm.Controls.Add($failedListBox)
        
            $progressForm.Show()
        
            foreach ($username in $usernames) {
                $progressLabel.Text = "Processing... $($usernames.IndexOf($username) + 1)/$total`nCurrent user: $username"
                Write-LogMessage "Processing user: $username" -level "BATCH"
                try {
                    $result = Get-ADUserProperties -username $username
                    if (0 -eq $result) {
                        $failed += $username
                        $failedListBox.Items.Add("Failed: $username")
                        Write-LogMessage "Failed to get user properties for $username" -level "BATCH"
                    }
                }
                catch {
                    $failed += $username
                    $failedListBox.Items.Add("Failed: $username")
                    Write-LogMessage "Failed to get user properties for $username" -level "BATCH"
                }
                
                UpdateResultTable
                [System.Windows.Forms.Application]::DoEvents()
            }
        
            $progressForm.Close()

            # Export failed users to a file
            if ($failed.Count -gt 0) {
                $failedFilePath = [System.IO.Path]::ChangeExtension($openFileDialog.FileName, "failed.txt")
                $failed | Out-File -FilePath $failedFilePath
                
                [System.Windows.Forms.MessageBox]::Show(
                    "Batch search completed.`nTotal processed: $total`nFailed: $($failed.Count)`n`nFailed users have been exported to:`n$failedFilePath",
                    "Batch Search Complete",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Information
                )
            }
            else {
                [System.Windows.Forms.MessageBox]::Show(
                    "Batch search completed successfully.`nTotal processed: $total`nAll users found!",
                    "Batch Search Complete",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Information
                )
            }
        }
    })

$resetButton.Add_Click({
        $result = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to clear the results?", "Confirm Reset", [System.Windows.Forms.MessageBoxButtons]::YesNo)
        if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
            $script:searchResults = @()
            $resultTable.Rows.Clear()
            $resultTable.Columns.Clear()
            $searchBox.Clear()
        }
    })

$exportAction = {
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx"
    $saveFileDialog.DefaultExt = "xlsx"
    $saveFileDialog.AddExtension = $true

    if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        function Export-ToExcel {
            param ($saveFileDialog)
            
            $excel = $null
            try {
                $excel = New-Object -ComObject Excel.Application
                $excel.Visible = $false
                $workbook = $excel.Workbooks.Add()
                $worksheet = $workbook.Worksheets.Item(1)
                
                for ($i = 0; $i -lt $resultTable.Columns.Count; $i++) {
                    $worksheet.Cells.Item(1, $i + 1) = $resultTable.Columns[$i].HeaderText
                }
                
                for ($i = 0; $i -lt $resultTable.Rows.Count; $i++) {
                    for ($j = 0; $j -lt $resultTable.Columns.Count; $j++) {
                        $worksheet.Cells.Item($i + 2, $j + 1) = $resultTable.Rows[$i].Cells[$j].Value
                    }
                }
                
                $usedRange = $worksheet.UsedRange
                $usedRange.EntireColumn.AutoFit() | Out-Null
                $usedRange.AutoFilter() | Out-Null
                
                $workbook.SaveAs($saveFileDialog.FileName)
                $workbook.Close()
            }
            catch {
                Write-LogMessage "Excel export failed: $_" -level "ERROR"
                throw
            }
            finally {
                if ($excel) {
                    $excel.Quit()
                    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
                    [System.GC]::Collect()
                }
            }
        }

        try {
            Export-ToExcel -saveFileDialog $saveFileDialog
            [System.Windows.Forms.MessageBox]::Show("Export completed successfully!", "Export Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            Write-LogMessage "Exported results to $($saveFileDialog.FileName)" -level "EXPORT"
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Export failed. Please check the log for details.", "Export Failed", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
}

$exportCsvAction = {
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "CSV Files (*.csv)|*.csv"
    $saveFileDialog.DefaultExt = "csv"
    $saveFileDialog.AddExtension = $true

    if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $csvContent = @()

        $headers = ($resultTable.Columns | ForEach-Object { $_.HeaderText }) -join ","
        $csvContent += $headers
        
        for ($i = 0; $i -lt $resultTable.Rows.Count; $i++) {
            $row = @()
            for ($j = 0; $j -lt $resultTable.Columns.Count; $j++) {
                $row += $resultTable.Rows[$i].Cells[$j].Value
            }
            $csvContent += ($row -join ",")
        }
        
        [System.IO.File]::WriteAllLines($saveFileDialog.FileName, $csvContent)

        [System.Windows.Forms.MessageBox]::Show("CSV export completed successfully!", "Export Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        Write-LogMessage "Exported results to $($saveFileDialog.FileName)" -level "EXPORT"
    }
}

$exportMenuItem.Add_Click($exportAction)
$exportCsvMenuItem.Add_Click($exportCsvAction)

function UpdateResultTable {
    $resultTable.Rows.Clear()
    $resultTable.Columns.Clear()
    
    $checkedProperties = $listBox.CheckedItems | ForEach-Object {
        $_.Split("`: ")[0]
    }
    
    foreach ($property in $checkedProperties) {
        $column = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $column.Name = $property
        $column.HeaderText = $property
        $resultTable.Columns.Add($column)
    }
    
    if ($checkedProperties.Count -gt 0) {
        foreach ($user in $script:searchResults) {
            $row = $resultTable.Rows.Add()
            foreach ($property in $checkedProperties) {
                $value = $user.$property
                $resultTable.Rows[$row].Cells[$property].Value = $value
            }
        }
    }
}

$listBox.Add_ItemCheck({
        param($eventSender, $e)
        
        $form.BeginInvoke([System.Action] {
                UpdateResultTable
            })
    })

$resultTable.Add_KeyDown({
        param($eventSender, $e)
        try {
            if ($e.KeyCode -eq 'Delete' -and $resultTable.SelectedRows.Count -gt 0) {
                $selectedIndex = $resultTable.SelectedRows[0].Index
                $resultTable.Rows.RemoveAt($selectedIndex)
                $script:searchResults = @($script:searchResults | Select-Object -First $selectedIndex) + @($script:searchResults | Select-Object -Skip ($selectedIndex + 1))
            }
        }
        catch {
            Write-LogMessage "Error deleting row: $_" -level "ERROR"
        }
    })

$searchBox.Add_KeyDown({
        param($senderObject, $e)
        if ($e.KeyCode -eq 'Enter') {
            $username = $searchBox.Text
            Get-ADUserProperties -username $username
            UpdateResultTable
        }
    })

$resetResultsMenuItem.Add_Click({
        $result = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to clear the results?", "Confirm Reset", [System.Windows.Forms.MessageBoxButtons]::YesNo)
        if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
            try {
                $script:searchResults = @()
                $resultTable.Rows.Clear()
                $resultTable.Columns.Clear()
                $searchBox.Clear()
                Write-LogMessage "Table cleared" -level "INFO"
            }
            catch {
                Write-LogMessage "Error clearing results: $_" -level "ERROR"
            }
        }
    })
    

$form.Add_Shown({ $form.Activate() })
[void]$form.ShowDialog()
