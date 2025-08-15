#GUI
function Show-UserGUI ([string] $initialDirectory) {

    Add-Type -AssemblyName System.Windows.Forms
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "PsSPM 0.3.3"
    $form.Size = New-Object System.Drawing.Size(290, 335)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $form.MaximizeBox = $false
    $form.TopMost = $true
    $groupBox1 = New-Object System.Windows.Forms.GroupBox
    $groupBox2 = New-Object System.Windows.Forms.GroupBox
    $checkBox1 = New-Object System.Windows.Forms.CheckBox
    $checkBox2 = New-Object System.Windows.Forms.CheckBox
    $Labelcsvbuf = New-Object System.Windows.Forms.Label
    $Labelsnmp = New-Object System.Windows.Forms.Label
    $LabelTcp = New-Object System.Windows.Forms.Label
    $LabelTcpMaxRetries = New-Object System.Windows.Forms.Label
    $LabelTcpRetryDelay = New-Object System.Windows.Forms.Label
    $LabelTcpThreads = New-Object System.Windows.Forms.Label
    $numericUpDownCsvBuf = New-Object System.Windows.Forms.NumericUpDown
    $numericUpDownSnmp = New-Object System.Windows.Forms.NumericUpDown
    $numericUpDownTcp = New-Object System.Windows.Forms.NumericUpDown
    $numericUpDownTcpMaxRetries = New-Object System.Windows.Forms.NumericUpDown
    $numericUpDownTcpRetryDelay = New-Object System.Windows.Forms.NumericUpDown
    $numericUpDownTcpThreads = New-Object System.Windows.Forms.NumericUpDown
    $form.Controls.Add($groupbox1)
    $form.Controls.Add($groupBox2)
    
    # GroupBox 1
    $groupBox1.Location = New-Object System.Drawing.Point(15,10)
    $groupBox1.size = New-Object System.Drawing.Size(110,130)
    $groupBox1.text = "Report"
    $groupBox1.Visible = $true
    $groupbox1.Controls.Add($checkBox1)
    $groupbox1.Controls.Add($checkBox2)
    $groupbox1.Controls.Add($Labelcsvbuf)
    $groupbox1.Controls.Add($numericUpDownCsvBuf)

    # GroupBox 2
    $groupBox2.Location = New-Object System.Drawing.Point(150,10)
    $groupBox2.size = New-Object System.Drawing.Size(110,270)
    $groupBox2.text = "TCP/SNMP"
    $groupBox2.Visible = $true
    $groupbox2.Controls.Add($Labelsnmp)
    $groupbox2.Controls.Add($numericUpDownSnmp)
    $groupbox2.Controls.Add($LabelTcp)
    $groupbox2.Controls.Add($numericUpDownTcp)
    $groupbox2.Controls.Add($LabelTcpMaxRetries)
    $groupbox2.Controls.Add($numericUpDownTcpMaxRetries)
    $groupbox2.Controls.Add($LabelTcpRetryDelay)
    $groupbox2.Controls.Add($numericUpDownTcpRetryDelay)
    $groupbox2.Controls.Add($LabelTcpThreads)
    $groupbox2.Controls.Add($numericUpDownTcpThreads)

    # CheckBox 1
    $checkBox1.Location = New-Object System.Drawing.Point(10,20)
    $checkBox1.Size = New-Object System.Drawing.Size(95,20)
    $checkBox1.Text = "HTML Report"
    $checkBox1.Checked = $script:HtmlFileReport
    $checkBox1.Add_CheckedChanged({ $script:HtmlFileReport = $checkBox1.Checked })

    # CheckBox 2
    $checkBox2.Location = New-Object System.Drawing.Point(10,44)
    $checkBox2.Size = New-Object System.Drawing.Size(95,20)
    $checkBox2.Text = "CSV Report"
    $checkBox2.Checked = $script:CsvFileReport
    $checkBox2.Add_CheckedChanged({ $script:CsvFileReport = $checkBox2.Checked })

    # CheckBox 3
    $checkBox3 = New-Object System.Windows.Forms.CheckBox
    $checkBox3.Location = New-Object System.Drawing.Point(20,145)
    $checkBox3.Size = New-Object System.Drawing.Size(95,20)
    $checkBox3.Text = "Log enable"
    $checkBox3.Checked = $script:WriteLog
    $checkBox3.Add_CheckedChanged({ $script:WriteLog = $checkBox3.Checked })
    $form.Controls.Add($checkBox3)

    # Label CsvBufferSize
    $Labelcsvbuf.Text = "CSV Buffer Size:"
    $Labelcsvbuf.Location = New-Object System.Drawing.Point(10, 80)
    $Labelcsvbuf.AutoSize = $true

    # Create a CsvBufferSize control
    $numericUpDownCsvBuf.Location = New-Object System.Drawing.Point(10, 100)
    $numericUpDownCsvBuf.Size = New-Object System.Drawing.Size(90,30)
    $numericUpDownCsvBuf.Minimum = 5
    $numericUpDownCsvBuf.Maximum = 500
    $numericUpDownCsvBuf.Value = $CsvBufferSize
    $numericUpDownCsvBuf.Increment = 5

    # Label SNMP timeout
    $Labelsnmp.Text = "SNMP timeout ms:"
    $Labelsnmp.Location = New-Object System.Drawing.Point(10, 20)
    $Labelsnmp.AutoSize = $true

    # Create a SNMP control
    $numericUpDownSnmp.Location = New-Object System.Drawing.Point(10, 40)
    $numericUpDownSnmp.Size = New-Object System.Drawing.Size(90,30)
    $numericUpDownSnmp.Minimum = 50
    $numericUpDownSnmp.Maximum = 10000
    $numericUpDownSnmp.Value = $TimeoutMsUDP
    $numericUpDownSnmp.Increment = 50

    # Label TCP timeout
    $LabelTcp.Text = "TCP timeout ms:"
    $LabelTcp.Location = New-Object System.Drawing.Point(10, 70)
    $LabelTcp.AutoSize = $true

    # Create a TCP control
    $numericUpDownTcp.Location = New-Object System.Drawing.Point(10, 90)
    $numericUpDownTcp.Size = New-Object System.Drawing.Size(90,30)
    $numericUpDownTcp.Minimum = 50
    $numericUpDownTcp.Maximum = 10000
    $numericUpDownTcp.Value = $TimeoutMsTCP
    $numericUpDownTcp.Increment = 50

    # Label TCP retry
    $LabelTcpMaxRetries.Text = "Max Retries:"
    $LabelTcpMaxRetries.Location = New-Object System.Drawing.Point(10, 120)
    $LabelTcpMaxRetries.AutoSize = $true

    # Create a TCP retry control
    $numericUpDownTcpMaxRetries.Location = New-Object System.Drawing.Point(10, 140)
    $numericUpDownTcpMaxRetries.Size = New-Object System.Drawing.Size(90,30)
    $numericUpDownTcpMaxRetries.Minimum = 1
    $numericUpDownTcpMaxRetries.Maximum = 20
    $numericUpDownTcpMaxRetries.Value = $MaxRetries
    $numericUpDownTcpMaxRetries.Increment = 1

    # Label TCP Retry Delay
    $LabelTcpRetryDelay.Text = "Retry Delay ms:"
    $LabelTcpRetryDelay.Location = New-Object System.Drawing.Point(10, 170)
    $LabelTcpRetryDelay.AutoSize = $true

    # Create a TCP Retry Delay control
    $numericUpDownTcpRetryDelay.Location = New-Object System.Drawing.Point(10, 190)
    $numericUpDownTcpRetryDelay.Size = New-Object System.Drawing.Size(90,30)
    $numericUpDownTcpRetryDelay.Minimum = 50
    $numericUpDownTcpRetryDelay.Maximum = 10000
    $numericUpDownTcpRetryDelay.Value = $RetryDelayMs
    $numericUpDownTcpRetryDelay.Increment = 50

    # Label TCP Threads
    $LabelTcpThreads.Text = "TCP Threads:"
    $LabelTcpThreads.Location = New-Object System.Drawing.Point(10, 220)
    $LabelTcpThreads.AutoSize = $true

    # Create a TCP Threads control
    $numericUpDownTcpThreads.Location = New-Object System.Drawing.Point(10, 240)
    $numericUpDownTcpThreads.Size = New-Object System.Drawing.Size(90,30)
    $numericUpDownTcpThreads.Minimum = 5
    $numericUpDownTcpThreads.Maximum = 1000
    $numericUpDownTcpThreads.Value = $TCPThreads
    $numericUpDownTcpThreads.Increment = 5

    # Label
    $Labelsf = New-Object System.Windows.Forms.Label
    $Labelsf.Text = "Select file:"
    $Labelsf.Location = New-Object System.Drawing.Point(15, 170)
    $Labelsf.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
    $Labelsf.AutoSize = $true
    $Form.Controls.Add($Labelsf)

    # ComboBox
    $files = Get-ChildItem -Path $InitialDirectory -File | Where-Object { $_.Extension -eq ".txt" -or $_.Extension -eq ".csv" } | Select-Object -ExpandProperty Name
    $comboBox = New-Object System.Windows.Forms.ComboBox
    $comboBox.Location = New-Object System.Drawing.Point(15, 190)
    $comboBox.Size = New-Object System.Drawing.Size(130, 30)
    $comboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $comboBox.Anchor = 'Top, Left, Right'
    $comboBox.Items.AddRange($files)
    $form.Controls.Add($comboBox)

    # Button
    $Button = New-Object System.Windows.Forms.Button
    $Button.Location = New-Object System.Drawing.Point(25,240)
    $Button.Size = New-Object System.Drawing.Size(100,30)
    $Button.Text = "Run"
    $Button.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
    $Button.Add_Click({
        if ($comboBox.SelectedItem -notlike $null) {
                $script:selectedFile = Join-Path -Path $PrinterListPath -ChildPath $comboBox.SelectedItem
                $script:CsvBufferSize = $numericUpDownCsvBuf.Value
                $script:TimeoutMsUDP = $numericUpDownSnmp.Value
                $script:TimeoutMsTCP = $numericUpDownTcp.Value
                $script:MaxRetries = $numericUpDownTcpMaxRetries.Value
                $script:RetryDelayMs = $numericUpDownTcpRetryDelay.Value
                $script:TcpThreads = $numericUpDownTcpThreads.Value
                $form.Close()
        } else {
            [System.Windows.Forms.MessageBox]::Show("Please select a file.", "Warning", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        }
    })
    $form.Controls.Add($Button)

    $form.ShowDialog() | Out-Null

    return $script:selectedFile
}