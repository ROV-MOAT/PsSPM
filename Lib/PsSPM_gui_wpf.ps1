Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms

$xaml = @'
<Window 
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="PsSPM 0.3.4" SizeToContent="WidthAndHeight" ResizeMode="NoResize" WindowStartupLocation="CenterScreen" Topmost="True">
        
    <Grid x:Name="grid1" ShowGridLines="false">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"></RowDefinition>
            <RowDefinition Height="Auto"></RowDefinition>
        </Grid.RowDefinitions>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="Auto"></ColumnDefinition>
            <ColumnDefinition Width="Auto"></ColumnDefinition>
        </Grid.ColumnDefinitions>

        <GroupBox Grid.Column="1" Grid.Row="0" Header="TCP Settings" Padding="5">
        <Border CornerRadius="5" BorderBrush="Gray" BorderThickness="0" Padding="0" Margin="0">
            <StackPanel Orientation="Vertical" HorizontalAlignment="Center" Margin="0">
                <StackPanel Orientation="Horizontal">
                    <DockPanel VerticalAlignment="Center" LastChildFill="False" Width="135">
                        <Label Content="Timeout ms:" DockPanel.Dock="Left" Margin="0"/>
                        <TextBox x:Name="TCPTimeout" TextAlignment="Center" DockPanel.Dock="Right" MaxLength="5" Width="40" Height="20" Margin="0"/>
                    </DockPanel>
                </StackPanel>
                <StackPanel Orientation="Horizontal">
                    <DockPanel VerticalAlignment="Center" LastChildFill="False" Width="135">
                        <Label Content="Max Retries:" DockPanel.Dock="Left" Margin="0"/>
                        <TextBox x:Name="TCPMaxRetries" TextAlignment="Center" DockPanel.Dock="Right" MaxLength="2" Width="40" Height="20" Margin="0"/>
                    </DockPanel>
                </StackPanel>
                <StackPanel Orientation="Horizontal">
                    <DockPanel VerticalAlignment="Center" LastChildFill="False" Width="135">
                        <Label Content="Retry Delay ms:" DockPanel.Dock="Left" Margin="0"/>
                        <TextBox x:Name="TCPRetryDelay" TextAlignment="Center" DockPanel.Dock="Right" MaxLength="5" Width="40" Height="20" Margin="0"/>
                    </DockPanel>
                </StackPanel>
                <StackPanel Orientation="Horizontal">
                    <DockPanel VerticalAlignment="Center" LastChildFill="False" Width="135">
                        <Label Content="Threads:" DockPanel.Dock="Left" Margin="0"/>
                        <TextBox x:Name="xTCPThreads" TextAlignment="Center" DockPanel.Dock="Right" MaxLength="4" Width="40" Height="20" Margin="0"/>
                    </DockPanel>
                </StackPanel>
            </StackPanel>
            </Border>
        </GroupBox>

        <GroupBox Grid.Column="0" Grid.Row="1" Header="Device pool" Padding="5">
        <StackPanel>
            <Label Content="Enter IP range:" Margin="0"/>
            <StackPanel Name="IpStackPanel" Orientation="Horizontal" HorizontalAlignment="Center" VerticalAlignment="Center">
                <TextBox x:Name="Octet1" TextAlignment="Center" MaxLength="3" Width="26" Height="20" Margin="5"/>
                <TextBlock Text="." VerticalAlignment="Center" Margin="0"/>
                <TextBox x:Name="Octet2" TextAlignment="Center" MaxLength="3" Width="26" Height="20" Margin="5"/>
                <TextBlock Text="." VerticalAlignment="Center" Margin="0"/>
                <TextBox x:Name="Octet3" TextAlignment="Center" MaxLength="3" Width="26" Height="20" Margin="5"/>
                <TextBlock Text="." VerticalAlignment="Center" Margin="0"/>
                <TextBox x:Name="Octet4" TextAlignment="Center" MaxLength="3" Width="26" Height="20" Margin="5"/>
                <TextBlock Text="to" VerticalAlignment="Center" Margin="0"/>
                <TextBox x:Name="Octet5" TextAlignment="Center" MaxLength="3" Width="26" Height="20" Margin="5"/>
            </StackPanel>
            <StackPanel Orientation="Horizontal" Margin="0,10,0,0">
                    <Label Content="or select file:" Margin="0"/>
                    <ComboBox x:Name="FileSelectComboBox" Margin="0" Width="118" Height="20" IsEditable="True"/>
            </StackPanel>
        </StackPanel>
        </GroupBox>


        <StackPanel Grid.Column="0" Grid.Row="0">
        <GroupBox Header="Report" Padding="5">
        <StackPanel Orientation="Vertical" HorizontalAlignment="Center" Margin="0">
            <StackPanel Orientation="Horizontal">
                <CheckBox x:Name="HtmlCheckBox" Content="HTML" Margin="5,0,0,0"/>
                <CheckBox x:Name="CsvCheckBox" Content="CSV" Margin="5,0,0,0"/>
                <CheckBox x:Name="LogCheckBox" Content="Log" Margin="5,0,0,0"/>
            </StackPanel>
            <StackPanel Orientation="Horizontal" Margin="0,5,0,0">
                <Label Content="CSV buffer (line):" Margin="0"/>
                <TextBox x:Name="CSVBuffer" TextAlignment="Center" HorizontalAlignment="Left" MaxLength="4" Width="40" Height="20" Margin="0"/>
            </StackPanel>
            </StackPanel>
        </GroupBox>
        <GroupBox Header="SNMP Settings" Padding="5">
        <StackPanel Orientation="Vertical" HorizontalAlignment="Center" Margin="0">
            <StackPanel Orientation="Horizontal">
                <DockPanel VerticalAlignment="Center" LastChildFill="False" Width="115">
                    <Label Content="Timeout ms:" DockPanel.Dock="Left" Margin="0"/>
                    <TextBox x:Name="SNMPTimeout" TextAlignment="Center" DockPanel.Dock="Right" MaxLength="5" Width="40" Height="20" Margin="0"/>
                </DockPanel>
            </StackPanel>
        </StackPanel>
        </GroupBox>
        </StackPanel>
        <Button x:Name="RunButton" Grid.Column="1" Grid.Row="1" Content="Run" Width="50" Height="30" Padding="5" Margin="5"/>
    </Grid>
</Window>
'@
function Show-UserGUIXaml {
    param (
        [string]$Directory = $PSScriptRoot  # Default to current directory
    )

$xmlDoc = New-Object System.Xml.XmlDocument
$xmlDoc.LoadXml($xaml)
$reader = (New-Object System.Xml.XmlNodeReader($xmlDoc))
$form = [Windows.Markup.XamlReader]::Load($reader)

$HtmlCheckBox = $form.FindName("HtmlCheckBox")
$HtmlCheckBox.isChecked = $script:HtmlFileReport
$HtmlCheckBox.Add_Checked({ $script:HtmlFileReport = $true })
$HtmlCheckBox.Add_Unchecked({ $script:HtmlFileReport = $false })

$CsvCheckBox = $form.FindName("CsvCheckBox")
$CsvCheckBox.isChecked = $script:CsvFileReport
$CsvCheckBox.Add_Checked({ $script:CsvFileReport = $true })
$CsvCheckBox.Add_Unchecked({ $script:CsvFileReport = $false })

$LogCheckBox = $form.FindName("LogCheckBox")
$LogCheckBox.isChecked = $script:WriteLog
$LogCheckBox.Add_Checked({ $script:WriteLog = $true })
$LogCheckBox.Add_Unchecked({ $script:WriteLog = $false })

$CSVBuffer = $form.FindName("CSVBuffer")
$CSVBuffer.text = $CsvBufferSize

$SNMPTimeout = $form.FindName("SNMPTimeout")
$SNMPTimeout.text = $TimeoutMsUDP

$TCPTimeout = $form.FindName("TCPTimeout")
$TCPTimeout.text = $TimeoutMsTCP

$TCPRetryDelay = $form.FindName("TCPRetryDelay")
$TCPRetryDelay.text = $RetryDelayMs

$TCPMaxRetries = $form.FindName("TCPMaxRetries")
$TCPMaxRetries.text = $MaxRetries

$xTCPThreads = $form.FindName("xTCPThreads")
$xTCPThreads.text = $TCPThreads

$Octet1 = $form.FindName("Octet1")
$Octet2 = $form.FindName("Octet2")
$Octet3 = $form.FindName("Octet3")
$Octet4 = $form.FindName("Octet4")
$Octet5 = $form.FindName("Octet5")
$FileComboBox = $form.FindName("FileSelectComboBox")
$Button1 = $form.FindName("RunButton")

$IpNumericTextBoxPreviewTextInput = {
    param($sender, $e)

    $focusedElement = [System.Windows.Input.FocusManager]::GetFocusedElement($form)
    $MyStackPanel = $form.FindName("IpStackPanel")
    $currentElement = $MyStackPanel.Children | Where-Object { $_.Name -eq "$($focusedElement.Name)" }

    if ($currentElement) {
        $currentIndex = $MyStackPanel.Children.IndexOf($currentElement)
        # Allow digits
        if ([char]::IsDigit($e.Text)) {
        $e.Handled = $false
        }
        # Allow only one dot
        elseif ($e.Text -eq '.' -and -not $sender.Text.Contains('.')) {
            if (($currentIndex + 2) -lt $MyStackPanel.Children.Count) {
                $nextElement = $MyStackPanel.Children[$currentIndex + 2]
                $nextElement.Focus()
                #Write-Host "The next element is: $($nextElement.Name)"
            }
            $e.Handled = $true
        }
        # Disallow other characters
        else {
            $e.Handled = $true
        }
    }
}

$IpTextBoxTextChanged = {
    $focusedElement = [System.Windows.Input.FocusManager]::GetFocusedElement($form)
    $MyStackPanel = $form.FindName("IpStackPanel")
    $currentElement = $MyStackPanel.Children | Where-Object { $_.Name -eq "$($focusedElement.Name)" }

    if ($currentElement) {
        $currentIndex = $MyStackPanel.Children.IndexOf($currentElement)

        if ($currentElement.Name -match "Octet1") {
            if ($currentElement.Text -ne "" -and [int]$currentElement.Text -gt 223) {
                [System.Windows.Forms.MessageBox]::Show("Octet must be ≤ 223", "Error", "OK", "Error")
                $currentElement.Text = ""
                $currentElement.SelectAll()
            } else {
                $text = $currentElement.Text
                if ($text.Length -eq 3) {
                    if (($currentIndex + 2) -lt $MyStackPanel.Children.Count) {
                        $nextElement = $MyStackPanel.Children[$currentIndex + 2]
                        $nextElement.Focus()
                    }
                }
            }
        } else {
            if ($currentElement.Text -ne "" -and [int]$currentElement.Text -gt 255) {
                [System.Windows.Forms.MessageBox]::Show("Octet must be ≤ 255", "Error", "OK", "Error")
                $currentElement.Text = ""
                $currentElement.SelectAll()
            } else {
                $text = $currentElement.Text
                if ($text.Length -eq 3) {
                    if (($currentIndex + 2) -lt $MyStackPanel.Children.Count) {
                        $nextElement = $MyStackPanel.Children[$currentIndex + 2]
                        $nextElement.Focus()
                    }
                }
            }
        }
    }
}

function Update-ComboBox {
    param (
        [string]$CbDirectory = $PSScriptRoot  # Default to current directory
    )
    
    try {
        # Get all files in the directory
        $files = Get-ChildItem -Path $CbDirectory -File | Where-Object { $_.Extension -eq ".txt" -or $_.Extension -eq ".csv" } | Select-Object -ExpandProperty Name
        
        # Clear and populate the ComboBox
        $FileComboBox.Items.Clear()
        foreach ($file in $files) {
            $FileComboBox.Items.Add($file) | Out-Null
        }
    }
    catch {
        [System.Windows.MessageBox]::Show("Error accessing directory: $_", "Error")
    }
}

function Get-SimpleIPRange {
    param (
        [string]$BaseIP = "10.10.10",
        [int]$StartIP = 1,
        [int]$EndIP = 120
    )

    if ($StartIP -gt $EndIP) {
        $Octet5.Text = ""
        $Octet5.Focus()
        [System.Windows.Forms.MessageBox]::Show("Octet 5 must be ≤ Octet 4", "Error", "OK", "Error")
        return @()
    }

    $ipList = @()
    for ($i = $StartIP; $i -le $EndIP; $i++) {
        $ipList += [PSCustomObject]@{ Value = "$BaseIP.$i" }
    }

    return $ipList
}

$Octet1.Add_PreviewTextInput($IpNumericTextBoxPreviewTextInput)
$Octet2.Add_PreviewTextInput($IpNumericTextBoxPreviewTextInput)
$Octet3.Add_PreviewTextInput($IpNumericTextBoxPreviewTextInput)
$Octet4.Add_PreviewTextInput($IpNumericTextBoxPreviewTextInput)
$Octet5.Add_PreviewTextInput($IpNumericTextBoxPreviewTextInput)
$Octet1.Add_TextChanged($IpTextBoxTextChanged)
$Octet2.Add_TextChanged($IpTextBoxTextChanged)
$Octet3.Add_TextChanged($IpTextBoxTextChanged)
$Octet4.Add_TextChanged($IpTextBoxTextChanged)
$Octet5.Add_TextChanged($IpTextBoxTextChanged)

$Button1.Add_Click({
    if ($Octet5.Text -ne "") {
                if ( ([Convert]::ToInt32(($Octet4.Text).ToString())) -gt ([Convert]::ToInt32(($Octet5.Text).ToString())) ) {
                    $Octet5.Text = ""
                    $Octet5.Focus()
                    [System.Windows.Forms.MessageBox]::Show("Octet 5 must be ≥ Octet 4", "Error", "OK", "Error")
                } else {
                    $script:PrinterRange = Get-SimpleIPRange -BaseIP "$($Octet1.Text).$($Octet2.Text).$($Octet3.Text)" -StartIP $($Octet4.Text) -EndIP $($Octet5.Text)
                    $script:CsvBufferSize = $CSVBuffer.text
                    $script:TimeoutMsUDP = $SNMPTimeout.text
                    $script:TimeoutMsTCP = $TCPTimeout.text
                    $script:MaxRetries = $TCPMaxRetries.text
                    $script:RetryDelayMs = $TCPRetryDelay.text
                    $script:TcpThreads = $xTCPThreads.text
                    $form.Close()
                }
        }
    elseif ($FileComboBox.SelectedItem -notlike $null) {
                $script:selectedFile = Join-Path -Path $Directory -ChildPath $FileComboBox.SelectedItem
                $script:CsvBufferSize = $CSVBuffer.text
                $script:TimeoutMsUDP = $SNMPTimeout.text
                $script:TimeoutMsTCP = $TCPTimeout.text
                $script:MaxRetries = $TCPMaxRetries.text
                $script:RetryDelayMs = $TCPRetryDelay.text
                $script:TcpThreads = $xTCPThreads.text
                $form.Close()
        } else {
            [System.Windows.Forms.MessageBox]::Show("Please enter IP range or select file.", "Warning", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        }
})

$form.Add_Loaded({ Update-ComboBox -CbDirectory $Directory })

$form.ShowDialog() | Out-Null

if ($PrinterRange -notlike $null) { return $script:PrinterRange } else { return $script:selectedFile }
}