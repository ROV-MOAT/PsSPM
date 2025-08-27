Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms

$xaml = @'
<Window 
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="PsSPM 0.3.5b" SizeToContent="WidthAndHeight" ResizeMode="NoResize" WindowStartupLocation="CenterScreen" Topmost="True" FontFamily="Segoe UI" FontSize="13">
        
    <Grid x:Name="grid1" ShowGridLines="false" Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"></RowDefinition>
            <RowDefinition Height="Auto"></RowDefinition>
        </Grid.RowDefinitions>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="Auto"></ColumnDefinition>
            <ColumnDefinition Width="Auto"></ColumnDefinition>
        </Grid.ColumnDefinitions>

        <StackPanel Grid.Column="0" Grid.Row="0" Margin="0,0,5,0">
            <GroupBox Header="Report">
                <StackPanel Orientation="Vertical" HorizontalAlignment="Center" Margin="0">
                    <StackPanel Orientation="Horizontal">
                        <CheckBox x:Name="HtmlCheckBox" Content="HTML" Margin="5"/>
                        <CheckBox x:Name="CsvCheckBox" Content="CSV" Margin="5"/>
                        <CheckBox x:Name="LogCheckBox" Content="Log" Margin="5"/>
                    </StackPanel>
                    <StackPanel Orientation="Horizontal" Margin="0,5,0,0">
                        <DockPanel VerticalAlignment="Center" LastChildFill="False" Width="163">
                            <Label Content="CSV buffer (line):" DockPanel.Dock="Left" Margin="0"/>
                            <TextBox x:Name="CSVBuffer" TextAlignment="Center" DockPanel.Dock="Right" HorizontalAlignment="Left" MaxLength="4" Width="45" Height="20" Margin="0"/>
                        </DockPanel>
                    </StackPanel>
                </StackPanel>
            </GroupBox>
            <GroupBox Header="SNMP Settings" Margin="0,30,0,0">
                <StackPanel Orientation="Vertical" HorizontalAlignment="Center" Margin="0">
                    <StackPanel Orientation="Horizontal">
                        <DockPanel VerticalAlignment="Center" LastChildFill="False" Width="160">
                            <Label Content="Timeout (ms):" DockPanel.Dock="Left" Margin="0"/>
                            <TextBox x:Name="SNMPTimeout" TextAlignment="Center" DockPanel.Dock="Right" MaxLength="5" Width="45" Height="20" Margin="0"/>
                        </DockPanel>
                    </StackPanel>
                </StackPanel>
            </GroupBox>
        </StackPanel>

        <GroupBox Grid.Column="1" Grid.Row="0" Header="TCP Settings">
        <Border CornerRadius="5" BorderBrush="Gray" BorderThickness="0" Padding="0" Margin="0">
            <StackPanel Orientation="Vertical" HorizontalAlignment="Center" Margin="0,0,7,3">
                <StackPanel Orientation="Horizontal">
                    <DockPanel VerticalAlignment="Center" LastChildFill="False" Width="160">
                        <Label Content="TCP Port:" DockPanel.Dock="Left" Margin="0"/>
                        <TextBox x:Name="TCPPortSet" TextAlignment="Center" DockPanel.Dock="Right" MaxLength="5" Width="45" Height="20" Margin="0"/>
                    </DockPanel>
                </StackPanel>
                <StackPanel Orientation="Horizontal">
                    <DockPanel VerticalAlignment="Center" LastChildFill="False" Width="160">
                        <Label Content="Timeout (ms):" DockPanel.Dock="Left" Margin="0"/>
                        <TextBox x:Name="TCPTimeout" TextAlignment="Center" DockPanel.Dock="Right" MaxLength="5" Width="45" Height="20" Margin="0"/>
                    </DockPanel>
                </StackPanel>
                <StackPanel Orientation="Horizontal">
                    <DockPanel VerticalAlignment="Center" LastChildFill="False" Width="160">
                        <Label Content="Max Retries:" DockPanel.Dock="Left" Margin="0"/>
                        <TextBox x:Name="TCPMaxRetries" TextAlignment="Center" DockPanel.Dock="Right" MaxLength="2" Width="45" Height="20" Margin="0"/>
                    </DockPanel>
                </StackPanel>
                <StackPanel Orientation="Horizontal">
                    <DockPanel VerticalAlignment="Center" LastChildFill="False" Width="160">
                        <Label Content="Retry Delay (ms):" DockPanel.Dock="Left" Margin="0"/>
                        <TextBox x:Name="TCPRetryDelay" TextAlignment="Center" DockPanel.Dock="Right" MaxLength="5" Width="45" Height="20" Margin="0"/>
                    </DockPanel>
                </StackPanel>
                <StackPanel Orientation="Horizontal">
                    <DockPanel VerticalAlignment="Center" LastChildFill="False" Width="160">
                        <Label Content="Threads:" DockPanel.Dock="Left" Margin="0"/>
                        <TextBox x:Name="xTCPThreads" TextAlignment="Center" DockPanel.Dock="Right" MaxLength="4" Width="45" Height="20" Margin="0"/>
                    </DockPanel>
                </StackPanel>
            </StackPanel>
            </Border>
        </GroupBox>

        <GroupBox Grid.Column="0" Grid.Row="1" Header="Device pool" Margin="0,0,5,0">
            <StackPanel>
                <Label Content="Enter IP range:" Margin="0"/>
                <StackPanel Name="IpStackPanel" Orientation="Horizontal" HorizontalAlignment="Center" VerticalAlignment="Center">
                    <TextBox x:Name="Octet1" TextAlignment="Center" MaxLength="3" Width="30" Height="20" Margin="5"/>
                    <TextBlock Text="." VerticalAlignment="Center" Margin="0"/>
                    <TextBox x:Name="Octet2" TextAlignment="Center" MaxLength="3" Width="30" Height="20" Margin="5"/>
                    <TextBlock Text="." VerticalAlignment="Center" Margin="0"/>
                    <TextBox x:Name="Octet3" TextAlignment="Center" MaxLength="3" Width="30" Height="20" Margin="5"/>
                    <TextBlock Text="." VerticalAlignment="Center" Margin="0"/>
                    <TextBox x:Name="Octet4" TextAlignment="Center" MaxLength="3" Width="30" Height="20" Margin="5"/>
                    <TextBlock Text="to" VerticalAlignment="Center" Margin="0"/>
                    <TextBox x:Name="Octet5" TextAlignment="Center" MaxLength="3" Width="30" Height="20" Margin="5"/>
                </StackPanel>
                <StackPanel Orientation="Horizontal" Margin="0,10,0,0">
                    <Label Content="or select file:" Margin="0"/>
                    <ComboBox x:Name="FileSelectComboBox" Margin="0" Width="133" Height="25" IsEditable="True"/>
                </StackPanel>
            </StackPanel>
        </GroupBox>

        <Button x:Name="RunButton" Grid.Column="1" Grid.Row="1" Content="Run" Width="80" Height="30" Padding="5" Margin="5"/>
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

    $TCPPortSet = $form.FindName("TCPPortSet")
    $TCPPortSet.text = $TCPPort

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

    function Update-ComboBox {
        param (
            [string]$CbDirectory = $PSScriptRoot  # Default to current directory
        )
    
        try {
            # Get all files in the directory
            $files = Get-ChildItem -Path $CbDirectory -File | Where-Object { $_.Extension -eq ".txt" -or $_.Extension -eq ".csv" } | Select-Object -ExpandProperty Name
        
            # Clear and populate the ComboBox
            $FileComboBox.Items.Clear()
            foreach ($file in $files) { $FileComboBox.Items.Add($file) | Out-Null }
        } catch { [System.Windows.MessageBox]::Show("Error accessing directory: $_", "Error") }
    }

    $TextBoxDigInputHandler = {
        param($sender, $e)
    
        # Only numbers allowed
        if ([char]::IsDigit($e.Text)) {
            $e.Handled = $false
            return
        }
    
        # Block all other characters
        $e.Handled = $true
    }

    $TextBoxPreviewInputHandler = {
            param($sender, $e)
    
            # Only numbers allowed
            if ([char]::IsDigit($e.Text)) {
                $e.Handled = $false
                return
            }
    
            # Processing a point (move to the next field)
            if ($e.Text -eq '.' -or $e.Text -eq ',') {
                $e.Handled = $true
                $currentIndex = [array]::IndexOf($octetBoxes, $sender)
                if ($currentIndex -lt 4) {
                    $octetBoxes[$currentIndex + 1].Focus()
                    $octetBoxes[$currentIndex + 1].SelectAll()
                }
                return
            }
    
            # Block all other characters
            $e.Handled = $true
        }

        $TextBoxBackspaceHandler = {
            param($sender, $e)
    
            if ($e.Key -eq [System.Windows.Input.Key]::Back) {
                $currentIndex = [array]::IndexOf($octetBoxes, $sender)
        
                if ($sender.Text.Length -eq 0 -and $currentIndex -gt 0) {
                    $octetBoxes[$currentIndex - 1].Focus()
                    $octetBoxes[$currentIndex - 1].CaretIndex = $octetBoxes[$currentIndex - 1].Text.Length
                    $e.Handled = $true
                }
            }
        }
        $SettingBoxes = @($CSVBuffer, $SNMPTimeout, $TCPPortSet, $TCPTimeout, $TCPRetryDelay, $TCPMaxRetries, $xTCPThreads)

        # Registering handlers
        foreach ($Settingbox in $SettingBoxes) {
            # 1. Paste interception
            [System.Windows.Input.CommandManager]::AddPreviewExecutedHandler($Settingbox, {
                param($sender, $e)

                if ($e.Command -eq [System.Windows.Input.ApplicationCommands]::Paste) {
                    $clipboardText = [System.Windows.Clipboard]::GetText()
        
                    # Checking format and values
                    if ($clipboardText -match '^(\d{1,5})$') {
                        $sender.SelectAll()
                        $e.Handled = $false
                        return
                    }
                    $e.Handled = $true
                }
            })

            $Settingbox.Add_PreviewTextInput($TextBoxDigInputHandler)

            $Settingbox.Add_TextChanged({
                param($sender, $e)
    
                # Check only for filled fields
                if ($sender.Text -match '^\d+$') {
                    $value = [int]$sender.Text
                
                    # For the first octet
                    if ($sender.Text -ne "" -and $sender.Name -eq "TCPPortSet") {
                        if ($value -lt 1 -or $value -gt 65535) {
                            [System.Windows.MessageBox]::Show("The Port must be not 0 and ≤ 65535", "Error")
                            $sender.Text = ""
                            $sender.Focus()
                            return
                        }
                    } elseif ($value -gt 99999) {
                        [System.Windows.MessageBox]::Show("Value must be ≤ 99999", "Error")
                        $sender.Text = ""
                        $sender.Focus()
                        return
                    }
                }            
            })
        }

        # All TextBox for octets
        $octetBoxes = @($Octet1, $Octet2, $Octet3, $Octet4, $Octet5)

        # Registering handlers
        foreach ($box in $octetBoxes) {
            # 1. Paste interception
            [System.Windows.Input.CommandManager]::AddPreviewExecutedHandler($box, {
            param($sender, $e)
            if ($e.Command -eq [System.Windows.Input.ApplicationCommands]::Paste) {
                $clipboardText = [System.Windows.Clipboard]::GetText()
        
                # Checking format and values
                if ($clipboardText -match '^(\d{1,3})\.(\d{1,3})\.(\d{1,3})\.(\d{1,3})$') {
                    $octets = $clipboardText.Split('.')

                    # Validate first octet (≤223)
                    if ([int]$octets[0] -lt 1 -or [int]$octets[0] -gt 223) {
                        [System.Windows.MessageBox]::Show("The first octet must be not 0 and ≤ 223", "Error")
                        $e.Handled = $true
                        return
                    }
            
                    # Validate remaining octets (≤255)
                    for ($i = 1; $i -lt 4; $i++) {
                        if ([int]$octets[$i] -gt 255) {
                            [System.Windows.MessageBox]::Show("Octet $($i+1) must be ≤ 255", "Error")
                            $e.Handled = $true
                            return
                        }
                    }
            
                    # Filling fields
                    $Octet1.Text = $octets[0]
                    $Octet2.Text = $octets[1]
                    $Octet3.Text = $octets[2]
                    $Octet4.Text = $octets[3]
                    $Octet5.Focus()
                }
                $e.Handled = $true
            }
        })
    
        #2. Go to point + validate numbers (PreviewTextInput)
        $box.Add_PreviewTextInput($TextBoxPreviewInputHandler)
        # Backspace
        $box.Add_PreviewKeyDown($TextBoxBackspaceHandler)
    
        # 3. Automatic transition at 3 digits (TextChanged)
        $box.Add_TextChanged({
            param($sender, $e)
    
            # Check only for filled fields
            if ($sender.Text -match '^\d+$') {
                $value = [int]$sender.Text
                
                # For the first octet
                if ($sender.Text -ne "" -and $sender.Name -eq "Octet1") {
                    if ($value -lt 1 -or $value -gt 223) {
                        [System.Windows.MessageBox]::Show("The first octet must be not 0 and ≤ 223", "Error")
                        $sender.Text = ""
                        $sender.Focus()
                        return
                    }
                } elseif ($value -gt 255) {
                    [System.Windows.MessageBox]::Show("Octet must be ≤ 255", "Error")
                    $sender.Text = ""
                    $sender.Focus()
                    return
                }
            }            
    
            # Auto-transition at 3 digits
            if ($sender.Text.Length -eq 3 -and $sender.Text -match '^\d+$') {
                $currentIndex = [array]::IndexOf($octetBoxes, $sender)
                if ($currentIndex -lt 4) {
                    $octetBoxes[$currentIndex + 1].Focus()
                    $octetBoxes[$currentIndex + 1].SelectAll()
                }
            }
        })
    }

    $Button1.Add_Click({
        if ($Octet5.Text -ne "" -and $Octet4.Text -ne "") {
            if ([int]$Octet4.Text -gt [int]$Octet5.Text) {
                [System.Windows.Forms.MessageBox]::Show("The ending address must be greater than the starting address.", "Error", "OK", "Error")
                $Octet5.Focus()
                return
            } else {
                # Create an array of IP address ranges
                $BaseIP = "$($Octet1.Text).$($Octet2.Text).$($Octet3.Text)"
                $ipList = @()

                for ($i = [int]$Octet4.Text; $i -le [int]$Octet5.Text; $i++) { $ipList += [PSCustomObject]@{ Value = "$BaseIP.$i" } }
                $script:PrinterRange = $ipList
            }
        } elseif ($FileComboBox.SelectedItem -notlike $null -and $FileComboBox.SelectedItem -ne "") {
            $script:selectedFile = Join-Path -Path $Directory -ChildPath $FileComboBox.SelectedItem
        } else { [System.Windows.Forms.MessageBox]::Show("Enter IP address or select a file.", "Warning", "OK", "Warning")
            return
        }

        # Save settings
        $script:CsvBufferSize = $CSVBuffer.Text
        $script:TimeoutMsUDP = $SNMPTimeout.Text
        $script:TCPPort = $TCPPortSet.Text
        $script:TimeoutMsTCP = $TCPTimeout.Text
        $script:MaxRetries = $TCPMaxRetries.Text
        $script:RetryDelayMs = $TCPRetryDelay.Text
        $script:TcpThreads = $xTCPThreads.Text

        $form.Close()
    })

    $form.Add_Loaded({ Update-ComboBox -CbDirectory $Directory })

    $form.ShowDialog() | Out-Null

    if ($PrinterRange -notlike $null) { return $script:PrinterRange } else { return $script:selectedFile }
}