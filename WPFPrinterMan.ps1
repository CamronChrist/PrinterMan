<#
    @author Camron Christ
            Camron.Christ@hcahealthcare.com

    This program allows a user to view and manage remote computers printers,
    printer drivers, and printer ports. Built with Powershell's
    PrinterManagement Module and PrintUI.exe.
#>

# Main window markup
$xaml = @'
<Window 
 xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" 
 xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" 
 Title="PrinterMan" SizeToContent="Manual" Height="500" Width="882">
    <Grid Name="gridMain">
        <Grid.RowDefinitions>
            <RowDefinition Height="40" />
            <RowDefinition Height="*" />
        </Grid.RowDefinitions>
        <TabControl Grid.Row="1" Name="tabControl">
            <TabItem Header="Printers" Name="tabItemPrinters">
                <Grid Background="#FFE5E5E5">
                    <ListView Name="listViewPrinters">
                        <ListView.ContextMenu>
                            <ContextMenu Name="contextMenuPrinters">
                                <MenuItem Name="menuItemPrinterAdd" Header="Add...">
                                    <MenuItem Name="menuItemPrinterAddQuickPrinter" Header="Printer (Quick)"/>
                                    <MenuItem Name="menuItemPrinterAddPrinter" Header="Printer"/>
                                    <MenuItem Name="menuItemPrinterAddDriver" Header="Driver"/>
                                </MenuItem>
                                <Separator/>
                                <MenuItem Name="menuItemPrinterPrintQueue" Header="See What's Printing"/>
                                <MenuItem Name="menuItemPrinterPrintTestPage" Header="Send Test Page"/>
                                <MenuItem Name="menuItemPrinterClearPrintJobs" Header="Clear Print Jobs"/>
                                <Separator/>
                                <MenuItem Name="menuItemPrinterPrinterProperties" Header="Printer Properties"/>
                                <MenuItem Name="menuItemPrinterRemovePrinter" Header="Remove Printer" InputGestureText="Del"/>
                            </ContextMenu>
                        </ListView.ContextMenu>
                        <ListView.View>
                            <GridView>
                                <GridViewColumn Width="200" Header="Name" DisplayMemberBinding="{Binding Name}"/>
                                <GridViewColumn Width="400" Header="DriverName" DisplayMemberBinding="{Binding Drivername}"/>
                                <GridViewColumn Width="100" Header="PortName">
                                    <GridViewColumn.CellTemplate>
                                        <DataTemplate>
                                            <TextBlock Text="{Binding PortName}"/>
                                        </DataTemplate>
                                    </GridViewColumn.CellTemplate>
                                </GridViewColumn>
                                <GridViewColumn Width="150" Header="Shared" DisplayMemberBinding="{Binding Shared}"/>
                            </GridView>
                        </ListView.View>
                    </ListView>
                </Grid>
            </TabItem>
            <TabItem Header="Drivers" Name="tabItemDrivers">
                <Grid Background="#FFE5E5E5">
                    <ListView Name="listViewDrivers">
                        <ListView.ContextMenu>
                            <ContextMenu Name="contextMenuDrivers">
                                <MenuItem Name="menuItemDriverAdd" Header="Add...">
                                    <MenuItem Name="menuItemDriverAddPrinter" Header="Printer"/>
                                    <MenuItem Name="menuItemDriverAddDriver" Header="Driver"/>
                                </MenuItem>
                                <Separator/>
                                <MenuItem Name="menuItemDriverRemoveDriver" Header="Remove Driver" InputGestureText="Del"/>
                            </ContextMenu>
                        </ListView.ContextMenu>
                        <ListView.View>
                            <GridView>
                                <GridViewColumn Width="400" Header="Name" DisplayMemberBinding="{Binding Name}"/>
                                <GridViewColumn Width="200" Header="Manufacturer" DisplayMemberBinding="{Binding Manufacturer}"/>
                            </GridView>
                        </ListView.View>
                    </ListView>
                </Grid>
            </TabItem>
            <TabItem Header="Ports" Name="tabItemPorts">
                <Grid Background="#FFE5E5E5">
                    <ListView Name="listViewPorts">
                        <ListView.ContextMenu>
                            <ContextMenu Name="contextMenuPorts">
                                <MenuItem Name="menuItemPortAdd" Header="Add...">
                                    <MenuItem Name="menuItemPortAddPrinter" Header="Printer"/>
                                    <MenuItem Name="menuItemPortAddDriver" Header="Driver"/>
                                </MenuItem>
                                <Separator/>
                                <MenuItem Name="menuItemPortRemovePort" Header="Remove Port" InputGestureText="Del"/>
                            </ContextMenu>
                        </ListView.ContextMenu>
                        <ListView.View>
                            <GridView>
                                <GridViewColumn Width="200" Header="Name" DisplayMemberBinding="{Binding Name}"/>
                                <GridViewColumn Width="200" Header="Host Address" DisplayMemberBinding="{Binding PrinterHostAddress}"/>
                                <GridViewColumn Width="200" Header="Description" DisplayMemberBinding="{Binding Description}"/>
                            </GridView>
                        </ListView.View>
                    </ListView>
                </Grid>
            </TabItem>
        </TabControl>

        <TextBox Name="textBoxComputerName" CharacterCasing="Upper" HorizontalAlignment="Left" Height="23" Margin="111,12,0,0" TextWrapping="Wrap" Text="TextBox" VerticalAlignment="Top" Width="167" FontSize="14"/>
        <Label Content="Computer Name:" Margin="5,9,0,0" VerticalAlignment="Top" HorizontalAlignment="Left" Width="101"/>
        <Button x:Name="btnLoad" Content="Load" HorizontalAlignment="Left" Margin="283,12,0,0" VerticalAlignment="Top" Width="74" Height="23"/>
        <Button x:Name="btnRestartSpooler" Content="Restart Spooler" HorizontalAlignment="Left" Margin="362,12,0,0" VerticalAlignment="Top" Width="107" Height="23"/>
        <Button x:Name="btnPrintServerProperties" Content="Server Properties" HorizontalAlignment="Left" Margin="474,12,0,0" VerticalAlignment="Top" Width="120" Height="23"/>
    </Grid>
</Window>

'@
# Quick install popup window markup
$xaml1 = @'
    <Window
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:TwoWindowTest"
        mc:Ignorable="d"
        Height="250" Width="400">
    <Grid Name="MainGrid" Margin="8,8">

        <Grid.RowDefinitions>
            <RowDefinition Height="24"/>
            <RowDefinition Height="24"/>
            <RowDefinition Height="24"/>
            <RowDefinition Height="*"/>
        </Grid.RowDefinitions>

        <Grid Grid.Row="0">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="3*"/>
                <ColumnDefinition Width="3*"/>
                <ColumnDefinition Width="3*"/>
                <ColumnDefinition Width="1*"/>
            </Grid.ColumnDefinitions>
            <Label Content="Printer Name" Margin="0,0,6,0" Grid.Column="0"/>
            <Label Content="Driver" Margin="0,0,6,0" Grid.Column="1"/>
            <ComboBox Name="comboBoxPrinterConnection" Grid.Column="2" Margin="0,0,6,0" SelectedIndex="0" />
        </Grid>

        <Grid Margin="0,0,0,0" Grid.Row="1">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="3*"/>
                <ColumnDefinition Width="3*"/>
                <ColumnDefinition Width="3*"/>
                <ColumnDefinition Width="1*"/>
            </Grid.ColumnDefinitions>
            <TextBox Grid.Column="0" x:Name="textBoxPrinterName" ToolTip="Printer Name" Margin="0,0,6,0"/>
            <ComboBox Grid.Column="1" x:Name="comboBoxDrivers" ToolTip="Driver" Margin="0,0,6,0"/>
            <TextBox Grid.Column="2" x:Name="textBoxPrinterConnection" ToolTip="IP Address / Server Connection" Margin="0,0,6,0"/>
            <Button Grid.Column="3" x:Name="buttonInstallPrinter" Content="Install" ToolTip="Install Printer"/>
        </Grid>

        <Grid Margin="0,0,0,0" Grid.Row="2">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="6*"/>
                <ColumnDefinition Width="1*"/>
                <ColumnDefinition Width="2*"/>
            </Grid.ColumnDefinitions>
            <Label Name="labelConsole" Content="Console:" Grid.Column="0" Margin="0,0,6,0"/>
        </Grid>

        <Grid Margin="0,0,0,0" Grid.Row="3">
            <TextBox x:Name="textBoxConsole" VerticalScrollBarVisibility="Auto" IsReadOnly="True" FontFamily="Lucida Console"/>
        </Grid>
    </Grid>
</Window>

'@
function Convert-XAMLtoWindow{
    param
    (
        [Parameter(Mandatory=$true)]
        [string]
        $XAML
    )
    
    Add-Type -AssemblyName PresentationFramework
    
    $reader = [XML.XMLReader]::Create([IO.StringReader]$XAML)
    $result = [Windows.Markup.XAMLReader]::Load($reader)
    $reader.Close()
    $reader = [XML.XMLReader]::Create([IO.StringReader]$XAML)
    while ($reader.Read())
    {
        $name=$reader.GetAttribute('Name')
        if (!$name) { $name=$reader.GetAttribute('x:Name') }
        if($name)
        {$result | Add-Member NoteProperty -Name $name -Value $result.FindName($name) -Force}
    }
    $reader.Close()
    $result
}

function Show-WPFWindow{
    param
    (
        [Parameter(Mandatory)]
        [Windows.Window]
        $Window
    )
    
    $result = $null
    $null = $window.Dispatcher.InvokeAsync{
        $result = $window.ShowDialog()
        Set-Variable -Name result -Value $result -Scope 1
    }.Wait()
    $result
}

# Checks if computer name is online then loads printers/drivers/ports
function Invoke-Update{
    $global:ComputerName = $window.textBoxComputerName.Text
    $window.Title = "$($global:ComputerName)  -  PrinterMan"
    if(Test-Connection -ComputerName $global:ComputerName -Count 2 -Quiet)
    {
        $window.listViewPrinters.ItemsSource = @(Get-Printer -ComputerName $global:ComputerName | Where-Object -Property DeviceType -Like 'print' | Sort-Object)
        $window.listViewDrivers.ItemsSource = @(Get-PrinterDriver -ComputerName $global:ComputerName | Sort-Object)
        $window.listViewPorts.ItemsSource = @(Get-PrinterPort -ComputerName $global:ComputerName | Sort-Object)
    }
    else
    {
        # if computer is not found, display error
        $tempObj = New-Object -TypeName PSObject
        $tempObj | Add-Member -MemberType NoteProperty -Name Name -Value "Computer not responding. `nCheck to make sure the computer `nis online"
        $window.listViewPrinters.ItemsSource = @($tempObj)
        $window.listViewDrivers.ItemsSource = @($tempObj)
        $window.listViewPorts.ItemsSource = @($tempObj)
    }

    # Subscribe-PortClickEvents

}

# Opens the add printer wizard for remote computer
function Add-Printer_Click{
    rundll32 printui.dll,PrintUIEntry /im /c\\$global:ComputerName | Wait-Process
    Invoke-Update
}

# Opens the add driver wizard for remote computer
function Add-Driver_Click{
    rundll32 printui.dll,PrintUIEntry /id /c\\$global:ComputerName | Wait-Process
    Invoke-Update
}

# Opens the print queue popup of selected printer
function View-PrintQueue{
    $printerCol = $window.listViewPrinters.SelectedItems
    foreach($p in $printerCol) {
        rundll32 printui.dll,PrintUIEntry /o /n\\$global:ComputerName\$($p.Name)
    }
}

# Sends test print to selected printers
function Send-TestPage{
    # Split into two options because you can't use psexec as system account on your own computer
    $printerCol = $window.listViewPrinters.SelectedItems
    if($global:ComputerName -like $env:COMPUTERNAME)
    {
        foreach($p in $printerCol)
        {
            rundll32 printui.dll,PrintUIEntry /k /n\\$global:ComputerName\$($p.name)
        }
    }
    else
    {
        foreach($p in $printerCol)
        {
            psexec /s \\$global:ComputerName rundll32 printui.dll,PrintUIEntry /k /n$($p.name)
        }
    }
}

# Clears print jobs of selected printers
function Clear-PrintQueue{
    $printerCol = $window.listViewPrinters.SelectedItems
    $jobs = @($printerCol | foreach {$_ | Get-PrintJob})
    if($jobs.count -gt 0)
    {
        $jobs | Out-GridView -Title "Select Jobs to delete" -OutputMode Multiple | Remove-PrintJob
    } else {
        [System.Windows.MessageBox]::Show("No print jobs to display", "PrinterMan",[System.Windows.MessageBoxButton]::OK)
    }
}

# Opens printer Porperties of selected printer
function Open-PrinterProperties{
    $printerCol = $window.listViewPrinters.SelectedItems
    foreach($p in $printerCol) {
        rundll32 printui.dll,PrintUIEntry /p /n\\$global:ComputerName\$($p.Name)
    }
}

# Opens print server properties of remote computer
function Open-PrintServerProperties{
        rundll32 printui.dll,PrintUIEntry /s /c\\$global:ComputerName
}

# Removes selected printers. Errors not shown.
function Remove-Printer_Click{
    $numSelectedItems = $window.listViewPrinters.SelectedItems.Count
    if($numSelectedItems -gt 0 -and [System.Windows.MessageBox]::Show("Delete selected printers?", "PrinterMan",[System.Windows.MessageBoxButton]::OKCancel) -eq "OK")
    {
        foreach($i in $window.listViewPrinters.SelectedItems)
        {
            Remove-Printer -InputObject $i
        
        }
    }
    
    Invoke-Update
}

# Removes selected drivers. Errors not shown.
function Remove-Driver_Click{
    $numSelectedItems = $window.listViewDrivers.SelectedItems.Count
    if($numSelectedItems -gt 0 -and [System.Windows.MessageBox]::Show("Delete selected drivers?", "PrinterMan",[System.Windows.MessageBoxButton]::OKCancel) -eq "OK")
    {
        foreach($i in $window.listViewDrivers.SelectedItems)
        {
            Remove-PrinterDriver -InputObject $i
        }
        Invoke-Update
    }
}

# Removes selected ports. Errors not shown
function Remove-Port_Click{
    $numSelectedItems = $window.listViewPorts.SelectedItems.Count
    if($numSelectedItems -gt 0 -and [System.Windows.MessageBox]::Show("Delete selected ports?", "PrinterMan",[System.Windows.MessageBoxButton]::OKCancel) -eq "OK")
    {
        foreach($i in $window.listViewPorts.SelectedItems)
        {
            Remove-PrinterPort -InputObject $i
        }
        Invoke-Update
    }
}

# Currently using doubleclick to open up the browser to the selected IP printers

# Subscribes the TextBlocks of the IP address to a click event
# function Subscribe-PortClickEvents{
#     #$ports = $window.listViewPrinters.Items.
#     $portAddress = ($window.listViewPorts.Items | ? {$_.Name -like $portName}).printerhostaddress
#     if( $portAddress -match '^\b(25[0-5]|2[0-4][0-9]|1[0-9][0-9]|[1-9]?[0-9])\.(25[0-5]|2[0-4][0-9]|1[0-9][0-9]|[1-9]?[0-9])\.(25[0-5]|2[0-4][0-9]|1[0-9][0-9]|[1-9]?[0-9])\.(25[0-5]|2[0-4][0-9]|1[0-9][0-9]|[1-9]?[0-9])\b$')
#     {
#         Start-Process "http://$portAddress"
#     }
# }

# Fills in Computer TextBox with local computer name and then Updates the list.
function Invoke-Initialize{
    $Global:PrinterConnections = @(
    "TCP/IP",
    "Local"
    )
    $global:ComputerName = $env:COMPUTERNAME
    $window.textBoxComputerName.Text = $global:ComputerName
    Invoke-Update
}

# Printer Add Pop-Up Window functions
function Validate-Printer{
    
    $isValidIPAddress = ($window1.comboBoxPrinterConnection.Text -like $Global:PrinterConnections[0] -and $window1.textBoxPrinterConnection.Text.Trim() -match
        '^\b(25[0-5]|2[0-4][0-9]|1[0-9][0-9]|[1-9]?[0-9])\.(25[0-5]|2[0-4][0-9]|1[0-9][0-9]|[1-9]?[0-9])\.(25[0-5]|2[0-4][0-9]|1[0-9][0-9]|[1-9]?[0-9])\.(25[0-5]|2[0-4][0-9]|1[0-9][0-9]|[1-9]?[0-9])\b$'
    )
    $isValidPrinterName = ($window1.textBoxPrinterName.Text -ne "")
    $isValidPrintServerConnection = ($window1.comboBoxPrinterConnection.Text -like $Global:PrinterConnections[1] -and
        $window1.textBoxPrinterConnection.Text.Trim() -match '^\\\\\w+\\\w+$'
    )
    
    if($isValidPrinterName -and ($isValidIPAddress -or $isValidPrintServerConnection))
    {
        if($window1.textBoxPrinterConnection.BorderBrush -like "#FFFF0000") {
            $window1.textBoxPrinterConnection.BorderBrush = "#FFABADB3"
        }
        # Creates Printer object
        $printer = New-Object -TypeName PSObject
        $printer | Add-Member -MemberType NoteProperty -Name Name -Value $window1.textBoxPrinterName.Text.Trim()
        $printer | Add-Member -MemberType NoteProperty -Name Driver -Value $window1.comboBoxDrivers.Text.Trim()
        $printer | Add-Member -MemberType NoteProperty -Name Port -Value $window1.textBoxPrinterConnection.Text.Trim()
        $printer | Add-Member -MemberType NoteProperty -Name PortType -Value $(if($isValidIPAddress){$Global:PrinterConnections[0]}else{$Global:PrinterConnections[1]})
        return $printer
    }

    if(-not $isValidPrinterName) {
        $window1.textBoxPrinterName.BorderBrush = "#FFFF0000"
    }

    if(-not ($isValidIPAddress -or $isValidPrintServerConnection))
    {
        $window1.textBoxPrinterConnection.BorderBrush = "red"
    }

    return $null
}
function Write-Console($str){
    $window1.textBoxConsole.Text += "$str`n"
    $window1.textboxConsole.ScrollToEnd()
}
function Invoke-ShowAddPrinterWindow{

    $window1 = Convert-XAMLtoWindow -XAML $xaml1
    $window1.Title = "Add Printer:  " + $window.Title

    $window1.comboBoxPrinterConnection.ItemsSource = $Global:PrinterConnections
    $window1.comboBoxDrivers.ItemsSource = (Get-PrinterDriver).Name | Sort-Object
    $window1.comboBoxDrivers.SelectedIndex = 0

    # Printer Add Window Events
    $window1.buttonInstallPrinter.add_click({
        Install-Printer
    })
    $window1.textBoxPrinterName.add_gotKeyboardFocus({
        if ($window1.textBoxPrinterName.BorderBrush -like "#FFFF0000") {
            $window1.textBoxPrinterName.BorderBrush = "#FFABADB3"
        }
    })
    $window1.textBoxPrinterName.add_keyDown({
        if ($args[1].key -eq 'Enter'){
            Install-Printer
        }
    })
    $window1.textBoxPrinterConnection.add_gotKeyboardFocus({
        if ($window1.textBoxPrinterConnection.BorderBrush -like "#FFFF0000") {
            $window1.textBoxPrinterConnection.BorderBrush = "#FFABADB3"
        }
    })
    $window1.textBoxPrinterConnection.add_keyDown({
        if ($args[1].key -eq 'Enter'){
            Install-Printer
        }
    })
    $window1.add_Closing({
        Install-Printer -stop
    })


    $null = Show-WPFWindow -Window $window1
}
function Install-Printer([switch] $stop){
    
    # Check if printer is valid
    $computer = $Global:ComputerName
    $printer = Validate-Printer

    # Check if buttonInstallPrinter is set to install or stop installation
    if($window1.buttonInstallPrinter.Content -like "Stop" -or $stop){
        $window1.buttonInstallPrinter.Content = "Install"
        $window1.buttonInstallPrinter.ToolTip = "Install printer"
        $global:aSync.PowerShell.Stop()
        Write-Console "Cancelled install"
        return
    } elseif($null -eq $printer) {
        return
    }

    $window1.buttonInstallPrinter.Content = "Stop"
    $window1.buttonInstallPrinter.ToolTip = "Cancel install"
    

    # Create Thread
    $InstallRunspace = [runspacefactory]::CreateRunspace()
    $InstallRunspace.ApartmentState = "STA"
    $InstallRunspace.ThreadOptions = "ReuseThread"
    $InstallRunspace.Name = "InstallRunspace"
    $InstallRunspace.Open()

    Write-Console("Initiating install...")

    $InstallRunspace.SessionStateProxy.SetVariable("window1",$window1) 
    $InstallRunspace.SessionStateProxy.SetVariable("printer",$printer)
    $InstallRunspace.SessionStateProxy.SetVariable("computer",$computer)
    $InstallRunspace.SessionStateProxy.SetVariable("PrinterConnections",$global:PrinterConnections)
    $global:aSync = "" | Select-Object PowerShell,Runspace,Job
    $global:aSync.Job = $InstallRunspace.Name

    $code = {
    function Write-Console($str){
        $window1.Dispatcher.invoke(
        [action]{
                    $window1.textBoxConsole.Text += "$str`n"
                    $window1.textboxConsole.ScrollToEnd()
        })
    }
    [string[]]$failed = @()
    foreach($c in $computer){
        if($null -ne $c -and $c.Trim() -ne ""){
            Write-Console "`n$c`: Testing Connection..."
            if(Test-Connection -ComputerName $c -Count 2 -Quiet -ErrorAction stop){
                foreach($p in $printer){
                    
                    Write-Console "`n  $($p.Name)`: Beginning installation..."

                    # Test whether the print spooler is reachable
                    Write-Console "`tChecking Print Spooler..."
                    try
                    {
                        Get-Printer -ComputerName $c -ErrorAction Stop > $null 2>&1
                    }
                    catch
                    {
                        reg add "\\$c\HKLM\Software\Policies\Microsoft\Windows NT\Printers" /v RegisterSpoolerRemoteRpcEndPoint /t REG_DWORD /d 1 /f
                        Get-Service -ComputerName $c -Name Spooler | Restart-Service
                        Write-Console "$($c): Waiting for Spooler to restart..."
                        (Get-Service -ComputerName $c -Name Spooler).('Running', '00:00:05')
                    }

                    # If port does not exist, add port
                    Write-Console "`tChecking Port..."
                    try
                    {
                        Get-PrinterPort -ComputerName $c -Name $p.Port -ErrorAction Stop > $null 2>&1
                    }
                    catch
                    {
                        Write-Console "`tAdding Port: `'$($p.PortType) $($p.Port)`'"
                        if($p.portType -like $printerConnections[0]) {
                            Add-PrinterPort -ComputerName $c -Name $p.Port -PrinterHostAddress $p.Port
                        }
                        elseif($p.portType -like $printerConnections[1]) {
                            Add-PrinterPort -ComputerName $c -Name $p.Port
                        }
                        else {
                            Write-Console "`tPort: `'$($p.PortType) $($p.Port)`' installation failed..."
                        }
                    }

                    # install printer
                    Write-Console "`tAdding Printer..."
                    #Add-Printer -ComputerName $c -Name $p.Name -PortName $p.Port -DriverName $p.Driver
                    rundll32 printui.dll,PrintUIEntry /if /c\\$c /b $($p.Name) /r $($p.Port) /m $($p.Driver) /q | Wait-Process
                    
                    # Verify successful installation
                    Write-Console "`tVerifying successful install..."
                    try
                    {
                        Get-Printer -ComputerName $c -Name $p.Name -ErrorAction Stop > $null 2>&1
                        Write-Console ("`t`'$($p.Name)`' Installed")
                    }
                    catch
                    {
                        Write-Console("`t`'$($p.Name)`' installation not successful")
                        $failed += "$($c): $($p.Name)"
                    }
                }
            
            }
            # PC not on network
            else
            {
                Write-Console( "`t$($c): Not found..." )
                $failed += "$($c): Not found..."
            }
        }
    }
    Write-Console("`n----------Finished------------")
    if($failed.Count -gt 0){
        Write-Console("")
        Write-Console("------Failed Device List------")
        Write-Console($failed -join "`n")
    }
    $window1.Dispatcher.invoke(
        [action]{
            $window1.buttonInstallPrinter.Content = "Install"
            $window1.buttonInstallPrinter.ToolTip = "Install printer to computer"
    })
    }

    $PSinstance = [powershell]::Create().AddScript($code)
    $PSinstance.Runspace = $InstallRunspace
    $global:aSync.PowerShell = $PSinstance
    $global:aSync.Runspace = $PSinstance.BeginInvoke()
    #$jobs.Add($global:aSync)
}



$window = Convert-XAMLtoWindow -XAML $xaml

#----------------------Events---------------------------

$window.textBoxComputerName.add_KeyDown({
    if ($args[1].key -eq 'Return')
    {
        Invoke-Update
    }
})
$window.listViewPrinters.add_KeyDown({
    if ($args[1].key -eq 'Delete')
    {
        Write-Host "Delete key pressed in listViewPrinters"
        Remove-Printer_Click
    }
})
$window.listViewPrinters.add_MouseDoubleClick({
    $portName = $window.listViewPrinters.SelectedItem.PortName
    $portAddress = ($window.listViewPorts.Items | ? {$_.Name -like $portName}).printerhostaddress
    if( $portAddress -match '^\b(25[0-5]|2[0-4][0-9]|1[0-9][0-9]|[1-9]?[0-9])\.(25[0-5]|2[0-4][0-9]|1[0-9][0-9]|[1-9]?[0-9])\.(25[0-5]|2[0-4][0-9]|1[0-9][0-9]|[1-9]?[0-9])\.(25[0-5]|2[0-4][0-9]|1[0-9][0-9]|[1-9]?[0-9])\b$')
    {
        Start-Process "http://$portAddress"
    }
})
$window.listViewDrivers.add_KeyDown({
    if ($args[1].key -eq 'Delete')
    {
        Write-Host "Delete key pressed in listViewDrivers"
        Remove-Driver_Click
    }
})
$window.listViewPorts.add_KeyDown({
    if ($args[1].key -eq 'Delete')
    {
        Write-Host "Delete key pressed in listViewPorts"
        Remove-Port_Click
    }
})

# Adds a reg value to allow PrinterMangement Module to view spooler information
# Then it restarts the spooler.
$window.btnRestartSpooler.add_Click({
    reg add "\\$global:ComputerName\HKLM\Software\Policies\Microsoft\Windows NT\Printers" /v RegisterSpoolerRemoteRpcEndPoint /t REG_DWORD /d 1 /f
    Get-Service -ComputerName $global:ComputerName -Name Spooler | Restart-Service
    (Get-Service -ComputerName $global:ComputerName -Name Spooler).('Running', '00:00:05')
    Invoke-Update
})
$window.btnLoad.add_Click({
    Invoke-Update
})
$window.btnPrintServerProperties.add_Click({
    Invoke-Update
    Open-PrintServerProperties
})

<# The following closed events trigger the appropriate external window to open based on
   the value that is set in the click event for each contextMenu.
   It circumvents a graphical error of contextMenu not closing properly.
#>
$window.contextMenuPrinters.add_Closed({
    $temp = $Global:clickedMenuItem
    $Global:clickedMenuItem = ""
    # if Add > Printer (Quick) is clicked
    if($temp -like $args[0].Items[0].Items[0].Name)
    {
        Invoke-ShowAddPrinterWindow
        Invoke-Update
    }
    # if Add > Printer is clicked
    if($temp -like $args[0].Items[0].Items[1].Name)
    {
        Add-Printer_Click
    }
    # if Add > Driver is clicked
    elseif($temp -like $args[0].Items[0].Items[2].Name)
    {
        Add-Driver_Click
    }
    # if 'Clear Print Jobs' is clicked
    elseif($temp -like $args[0].Items[4].Name)
    {
        Clear-PrintQueue
    }
    # if 'Send Test Page' is clicked
    elseif($temp -like $args[0].Items[3].Name)
    {
        Send-TestPage
    }
})
$window.contextMenuDrivers.add_Closed({
    $temp = $Global:clickedMenuItem
    $Global:clickedMenuItem = ""
    if($temp -eq $args[0].Items[0].Items[0].Name)
    {
        Add-Printer_Click
    }
    elseif($temp -eq $args[0].Items[0].Items[1].Name)
    {
        Add-Driver_Click
    }
})
$window.contextMenuPorts.add_Closed({
    $temp = $Global:clickedMenuItem
    $Global:clickedMenuItem = ""
    if($temp -eq $args[0].Items[0].Items[0].Name)
    {
        Add-Printer_Click
    }
    elseif($temp -eq $args[0].Items[0].Items[1].Name)
    {
        Add-Driver_Click
    }
})

# Printer Tab: Context menu click events
$window.menuItemPrinterAddPrinter.add_Click({
    $Global:clickedMenuItem = $args[0].Name
})
$window.menuItemPrinterAddQuickPrinter.add_Click({
    $Global:clickedMenuItem = $args[0].Name
})
$window.menuItemPrinterAddDriver.add_Click({
    $Global:clickedMenuItem = $args[0].Name
})
$window.menuItemPrinterPrintQueue.add_Click({
    View-PrintQueue
})
$window.menuItemPrinterPrintTestPage.add_Click({
    $Global:clickedMenuItem = $args[0].Name
    #Send-TestPage
})
$window.menuItemPrinterClearPrintJobs.add_Click({
    $Global:clickedMenuItem = $args[0].Name
})
$window.menuItemPrinterPrinterProperties.add_Click({
    Open-PrinterProperties
})
$window.menuItemPrinterRemovePrinter.add_Click({
    Remove-Printer_Click
})

# Driver Tab: Context menu click events
$window.menuItemDriverAddPrinter.add_Click({
    $Global:clickedMenuItem = $args[0].Name
})
$window.menuItemDriverAddDriver.add_Click({
    $Global:clickedMenuItem = $args[0].Name
})
$window.menuItemDriverRemoveDriver.add_Click({
    Remove-Driver_Click
})

#Port Tab: Context menu click events
$window.menuItemPortAddPrinter.add_Click({
    $Global:clickedMenuItem = $args[0].Name
})
$window.menuItemPortAddDriver.add_Click({
    $Global:clickedMenuItem = $args[0].Name
})
$window.menuItemPortRemovePort.add_Click({
    Remove-Port_Click
})

#------------------------Code-----------------------------

Invoke-Initialize

$null = Show-WPFWindow -Window $window