<#
.SYNOPSIS
    Toont alle printers in de huidige (Citrix) sessie in een overzichtelijke GUI.

.DESCRIPTION
    - Haalt alle printers op die zichtbaar zijn voor de ingelogde gebruiker in deze sessie
      (Citrix clientprinters, netwerkprinters, lokale en virtuele printers).
    - Toont sessie-informatie: gebruiker, clientnaam/IP en standaardprinter.
    - Opgemaakt in de huisstijl van het Streekziekenhuis Koningin Beatrix (SKB), logo is ingebouwd.
    - Zoeken en filteren op type, vernieuwen (handmatig of automatisch elke 30 seconden).
    - Acties: als standaard instellen en afdrukwachtrij openen.

.NOTES
    Versie:   26112301
    Formaat:  JJMMDDnn  (datum + volgnummer van die dag, 01 = eerste versie van de dag)

    Versiegeschiedenis:
    26112301  - Kolom "Taken" verwijderd, kolom "Opmerking" (Comment) toegevoegd naast "Locatie".
              - Zoeken doorzoekt nu ook de opmerking.
              - Verborgen foto verwijderd.
              - Versienummer toegevoegd (in de broncode en in de statusbalk).

    Starten:  powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Show-CitrixPrinters.ps1
    Zonder consolevenster: voeg -WindowStyle Hidden toe.
    Geen admin-rechten nodig. Werkt op Windows PowerShell 5.1 en PowerShell 7.
    Bewust zonder schaduw-effecten/animaties, zodat het ook in Citrix (software rendering) vlot blijft.
#>

#Requires -Version 5.1

# Versienummer: JJMMDDnn. Bij elke wijziging ophogen en de versiegeschiedenis hierboven bijwerken.
$ScriptVersion = '26112301'

# --- WPF vereist een STA-thread; indien nodig opnieuw starten met -STA -----------------------
if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') {
    $exe = (Get-Process -Id $PID).Path
    Start-Process -FilePath $exe -ArgumentList @('-NoProfile', '-STA', '-ExecutionPolicy', 'Bypass', '-File', "`"$PSCommandPath`"")
    return
}

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase

# =============================================================================================
#  Gegevens ophalen
# =============================================================================================

function Get-SessionInfo {
    $sessionId   = (Get-Process -Id $PID).SessionId
    $sessionName = $env:SESSIONNAME
    $clientName  = $env:CLIENTNAME
    $clientIp    = $null
    $isCitrix    = $false

    # Citrix registreert per sessie gegevens onder HKLM:\SOFTWARE\Citrix\Ica\Session\<id>
    $icaKey = "HKLM:\SOFTWARE\Citrix\Ica\Session\$sessionId\Connection"
    if (Test-Path $icaKey) {
        $isCitrix = $true
        $conn = Get-ItemProperty -Path $icaKey -ErrorAction SilentlyContinue
        if ($conn.ClientName)    { $clientName = $conn.ClientName }
        if ($conn.ClientAddress) { $clientIp   = $conn.ClientAddress }
    }
    if ($sessionName -like 'ICA*') { $isCitrix = $true }
    if (-not $isCitrix -and (Test-Path 'HKLM:\SOFTWARE\Citrix\VirtualDesktopAgent')) { $isCitrix = $true }

    [pscustomobject]@{
        User        = "$env:USERDOMAIN\$env:USERNAME"
        Computer    = $env:COMPUTERNAME
        SessionId   = $sessionId
        SessionName = if ($sessionName) { $sessionName } else { 'onbekend' }
        ClientName  = if ($clientName)  { $clientName }  else { '-' }
        ClientIp    = $clientIp
        IsCitrix    = $isCitrix
    }
}

function Get-SessionPrinters {
    $virtualPattern = 'Microsoft Print to PDF|Microsoft XPS|OneNote|^Fax$|Send To OneNote|PDF24|Adobe PDF'

    foreach ($p in (Get-CimInstance -ClassName Win32_Printer -ErrorAction Stop)) {
        $name   = [string]$p.Name
        $port   = [string]$p.PortName
        $driver = [string]$p.DriverName

        # --- Type bepalen ---
        if ($port -like 'CLIENT:*' -or $name -match '\((from|van) .+?\)' -or $name -match '\(redirected \d+\)' -or $name -match 'in (session|sessie) \d+') {
            $type = 'Citrix clientprinter'; $typeCat = 'Citrix'
        }
        elseif ($name -like '\\*' -or $p.Network) {
            $type = 'Netwerkprinter'; $typeCat = 'Network'
        }
        elseif ($driver -like 'Citrix Universal*' -or $name -like 'Citrix*') {
            $type = 'Citrix clientprinter'; $typeCat = 'Citrix'
        }
        elseif ($name -match $virtualPattern -or $driver -match $virtualPattern) {
            $type = 'Virtueel'; $typeCat = 'Virtual'
        }
        else {
            $type = 'Lokaal'; $typeCat = 'Local'
        }

        # --- Bron (client / printserver / deze server) ---
        $source = $env:COMPUTERNAME
        if ($name -match '\((?:from|van) (.+?)\)')      { $source = "Client: $($Matches[1])" }
        elseif ($name -match '^\\\\([^\\]+)\\')         { $source = $Matches[1] }
        elseif ($p.ServerName)                          { $source = ([string]$p.ServerName).TrimStart('\') }
        elseif ($typeCat -eq 'Citrix' -and $env:CLIENTNAME) { $source = "Client: $env:CLIENTNAME" }

        # --- Status ---
        $status = switch ([int]$p.PrinterStatus) {
            3 { 'Gereed' }
            4 { 'Bezig met afdrukken' }
            5 { 'Opwarmen' }
            6 { 'Gestopt' }
            7 { 'Offline' }
            1 { 'Overig' }
            default { 'Onbekend' }
        }
        if ($p.WorkOffline) { $status = 'Offline' }
        $statusCat = switch ($status) {
            { $_ -in 'Gereed', 'Bezig met afdrukken', 'Opwarmen' } { 'Ok'; break }
            { $_ -in 'Gestopt', 'Offline' }                         { 'Error'; break }
            default                                                 { 'Warning' }
        }

        [pscustomobject]@{
            IsDefault      = [bool]$p.Default
            Name           = $name
            Type           = $type
            TypeCategory   = $typeCat
            Status         = $status
            StatusCategory = $statusCat
            Driver         = $driver
            Port           = $port
            Source         = $source
            Location       = [string]$p.Location
            Comment        = (([string]$p.Comment) -replace '\s*\r?\n\s*', ' | ').Trim()
            Shared         = if ($p.Shared) { 'Ja' } else { 'Nee' }
            CimObject      = $p
        }
    }
}

# =============================================================================================
#  Logo SKB (ingebouwd als base64-PNG, zodat het script los te gebruiken is)
# =============================================================================================

$LogoBase64 = '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'

# =============================================================================================
#  GUI (XAML)
# =============================================================================================

[xml]$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Printeroverzicht - Streekziekenhuis Koningin Beatrix" Width="1240" Height="740" MinWidth="900" MinHeight="500"
        WindowStartupLocation="CenterScreen" FontFamily="Segoe UI" FontSize="13"
        Background="#F7F4F8" UseLayoutRounding="True" TextOptions.TextFormattingMode="Display">

  <Window.Resources>
    <!-- Knoppen -->
    <Style x:Key="Btn" TargetType="Button">
      <Setter Property="Background" Value="#FFFFFF"/>
      <Setter Property="Foreground" Value="#6B257C"/>
      <Setter Property="BorderBrush" Value="#CDB6D4"/>
      <Setter Property="Padding" Value="14,7"/>
      <Setter Property="Margin" Value="8,0,0,0"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="bd" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}"
                    BorderThickness="1" CornerRadius="6" Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="bd" Property="Opacity" Value="0.82"/>
              </Trigger>
              <Trigger Property="IsPressed" Value="True">
                <Setter TargetName="bd" Property="Opacity" Value="0.65"/>
              </Trigger>
              <Trigger Property="IsEnabled" Value="False">
                <Setter TargetName="bd" Property="Opacity" Value="0.4"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="PrimaryBtn" TargetType="Button" BasedOn="{StaticResource Btn}">
      <Setter Property="Background" Value="#6B257C"/>
      <Setter Property="BorderBrush" Value="#6B257C"/>
      <Setter Property="Foreground" Value="White"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
    </Style>

    <!-- Infokaarten -->
    <Style x:Key="Card" TargetType="Border">
      <Setter Property="Background" Value="White"/>
      <Setter Property="BorderBrush" Value="#E7DFEA"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="CornerRadius" Value="8"/>
      <Setter Property="Padding" Value="14,10"/>
    </Style>
    <Style x:Key="CardLabel" TargetType="TextBlock">
      <Setter Property="FontSize" Value="11"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Foreground" Value="#8A5A98"/>
      <Setter Property="Margin" Value="0,0,0,3"/>
    </Style>
    <Style x:Key="CardValue" TargetType="TextBlock">
      <Setter Property="FontSize" Value="14.5"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Foreground" Value="#111827"/>
      <Setter Property="TextTrimming" Value="CharacterEllipsis"/>
    </Style>

    <!-- DataGrid -->
    <Style TargetType="DataGridColumnHeader">
      <Setter Property="Background" Value="#F8F3FA"/>
      <Setter Property="Foreground" Value="#6B257C"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="FontSize" Value="12"/>
      <Setter Property="Padding" Value="10,9"/>
      <Setter Property="BorderBrush" Value="#E5E7EB"/>
      <Setter Property="BorderThickness" Value="0,0,0,1"/>
    </Style>
    <Style TargetType="DataGridCell">
      <Setter Property="Background" Value="Transparent"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Padding" Value="10,0"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="DataGridCell">
            <Border Background="{TemplateBinding Background}" Padding="{TemplateBinding Padding}">
              <ContentPresenter VerticalAlignment="Center"/>
            </Border>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
      <Style.Triggers>
        <Trigger Property="IsSelected" Value="True">
          <Setter Property="Background" Value="#EBDDF0"/>
          <Setter Property="Foreground" Value="#111827"/>
        </Trigger>
      </Style.Triggers>
    </Style>
    <Style TargetType="DataGridRow">
      <Style.Triggers>
        <Trigger Property="IsMouseOver" Value="True">
          <Setter Property="Background" Value="#FAF6FB"/>
        </Trigger>
      </Style.Triggers>
    </Style>
  </Window.Resources>

  <Grid>
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>

    <!-- Header in SKB-huisstijl -->
    <Border Grid.Row="0" Background="White" Padding="24,14" BorderBrush="#B8C248" BorderThickness="0,0,0,4">
      <DockPanel>
        <Image x:Name="imgLogo" DockPanel.Dock="Left" Height="62" VerticalAlignment="Center"
               RenderOptions.BitmapScalingMode="HighQuality"/>
        <Border x:Name="bdCitrix" DockPanel.Dock="Right" VerticalAlignment="Center" CornerRadius="14"
                Padding="12,5" Background="#F3E8F5">
          <TextBlock x:Name="txtCitrix" Foreground="#6B257C" FontWeight="SemiBold" Text="..."/>
        </Border>
        <Border DockPanel.Dock="Left" Width="1" Background="#E5DDE8" Margin="24,6,24,6"/>
        <StackPanel VerticalAlignment="Center">
          <TextBlock Text="Printeroverzicht" Foreground="#6B257C" FontSize="24" FontWeight="SemiBold"
                     HorizontalAlignment="Left"/>
          <TextBlock x:Name="txtSubtitle" Text="Alle printers die beschikbaar zijn in deze sessie"
                     Foreground="#6B7280" Margin="0,2,0,0"/>
        </StackPanel>
      </DockPanel>
    </Border>

    <!-- Infokaarten -->
    <UniformGrid Grid.Row="1" Columns="3" Margin="20,16,8,0">
      <Border Style="{StaticResource Card}" Margin="0,0,12,0">
        <StackPanel>
          <TextBlock Style="{StaticResource CardLabel}" Text="GEBRUIKER"/>
          <TextBlock x:Name="txtUser" Style="{StaticResource CardValue}"/>
          <TextBlock x:Name="txtServer" FontSize="11.5" Foreground="#6B7280" Margin="0,2,0,0"/>
        </StackPanel>
      </Border>
      <Border Style="{StaticResource Card}" Margin="0,0,12,0">
        <StackPanel>
          <TextBlock Style="{StaticResource CardLabel}" Text="CLIENT (EINDGEBRUIKERSAPPARAAT)"/>
          <TextBlock x:Name="txtClient" Style="{StaticResource CardValue}"/>
          <TextBlock x:Name="txtClientIp" FontSize="11.5" Foreground="#6B7280" Margin="0,2,0,0"/>
        </StackPanel>
      </Border>
      <Border Style="{StaticResource Card}" Margin="0,0,12,0">
        <StackPanel>
          <TextBlock Style="{StaticResource CardLabel}" Text="STANDAARDPRINTER"/>
          <TextBlock x:Name="txtDefault" Style="{StaticResource CardValue}"/>
          <TextBlock x:Name="txtCount" FontSize="11.5" Foreground="#6B7280" Margin="0,2,0,0"/>
        </StackPanel>
      </Border>
    </UniformGrid>

    <!-- Hoofdpaneel -->
    <Border Grid.Row="2" Style="{StaticResource Card}" Margin="20,16,20,0" Padding="0">
      <Grid>
        <Grid.RowDefinitions>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="*"/>
        </Grid.RowDefinitions>

        <!-- Werkbalk -->
        <DockPanel Grid.Row="0" Margin="14,12,14,12" LastChildFill="False">
          <Grid DockPanel.Dock="Left" Width="260">
            <TextBox x:Name="txtSearch" Padding="8,6" VerticalContentAlignment="Center"
                     BorderBrush="#CDB6D4" ToolTip="Zoek op naam, bron, locatie of opmerking"/>
            <TextBlock x:Name="txtPlaceholder" Text="Zoeken..." Foreground="#9CA3AF" Margin="11,0,0,0"
                       VerticalAlignment="Center" IsHitTestVisible="False"/>
          </Grid>
          <ComboBox x:Name="cmbType" DockPanel.Dock="Left" Width="180" Margin="8,0,0,0"
                    SelectedIndex="0" VerticalContentAlignment="Center" Padding="8,6">
            <ComboBoxItem Content="Alle types"/>
            <ComboBoxItem Content="Citrix clientprinter"/>
            <ComboBoxItem Content="Netwerkprinter"/>
            <ComboBoxItem Content="Lokaal"/>
            <ComboBoxItem Content="Virtueel"/>
          </ComboBox>

          <Button x:Name="btnQueue"   DockPanel.Dock="Right" Style="{StaticResource Btn}" Content="Wachtrij openen"/>
          <Button x:Name="btnDefault" DockPanel.Dock="Right" Style="{StaticResource Btn}" Content="Als standaard"/>
          <Button x:Name="btnRefresh" DockPanel.Dock="Right" Style="{StaticResource PrimaryBtn}" Content="Vernieuwen"/>
        </DockPanel>

        <!-- Printerlijst -->
        <DataGrid x:Name="grid" Grid.Row="1" AutoGenerateColumns="False" IsReadOnly="True"
                  CanUserAddRows="False" CanUserDeleteRows="False" SelectionMode="Single"
                  SelectionUnit="FullRow" HeadersVisibility="Column" GridLinesVisibility="Horizontal"
                  HorizontalGridLinesBrush="#EEF0F3" BorderThickness="0,1,0,0" BorderBrush="#E5E7EB"
                  Background="White" RowBackground="White" AlternatingRowBackground="#FCFAFD"
                  RowHeight="36" CanUserResizeRows="False">
          <DataGrid.ContextMenu>
            <ContextMenu>
              <MenuItem Header="Als standaardprinter instellen"/>
              <MenuItem Header="Afdrukwachtrij openen"/>
              <Separator/>
              <MenuItem Header="Printernaam kopieren"/>
            </ContextMenu>
          </DataGrid.ContextMenu>
          <DataGrid.Columns>

            <DataGridTemplateColumn Header="" Width="36" SortMemberPath="IsDefault">
              <DataGridTemplateColumn.CellTemplate>
                <DataTemplate>
                  <TextBlock x:Name="star" Text="&#x2605;" Foreground="#9AA634" FontSize="15"
                             HorizontalAlignment="Center" Visibility="Hidden" ToolTip="Standaardprinter"/>
                  <DataTemplate.Triggers>
                    <DataTrigger Binding="{Binding IsDefault}" Value="True">
                      <Setter TargetName="star" Property="Visibility" Value="Visible"/>
                    </DataTrigger>
                  </DataTemplate.Triggers>
                </DataTemplate>
              </DataGridTemplateColumn.CellTemplate>
            </DataGridTemplateColumn>

            <DataGridTemplateColumn Header="Printernaam" Width="2.4*" SortMemberPath="Name">
              <DataGridTemplateColumn.CellTemplate>
                <DataTemplate>
                  <TextBlock x:Name="nm" Text="{Binding Name}" TextTrimming="CharacterEllipsis"
                             ToolTip="{Binding Name}" Foreground="#111827"/>
                  <DataTemplate.Triggers>
                    <DataTrigger Binding="{Binding IsDefault}" Value="True">
                      <Setter TargetName="nm" Property="FontWeight" Value="SemiBold"/>
                    </DataTrigger>
                  </DataTemplate.Triggers>
                </DataTemplate>
              </DataGridTemplateColumn.CellTemplate>
            </DataGridTemplateColumn>

            <DataGridTemplateColumn Header="Type" Width="160" SortMemberPath="Type">
              <DataGridTemplateColumn.CellTemplate>
                <DataTemplate>
                  <Border x:Name="b" CornerRadius="10" Padding="9,2" HorizontalAlignment="Left" Background="#F3F4F6">
                    <TextBlock x:Name="t" Text="{Binding Type}" FontSize="11.5" FontWeight="SemiBold" Foreground="#374151"/>
                  </Border>
                  <DataTemplate.Triggers>
                    <DataTrigger Binding="{Binding TypeCategory}" Value="Citrix">
                      <Setter TargetName="b" Property="Background" Value="#F1E6F4"/>
                      <Setter TargetName="t" Property="Foreground" Value="#6B257C"/>
                    </DataTrigger>
                    <DataTrigger Binding="{Binding TypeCategory}" Value="Network">
                      <Setter TargetName="b" Property="Background" Value="#EFF2D3"/>
                      <Setter TargetName="t" Property="Foreground" Value="#55600F"/>
                    </DataTrigger>
                    <DataTrigger Binding="{Binding TypeCategory}" Value="Local">
                      <Setter TargetName="b" Property="Background" Value="#E9E1EC"/>
                      <Setter TargetName="t" Property="Foreground" Value="#4E2B5A"/>
                    </DataTrigger>
                  </DataTemplate.Triggers>
                </DataTemplate>
              </DataGridTemplateColumn.CellTemplate>
            </DataGridTemplateColumn>

            <DataGridTemplateColumn Header="Status" Width="165" SortMemberPath="Status">
              <DataGridTemplateColumn.CellTemplate>
                <DataTemplate>
                  <StackPanel Orientation="Horizontal">
                    <Ellipse x:Name="dot" Width="9" Height="9" Fill="#F59E0B" Margin="0,0,7,0" VerticalAlignment="Center"/>
                    <TextBlock Text="{Binding Status}" Foreground="#374151"/>
                  </StackPanel>
                  <DataTemplate.Triggers>
                    <DataTrigger Binding="{Binding StatusCategory}" Value="Ok">
                      <Setter TargetName="dot" Property="Fill" Value="#10B981"/>
                    </DataTrigger>
                    <DataTrigger Binding="{Binding StatusCategory}" Value="Error">
                      <Setter TargetName="dot" Property="Fill" Value="#EF4444"/>
                    </DataTrigger>
                  </DataTemplate.Triggers>
                </DataTemplate>
              </DataGridTemplateColumn.CellTemplate>
            </DataGridTemplateColumn>

            <DataGridTextColumn Header="Bron"    Binding="{Binding Source}"   Width="1.1*"/>
            <DataGridTextColumn Header="Locatie" Binding="{Binding Location}" Width="1.2*"/>
            <DataGridTemplateColumn Header="Opmerking" Width="1.4*" SortMemberPath="Comment">
              <DataGridTemplateColumn.CellTemplate>
                <DataTemplate>
                  <TextBlock Text="{Binding Comment}" TextTrimming="CharacterEllipsis" ToolTip="{Binding Comment}"/>
                </DataTemplate>
              </DataGridTemplateColumn.CellTemplate>
            </DataGridTemplateColumn>
          </DataGrid.Columns>
        </DataGrid>
      </Grid>
    </Border>

    <!-- Statusbalk -->
    <DockPanel Grid.Row="3" Margin="22,10,22,12">
      <CheckBox x:Name="chkAuto" DockPanel.Dock="Right" Content="Automatisch vernieuwen (30 s)"
                Foreground="#4B5563" VerticalAlignment="Center"/>
      <TextBlock x:Name="lblStatus" Foreground="#6B7280" VerticalAlignment="Center" Text="Gereed"/>
    </DockPanel>
  </Grid>
</Window>
'@

$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [Windows.Markup.XamlReader]::Load($reader)

$ui = @{}
foreach ($n in 'bdCitrix','txtCitrix','txtSubtitle','imgLogo','txtUser','txtServer',
               'txtClient','txtClientIp','txtDefault','txtCount','txtSearch','txtPlaceholder','cmbType',
               'btnRefresh','btnDefault','btnQueue','grid','chkAuto','lblStatus') {
    $ui[$n] = $window.FindName($n)
}
$grid = $ui.grid

$script:AllPrinters = @()

# =============================================================================================
#  Hulpfuncties voor de GUI
# =============================================================================================

function ConvertTo-Bitmap([string]$Base64) {
    $bmp = New-Object System.Windows.Media.Imaging.BitmapImage
    $bmp.BeginInit()
    $bmp.StreamSource = New-Object System.IO.MemoryStream (,[Convert]::FromBase64String($Base64))
    $bmp.CacheOption  = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
    $bmp.EndInit()
    $bmp.Freeze()
    return $bmp
}

function Show-Message([string]$Text, [string]$Icon = 'Information', [string]$Buttons = 'OK') {
    [System.Windows.MessageBox]::Show($window, $Text, 'Printeroverzicht', $Buttons, $Icon)
}

function Test-Contains([string]$Value, [string]$Query) {
    return ($Value -and $Value.IndexOf($Query, [System.StringComparison]::OrdinalIgnoreCase) -ge 0)
}

function Update-SessionPanel {
    $s = Get-SessionInfo
    $ui.txtUser.Text      = $s.User
    $ui.txtServer.Text    = "Server: $($s.Computer)"
    $ui.txtClient.Text    = $s.ClientName
    $ui.txtClientIp.Text  = if ($s.ClientIp) { "IP-adres: $($s.ClientIp)" } else { 'IP-adres: onbekend' }

    if ($s.IsCitrix) {
        $ui.txtCitrix.Text = 'Citrix-sessie gedetecteerd'
        $bc = [System.Windows.Media.BrushConverter]::new()
        $ui.bdCitrix.Background = $bc.ConvertFromString('#EFF2D3')
        $ui.txtCitrix.Foreground = $bc.ConvertFromString('#55600F')
    } else {
        $ui.txtCitrix.Text = 'Geen Citrix-sessie gedetecteerd'
    }
}

function Update-View {
    $q = $ui.txtSearch.Text.Trim()
    $t = [string]$ui.cmbType.SelectedItem.Content

    $rows = @($script:AllPrinters | Where-Object {
        ($t -eq 'Alle types' -or $_.Type -eq $t) -and
        ([string]::IsNullOrEmpty($q) -or
            (Test-Contains $_.Name $q) -or
            (Test-Contains $_.Source $q) -or (Test-Contains $_.Location $q) -or (Test-Contains $_.Comment $q))
    })
    $grid.ItemsSource = $rows

    $all = $script:AllPrinters
    $c = @($all | Where-Object TypeCategory -eq 'Citrix').Count
    $n = @($all | Where-Object TypeCategory -eq 'Network').Count
    $l = @($all | Where-Object TypeCategory -eq 'Local').Count
    $v = @($all | Where-Object TypeCategory -eq 'Virtual').Count
    $ui.txtCount.Text = "$($all.Count) printers: $c Citrix, $n netwerk, $l lokaal, $v virtueel"

    $shown = if ($rows.Count -ne $all.Count) { "$($rows.Count) van $($all.Count) printers getoond" } else { "$($all.Count) printers" }
    $ui.lblStatus.Text = "$shown  |  Laatst bijgewerkt: $(Get-Date -Format 'HH:mm:ss')  |  Versie $ScriptVersion"
    Update-Buttons
}

function Update-PrinterList([switch]$Quiet) {
    $selectedName = if ($grid.SelectedItem) { $grid.SelectedItem.Name } else { $null }

    $ui.lblStatus.Text = 'Printers ophalen...'
    $window.Cursor = [System.Windows.Input.Cursors]::Wait
    $window.Dispatcher.Invoke([action]{}, [System.Windows.Threading.DispatcherPriority]::Render)

    try {
        $order = @{ Citrix = 0; Network = 1; Local = 2; Virtual = 3 }
        $script:AllPrinters = @(Get-SessionPrinters |
            Sort-Object @{ Expression = { -not $_.IsDefault } }, @{ Expression = { $order[$_.TypeCategory] } }, Name)
    }
    catch {
        $script:AllPrinters = @()
        if (-not $Quiet) { Show-Message "Printers konden niet worden opgehaald:`n`n$($_.Exception.Message)" 'Error' }
    }
    finally {
        $window.Cursor = $null
    }

    $def = $script:AllPrinters | Where-Object IsDefault | Select-Object -First 1
    $ui.txtDefault.Text    = if ($def) { $def.Name } else { '(geen)' }
    $ui.txtDefault.ToolTip = $ui.txtDefault.Text

    Update-View

    if ($selectedName) {
        $match = @($grid.ItemsSource) | Where-Object Name -eq $selectedName | Select-Object -First 1
        if ($match) { $grid.SelectedItem = $match; $grid.ScrollIntoView($match) }
    }
}

function Update-Buttons {
    $sel = $null -ne $grid.SelectedItem
    $ui.btnDefault.IsEnabled = $sel -and -not $grid.SelectedItem.IsDefault
    $ui.btnQueue.IsEnabled   = $sel
}

function Set-SelectedAsDefault {
    $p = $grid.SelectedItem
    if (-not $p) { return }
    try {
        $r = Invoke-CimMethod -InputObject $p.CimObject -MethodName SetDefaultPrinter -ErrorAction Stop
        if ($r.ReturnValue -ne 0) { throw "WMI gaf foutcode $($r.ReturnValue) terug." }
        Update-PrinterList
        $ui.lblStatus.Text = "Standaardprinter ingesteld op: $($p.Name)"
    }
    catch { Show-Message "Kon de standaardprinter niet instellen:`n`n$($_.Exception.Message)" 'Error' }
}

function Open-SelectedQueue {
    $p = $grid.SelectedItem
    if (-not $p) { return }
    Start-Process -FilePath 'rundll32.exe' -ArgumentList "printui.dll,PrintUIEntry /o /n `"$($p.Name)`""
}

# =============================================================================================
#  Events
# =============================================================================================

try { $ui.imgLogo.Source = ConvertTo-Bitmap $LogoBase64 } catch { $ui.imgLogo.Visibility = 'Collapsed' }

$ui.btnRefresh.Add_Click({ Update-SessionPanel; Update-PrinterList })
$ui.btnDefault.Add_Click({ Set-SelectedAsDefault })
$ui.btnQueue.Add_Click({ Open-SelectedQueue })

$ui.txtSearch.Add_TextChanged({
    $ui.txtPlaceholder.Visibility = if ($ui.txtSearch.Text) { 'Collapsed' } else { 'Visible' }
    Update-View
})
$ui.cmbType.Add_SelectionChanged({ if ($window.IsLoaded) { Update-View } })

$grid.Add_SelectionChanged({ Update-Buttons })
$grid.Add_MouseDoubleClick({ if ($grid.SelectedItem) { Open-SelectedQueue } })

$cm = $grid.ContextMenu
$cm.Items[0].Add_Click({ Set-SelectedAsDefault })
$cm.Items[1].Add_Click({ Open-SelectedQueue })
$cm.Items[3].Add_Click({ if ($grid.SelectedItem) { [System.Windows.Clipboard]::SetText($grid.SelectedItem.Name) } })
$cm.Add_Opened({
    $sel = $null -ne $grid.SelectedItem
    foreach ($i in 0, 1, 3) { $cm.Items[$i].IsEnabled = $sel }
    if ($sel -and $grid.SelectedItem.IsDefault) { $cm.Items[0].IsEnabled = $false }
})

# Automatisch vernieuwen (handig omdat Citrix clientprinters soms pas na het inloggen verschijnen)
$timer = New-Object System.Windows.Threading.DispatcherTimer
$timer.Interval = [TimeSpan]::FromSeconds(30)
$timer.Add_Tick({ Update-PrinterList -Quiet })
$ui.chkAuto.Add_Checked({ $timer.Start() })
$ui.chkAuto.Add_Unchecked({ $timer.Stop() })

$window.Add_KeyDown({
    param($s, $e)
    if ($e.Key -eq 'F5') { Update-SessionPanel; Update-PrinterList }
    if ($e.Key -eq 'F' -and [System.Windows.Input.Keyboard]::Modifiers -eq 'Control') { $ui.txtSearch.Focus() | Out-Null }
})

$window.Add_ContentRendered({ Update-SessionPanel; Update-PrinterList })
$window.Add_Closed({ $timer.Stop() })

[void]$window.ShowDialog()
