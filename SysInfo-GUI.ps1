<#
.SYNOPSIS
    SysInfo-GUI - WPF graphical front-end for the Sysinfo-Scraper tool.

.DESCRIPTION
    Modern WPF GUI that dot-sources SysInfo-Core.ps1 and presents system
    information in a tabbed, dark-themed interface.  Supports background
    scanning with real-time progress, hardware image display, and export
    to TXT / CSV formats.

.AUTHOR
    @13X

.VERSION
    2.0

.NOTES
    Requires PowerShell 5.1+ and WPF assemblies (Windows only).
#>

#Requires -Version 5.1

# ── Load WPF assemblies ────────────────────────────────────────────────────────
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms   # for FolderBrowserDialog

# ── Dot-source the core module ─────────────────────────────────────────────────
. "$PSScriptRoot\SysInfo-Core.ps1"

# ── XAML Definition ────────────────────────────────────────────────────────────

[xml]$xaml = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Sysinfo Scraper v2.0 - GUI"
    Width="1050" Height="720"
    MinWidth="900" MinHeight="650"
    WindowStartupLocation="CenterScreen"
    Background="#1E1E1E">

    <Window.Resources>
        <!-- Accent button style -->
        <Style x:Key="AccentButton" TargetType="Button">
            <Setter Property="Background" Value="#0078D4"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="FontFamily" Value="Segoe UI"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Padding" Value="16,8"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="border" Background="{TemplateBinding Background}"
                                CornerRadius="4" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#1A8AD4"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#005A9E"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="border" Property="Background" Value="#555555"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Standard button style -->
        <Style x:Key="StdButton" TargetType="Button">
            <Setter Property="Background" Value="#3C3C3C"/>
            <Setter Property="Foreground" Value="#CCCCCC"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="FontFamily" Value="Segoe UI"/>
            <Setter Property="Padding" Value="12,6"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="border" Background="{TemplateBinding Background}"
                                CornerRadius="3" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#505050"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#2A2A2A"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="border" Property="Background" Value="#2A2A2A"/>
                                <Setter Property="Foreground" Value="#666666"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- DataGrid style -->
        <Style TargetType="DataGrid">
            <Setter Property="Background" Value="#252525"/>
            <Setter Property="Foreground" Value="#CCCCCC"/>
            <Setter Property="BorderBrush" Value="#3C3C3C"/>
            <Setter Property="RowBackground" Value="#252525"/>
            <Setter Property="AlternatingRowBackground" Value="#2D2D2D"/>
            <Setter Property="GridLinesVisibility" Value="None"/>
            <Setter Property="HeadersVisibility" Value="Column"/>
            <Setter Property="FontFamily" Value="Segoe UI"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="IsReadOnly" Value="True"/>
            <Setter Property="AutoGenerateColumns" Value="True"/>
            <Setter Property="CanUserAddRows" Value="False"/>
            <Setter Property="SelectionMode" Value="Single"/>
        </Style>

        <Style TargetType="DataGridColumnHeader">
            <Setter Property="Background" Value="#333333"/>
            <Setter Property="Foreground" Value="#E0E0E0"/>
            <Setter Property="Padding" Value="8,4"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="BorderBrush" Value="#444444"/>
            <Setter Property="BorderThickness" Value="0,0,1,1"/>
        </Style>

        <Style TargetType="DataGridCell">
            <Setter Property="Padding" Value="6,3"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Foreground" Value="#CCCCCC"/>
            <Style.Triggers>
                <Trigger Property="IsSelected" Value="True">
                    <Setter Property="Background" Value="#0078D4"/>
                    <Setter Property="Foreground" Value="White"/>
                </Trigger>
            </Style.Triggers>
        </Style>
    </Window.Resources>

    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Title bar area -->
        <Border Grid.Row="0" Background="#252525" Padding="16,10">
            <StackPanel>
                <TextBlock Text="Sysinfo Scraper v2.0 - GUI"
                           FontSize="20" FontWeight="Bold" FontFamily="Segoe UI"
                           Foreground="#E0E0E0"/>
                <TextBlock Text="by @13X" FontSize="11" FontFamily="Segoe UI"
                           Foreground="#888888" Margin="0,2,0,0"/>
            </StackPanel>
        </Border>

        <!-- Main content -->
        <Grid Grid.Row="1" Margin="8">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="220"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="200"/>
            </Grid.ColumnDefinitions>

            <!-- Left panel: hardware image -->
            <Border Grid.Column="0" Background="#252525" CornerRadius="4" Margin="4" Padding="10">
                <StackPanel VerticalAlignment="Top">
                    <TextBlock Text="Hardware" FontSize="14" FontWeight="SemiBold"
                               Foreground="#E0E0E0" FontFamily="Segoe UI"
                               HorizontalAlignment="Center" Margin="0,0,0,8"/>
                    <Border Background="#1E1E1E" CornerRadius="4"
                            Width="200" Height="200" Margin="0,0,0,8">
                        <Grid>
                            <TextBlock x:Name="ImagePlaceholder"
                                       Text="No image available"
                                       Foreground="#666666" FontFamily="Segoe UI"
                                       FontSize="12" FontStyle="Italic"
                                       HorizontalAlignment="Center" VerticalAlignment="Center"/>
                            <Image x:Name="HardwareImage"
                                   Width="190" Height="190"
                                   Stretch="Uniform" Visibility="Collapsed"/>
                        </Grid>
                    </Border>
                    <TextBlock x:Name="MakeLabel" Text="Make: Unknown"
                               Foreground="#AAAAAA" FontFamily="Segoe UI" FontSize="12"
                               HorizontalAlignment="Center" Margin="0,2,0,0"
                               TextWrapping="Wrap" TextAlignment="Center"/>
                    <TextBlock x:Name="ModelLabel" Text="Model: Unknown"
                               Foreground="#AAAAAA" FontFamily="Segoe UI" FontSize="12"
                               HorizontalAlignment="Center" Margin="0,2,0,0"
                               TextWrapping="Wrap" TextAlignment="Center"/>
                </StackPanel>
            </Border>

            <!-- Center panel: tab control -->
            <TabControl x:Name="MainTabs" Grid.Column="1" Margin="4"
                        Background="#252525" BorderBrush="#3C3C3C"
                        FontFamily="Segoe UI" FontSize="12"
                        Foreground="#CCCCCC">

                <!-- System tab -->
                <TabItem Header="System">
                    <ScrollViewer VerticalScrollBarVisibility="Auto" Margin="8">
                        <Grid x:Name="SystemGrid">
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="140"/>
                                <ColumnDefinition Width="*"/>
                            </Grid.ColumnDefinitions>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/><RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/><RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/><RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/><RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                            </Grid.RowDefinitions>
                            <TextBlock Grid.Row="0" Grid.Column="0" Text="Computer Name:" Foreground="#AAAAAA" Margin="0,4"/>
                            <TextBlock Grid.Row="0" Grid.Column="1" x:Name="valComputerName" Text="Unknown" Foreground="#E0E0E0" Margin="8,4"/>
                            <TextBlock Grid.Row="1" Grid.Column="0" Text="Current User:" Foreground="#AAAAAA" Margin="0,4"/>
                            <TextBlock Grid.Row="1" Grid.Column="1" x:Name="valCurrentUser" Text="Unknown" Foreground="#E0E0E0" Margin="8,4"/>
                            <TextBlock Grid.Row="2" Grid.Column="0" Text="Domain:" Foreground="#AAAAAA" Margin="0,4"/>
                            <TextBlock Grid.Row="2" Grid.Column="1" x:Name="valDomain" Text="Unknown" Foreground="#E0E0E0" Margin="8,4"/>
                            <TextBlock Grid.Row="3" Grid.Column="0" Text="Manufacturer:" Foreground="#AAAAAA" Margin="0,4"/>
                            <TextBlock Grid.Row="3" Grid.Column="1" x:Name="valMake" Text="Unknown" Foreground="#E0E0E0" Margin="8,4"/>
                            <TextBlock Grid.Row="4" Grid.Column="0" Text="Model:" Foreground="#AAAAAA" Margin="0,4"/>
                            <TextBlock Grid.Row="4" Grid.Column="1" x:Name="valModel" Text="Unknown" Foreground="#E0E0E0" Margin="8,4"/>
                            <TextBlock Grid.Row="5" Grid.Column="0" Text="System Type:" Foreground="#AAAAAA" Margin="0,4"/>
                            <TextBlock Grid.Row="5" Grid.Column="1" x:Name="valSystemType" Text="Unknown" Foreground="#E0E0E0" Margin="8,4"/>
                            <TextBlock Grid.Row="6" Grid.Column="0" Text="Serial Number:" Foreground="#AAAAAA" Margin="0,4"/>
                            <TextBlock Grid.Row="6" Grid.Column="1" x:Name="valSerialNumber" Text="Unknown" Foreground="#E0E0E0" Margin="8,4"/>
                            <TextBlock Grid.Row="7" Grid.Column="0" Text="Asset Tag:" Foreground="#AAAAAA" Margin="0,4"/>
                            <TextBlock Grid.Row="7" Grid.Column="1" x:Name="valAssetTag" Text="Unknown" Foreground="#E0E0E0" Margin="8,4"/>
                            <TextBlock Grid.Row="8" Grid.Column="0" Text="Chassis Type:" Foreground="#AAAAAA" Margin="0,4"/>
                            <TextBlock Grid.Row="8" Grid.Column="1" x:Name="valChassisType" Text="Unknown" Foreground="#E0E0E0" Margin="8,4"/>
                        </Grid>
                    </ScrollViewer>
                </TabItem>

                <!-- OS tab -->
                <TabItem Header="OS">
                    <ScrollViewer VerticalScrollBarVisibility="Auto" Margin="8">
                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="140"/>
                                <ColumnDefinition Width="*"/>
                            </Grid.ColumnDefinitions>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/><RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/><RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/><RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/><RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                            </Grid.RowDefinitions>
                            <TextBlock Grid.Row="0" Grid.Column="0" Text="OS Name:" Foreground="#AAAAAA" Margin="0,4"/>
                            <TextBlock Grid.Row="0" Grid.Column="1" x:Name="valOSName" Text="Unknown" Foreground="#E0E0E0" Margin="8,4" TextWrapping="Wrap"/>
                            <TextBlock Grid.Row="1" Grid.Column="0" Text="Version:" Foreground="#AAAAAA" Margin="0,4"/>
                            <TextBlock Grid.Row="1" Grid.Column="1" x:Name="valOSVersion" Text="Unknown" Foreground="#E0E0E0" Margin="8,4"/>
                            <TextBlock Grid.Row="2" Grid.Column="0" Text="Build:" Foreground="#AAAAAA" Margin="0,4"/>
                            <TextBlock Grid.Row="2" Grid.Column="1" x:Name="valOSBuild" Text="Unknown" Foreground="#E0E0E0" Margin="8,4"/>
                            <TextBlock Grid.Row="3" Grid.Column="0" Text="Architecture:" Foreground="#AAAAAA" Margin="0,4"/>
                            <TextBlock Grid.Row="3" Grid.Column="1" x:Name="valOSArch" Text="Unknown" Foreground="#E0E0E0" Margin="8,4"/>
                            <TextBlock Grid.Row="4" Grid.Column="0" Text="Install Date:" Foreground="#AAAAAA" Margin="0,4"/>
                            <TextBlock Grid.Row="4" Grid.Column="1" x:Name="valInstallDate" Text="Unknown" Foreground="#E0E0E0" Margin="8,4"/>
                            <TextBlock Grid.Row="5" Grid.Column="0" Text="Last Boot:" Foreground="#AAAAAA" Margin="0,4"/>
                            <TextBlock Grid.Row="5" Grid.Column="1" x:Name="valLastBoot" Text="Unknown" Foreground="#E0E0E0" Margin="8,4"/>
                            <TextBlock Grid.Row="6" Grid.Column="0" Text="Uptime:" Foreground="#AAAAAA" Margin="0,4"/>
                            <TextBlock Grid.Row="6" Grid.Column="1" x:Name="valUptime" Text="Unknown" Foreground="#E0E0E0" Margin="8,4"/>
                            <TextBlock Grid.Row="7" Grid.Column="0" Text="Registered Owner:" Foreground="#AAAAAA" Margin="0,4"/>
                            <TextBlock Grid.Row="7" Grid.Column="1" x:Name="valRegOwner" Text="Unknown" Foreground="#E0E0E0" Margin="8,4"/>
                            <TextBlock Grid.Row="8" Grid.Column="0" Text="Product ID:" Foreground="#AAAAAA" Margin="0,4"/>
                            <TextBlock Grid.Row="8" Grid.Column="1" x:Name="valProductID" Text="Unknown" Foreground="#E0E0E0" Margin="8,4"/>
                        </Grid>
                    </ScrollViewer>
                </TabItem>

                <!-- Processor tab -->
                <TabItem Header="Processor">
                    <ScrollViewer VerticalScrollBarVisibility="Auto" Margin="8">
                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="160"/>
                                <ColumnDefinition Width="*"/>
                            </Grid.ColumnDefinitions>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/><RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/><RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/><RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/><RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/><RowDefinition Height="Auto"/>
                            </Grid.RowDefinitions>
                            <TextBlock Grid.Row="0" Grid.Column="0" Text="Name:" Foreground="#AAAAAA" Margin="0,4"/>
                            <TextBlock Grid.Row="0" Grid.Column="1" x:Name="valCPUName" Text="Unknown" Foreground="#E0E0E0" Margin="8,4" TextWrapping="Wrap"/>
                            <TextBlock Grid.Row="1" Grid.Column="0" Text="Manufacturer:" Foreground="#AAAAAA" Margin="0,4"/>
                            <TextBlock Grid.Row="1" Grid.Column="1" x:Name="valCPUMfg" Text="Unknown" Foreground="#E0E0E0" Margin="8,4"/>
                            <TextBlock Grid.Row="2" Grid.Column="0" Text="Cores:" Foreground="#AAAAAA" Margin="0,4"/>
                            <TextBlock Grid.Row="2" Grid.Column="1" x:Name="valCPUCores" Text="Unknown" Foreground="#E0E0E0" Margin="8,4"/>
                            <TextBlock Grid.Row="3" Grid.Column="0" Text="Logical Processors:" Foreground="#AAAAAA" Margin="0,4"/>
                            <TextBlock Grid.Row="3" Grid.Column="1" x:Name="valCPULP" Text="Unknown" Foreground="#E0E0E0" Margin="8,4"/>
                            <TextBlock Grid.Row="4" Grid.Column="0" Text="Max Clock (MHz):" Foreground="#AAAAAA" Margin="0,4"/>
                            <TextBlock Grid.Row="4" Grid.Column="1" x:Name="valCPUMaxClk" Text="Unknown" Foreground="#E0E0E0" Margin="8,4"/>
                            <TextBlock Grid.Row="5" Grid.Column="0" Text="Current Clock (MHz):" Foreground="#AAAAAA" Margin="0,4"/>
                            <TextBlock Grid.Row="5" Grid.Column="1" x:Name="valCPUCurClk" Text="Unknown" Foreground="#E0E0E0" Margin="8,4"/>
                            <TextBlock Grid.Row="6" Grid.Column="0" Text="Architecture:" Foreground="#AAAAAA" Margin="0,4"/>
                            <TextBlock Grid.Row="6" Grid.Column="1" x:Name="valCPUArch" Text="Unknown" Foreground="#E0E0E0" Margin="8,4"/>
                            <TextBlock Grid.Row="7" Grid.Column="0" Text="L2 Cache (KB):" Foreground="#AAAAAA" Margin="0,4"/>
                            <TextBlock Grid.Row="7" Grid.Column="1" x:Name="valCPUL2" Text="Unknown" Foreground="#E0E0E0" Margin="8,4"/>
                            <TextBlock Grid.Row="8" Grid.Column="0" Text="L3 Cache (KB):" Foreground="#AAAAAA" Margin="0,4"/>
                            <TextBlock Grid.Row="8" Grid.Column="1" x:Name="valCPUL3" Text="Unknown" Foreground="#E0E0E0" Margin="8,4"/>
                            <TextBlock Grid.Row="9" Grid.Column="0" Text="Socket:" Foreground="#AAAAAA" Margin="0,4"/>
                            <TextBlock Grid.Row="9" Grid.Column="1" x:Name="valCPUSocket" Text="Unknown" Foreground="#E0E0E0" Margin="8,4"/>
                        </Grid>
                    </ScrollViewer>
                </TabItem>

                <!-- Memory tab -->
                <TabItem Header="Memory">
                    <ScrollViewer VerticalScrollBarVisibility="Auto" Margin="8">
                        <StackPanel>
                            <Grid Margin="0,0,0,12">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="140"/>
                                    <ColumnDefinition Width="*"/>
                                </Grid.ColumnDefinitions>
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto"/><RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/><RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/><RowDefinition Height="Auto"/>
                                </Grid.RowDefinitions>
                                <TextBlock Grid.Row="0" Grid.Column="0" Text="Total (GB):" Foreground="#AAAAAA" Margin="0,4"/>
                                <TextBlock Grid.Row="0" Grid.Column="1" x:Name="valMemTotal" Text="Unknown" Foreground="#E0E0E0" Margin="8,4"/>
                                <TextBlock Grid.Row="1" Grid.Column="0" Text="Available (GB):" Foreground="#AAAAAA" Margin="0,4"/>
                                <TextBlock Grid.Row="1" Grid.Column="1" x:Name="valMemAvail" Text="Unknown" Foreground="#E0E0E0" Margin="8,4"/>
                                <TextBlock Grid.Row="2" Grid.Column="0" Text="Used (GB):" Foreground="#AAAAAA" Margin="0,4"/>
                                <TextBlock Grid.Row="2" Grid.Column="1" x:Name="valMemUsed" Text="Unknown" Foreground="#E0E0E0" Margin="8,4"/>
                                <TextBlock Grid.Row="3" Grid.Column="0" Text="Usage (%):" Foreground="#AAAAAA" Margin="0,4"/>
                                <TextBlock Grid.Row="3" Grid.Column="1" x:Name="valMemPct" Text="Unknown" Foreground="#E0E0E0" Margin="8,4"/>
                                <TextBlock Grid.Row="4" Grid.Column="0" Text="Total Slots:" Foreground="#AAAAAA" Margin="0,4"/>
                                <TextBlock Grid.Row="4" Grid.Column="1" x:Name="valMemSlots" Text="Unknown" Foreground="#E0E0E0" Margin="8,4"/>
                                <TextBlock Grid.Row="5" Grid.Column="0" Text="Used Slots:" Foreground="#AAAAAA" Margin="0,4"/>
                                <TextBlock Grid.Row="5" Grid.Column="1" x:Name="valMemUsedSlots" Text="Unknown" Foreground="#E0E0E0" Margin="8,4"/>
                            </Grid>
                            <TextBlock Text="DIMMs" Foreground="#E0E0E0" FontWeight="SemiBold" Margin="0,0,0,4"/>
                            <DataGrid x:Name="gridDIMMs" Height="180"/>
                        </StackPanel>
                    </ScrollViewer>
                </TabItem>

                <!-- Storage tab -->
                <TabItem Header="Storage">
                    <ScrollViewer VerticalScrollBarVisibility="Auto" Margin="8">
                        <StackPanel>
                            <TextBlock Text="Physical Disks" Foreground="#E0E0E0" FontWeight="SemiBold" Margin="0,0,0,4"/>
                            <DataGrid x:Name="gridPhysDisks" Height="180" Margin="0,0,0,12"/>
                            <TextBlock Text="Logical Volumes" Foreground="#E0E0E0" FontWeight="SemiBold" Margin="0,0,0,4"/>
                            <DataGrid x:Name="gridLogVols" Height="180"/>
                        </StackPanel>
                    </ScrollViewer>
                </TabItem>

                <!-- Graphics tab -->
                <TabItem Header="Graphics">
                    <ScrollViewer VerticalScrollBarVisibility="Auto" Margin="8">
                        <StackPanel>
                            <TextBlock Text="Graphics Adapters" Foreground="#E0E0E0" FontWeight="SemiBold" Margin="0,0,0,4"/>
                            <DataGrid x:Name="gridGPU" Height="250"/>
                        </StackPanel>
                    </ScrollViewer>
                </TabItem>

                <!-- Network tab -->
                <TabItem Header="Network">
                    <ScrollViewer VerticalScrollBarVisibility="Auto" Margin="8">
                        <StackPanel>
                            <TextBlock Text="Network Adapters" Foreground="#E0E0E0" FontWeight="SemiBold" Margin="0,0,0,4"/>
                            <DataGrid x:Name="gridNetwork" Height="350"/>
                        </StackPanel>
                    </ScrollViewer>
                </TabItem>

                <!-- BIOS/Board tab -->
                <TabItem Header="BIOS/Board">
                    <ScrollViewer VerticalScrollBarVisibility="Auto" Margin="8">
                        <StackPanel>
                            <TextBlock Text="BIOS" Foreground="#E0E0E0" FontWeight="SemiBold" FontSize="14" Margin="0,0,0,6"/>
                            <Grid Margin="0,0,0,16">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="150"/>
                                    <ColumnDefinition Width="*"/>
                                </Grid.ColumnDefinitions>
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto"/><RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/><RowDefinition Height="Auto"/>
                                </Grid.RowDefinitions>
                                <TextBlock Grid.Row="0" Grid.Column="0" Text="Manufacturer:" Foreground="#AAAAAA" Margin="0,4"/>
                                <TextBlock Grid.Row="0" Grid.Column="1" x:Name="valBIOSMfg" Text="Unknown" Foreground="#E0E0E0" Margin="8,4" TextWrapping="Wrap"/>
                                <TextBlock Grid.Row="1" Grid.Column="0" Text="Version:" Foreground="#AAAAAA" Margin="0,4"/>
                                <TextBlock Grid.Row="1" Grid.Column="1" x:Name="valBIOSVer" Text="Unknown" Foreground="#E0E0E0" Margin="8,4" TextWrapping="Wrap"/>
                                <TextBlock Grid.Row="2" Grid.Column="0" Text="Date:" Foreground="#AAAAAA" Margin="0,4"/>
                                <TextBlock Grid.Row="2" Grid.Column="1" x:Name="valBIOSDate" Text="Unknown" Foreground="#E0E0E0" Margin="8,4"/>
                                <TextBlock Grid.Row="3" Grid.Column="0" Text="SMBIOS Version:" Foreground="#AAAAAA" Margin="0,4"/>
                                <TextBlock Grid.Row="3" Grid.Column="1" x:Name="valSMBIOS" Text="Unknown" Foreground="#E0E0E0" Margin="8,4"/>
                            </Grid>
                            <TextBlock Text="Motherboard" Foreground="#E0E0E0" FontWeight="SemiBold" FontSize="14" Margin="0,0,0,6"/>
                            <Grid>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="150"/>
                                    <ColumnDefinition Width="*"/>
                                </Grid.ColumnDefinitions>
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto"/><RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/><RowDefinition Height="Auto"/>
                                </Grid.RowDefinitions>
                                <TextBlock Grid.Row="0" Grid.Column="0" Text="Manufacturer:" Foreground="#AAAAAA" Margin="0,4"/>
                                <TextBlock Grid.Row="0" Grid.Column="1" x:Name="valMBMfg" Text="Unknown" Foreground="#E0E0E0" Margin="8,4"/>
                                <TextBlock Grid.Row="1" Grid.Column="0" Text="Product:" Foreground="#AAAAAA" Margin="0,4"/>
                                <TextBlock Grid.Row="1" Grid.Column="1" x:Name="valMBProduct" Text="Unknown" Foreground="#E0E0E0" Margin="8,4"/>
                                <TextBlock Grid.Row="2" Grid.Column="0" Text="Version:" Foreground="#AAAAAA" Margin="0,4"/>
                                <TextBlock Grid.Row="2" Grid.Column="1" x:Name="valMBVersion" Text="Unknown" Foreground="#E0E0E0" Margin="8,4"/>
                                <TextBlock Grid.Row="3" Grid.Column="0" Text="Serial Number:" Foreground="#AAAAAA" Margin="0,4"/>
                                <TextBlock Grid.Row="3" Grid.Column="1" x:Name="valMBSerial" Text="Unknown" Foreground="#E0E0E0" Margin="8,4"/>
                            </Grid>
                        </StackPanel>
                    </ScrollViewer>
                </TabItem>

                <!-- Battery tab -->
                <TabItem Header="Battery">
                    <ScrollViewer VerticalScrollBarVisibility="Auto" Margin="8">
                        <StackPanel>
                            <TextBlock x:Name="NoBatteryText"
                                       Text="No battery detected"
                                       Foreground="#888888" FontSize="14" FontStyle="Italic"
                                       HorizontalAlignment="Center" VerticalAlignment="Center"
                                       Margin="0,40,0,0"/>
                            <Grid x:Name="BatteryGrid" Visibility="Collapsed">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="160"/>
                                    <ColumnDefinition Width="*"/>
                                </Grid.ColumnDefinitions>
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto"/><RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/><RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/><RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                </Grid.RowDefinitions>
                                <TextBlock Grid.Row="0" Grid.Column="0" Text="Status:" Foreground="#AAAAAA" Margin="0,4"/>
                                <TextBlock Grid.Row="0" Grid.Column="1" x:Name="valBatStatus" Text="Unknown" Foreground="#E0E0E0" Margin="8,4"/>
                                <TextBlock Grid.Row="1" Grid.Column="0" Text="Charge (%):" Foreground="#AAAAAA" Margin="0,4"/>
                                <TextBlock Grid.Row="1" Grid.Column="1" x:Name="valBatCharge" Text="Unknown" Foreground="#E0E0E0" Margin="8,4"/>
                                <TextBlock Grid.Row="2" Grid.Column="0" Text="Est. Runtime (min):" Foreground="#AAAAAA" Margin="0,4"/>
                                <TextBlock Grid.Row="2" Grid.Column="1" x:Name="valBatRuntime" Text="Unknown" Foreground="#E0E0E0" Margin="8,4"/>
                                <TextBlock Grid.Row="3" Grid.Column="0" Text="Design Capacity:" Foreground="#AAAAAA" Margin="0,4"/>
                                <TextBlock Grid.Row="3" Grid.Column="1" x:Name="valBatDesign" Text="Unknown" Foreground="#E0E0E0" Margin="8,4"/>
                                <TextBlock Grid.Row="4" Grid.Column="0" Text="Full Charge Cap.:" Foreground="#AAAAAA" Margin="0,4"/>
                                <TextBlock Grid.Row="4" Grid.Column="1" x:Name="valBatFullCharge" Text="Unknown" Foreground="#E0E0E0" Margin="8,4"/>
                                <TextBlock Grid.Row="5" Grid.Column="0" Text="Battery Health (%):" Foreground="#AAAAAA" Margin="0,4"/>
                                <TextBlock Grid.Row="5" Grid.Column="1" x:Name="valBatHealth" Text="Unknown" Foreground="#E0E0E0" Margin="8,4"/>
                            </Grid>
                        </StackPanel>
                    </ScrollViewer>
                </TabItem>

                <!-- Other tab -->
                <TabItem Header="Other">
                    <ScrollViewer VerticalScrollBarVisibility="Auto" Margin="8">
                        <StackPanel>
                            <TextBlock Text="Installed Hotfixes" Foreground="#E0E0E0" FontWeight="SemiBold" Margin="0,0,0,4"/>
                            <DataGrid x:Name="gridHotfixes" Height="200" Margin="0,0,0,12"/>
                            <TextBlock Text="Startup Programs" Foreground="#E0E0E0" FontWeight="SemiBold" Margin="0,0,0,4"/>
                            <DataGrid x:Name="gridStartup" Height="200"/>
                        </StackPanel>
                    </ScrollViewer>
                </TabItem>

            </TabControl>

            <!-- Right panel: actions -->
            <Border Grid.Column="2" Background="#252525" CornerRadius="4" Margin="4" Padding="10">
                <StackPanel VerticalAlignment="Top">
                    <TextBlock Text="Actions" FontSize="14" FontWeight="SemiBold"
                               Foreground="#E0E0E0" FontFamily="Segoe UI"
                               Margin="0,0,0,12"/>

                    <Button x:Name="btnScan" Content="&#x1F50D;  Run Scan"
                            Style="{StaticResource AccentButton}" Margin="0,0,0,16"
                            HorizontalAlignment="Stretch"/>

                    <Separator Background="#3C3C3C" Margin="0,0,0,12"/>

                    <TextBlock Text="Export Format" Foreground="#AAAAAA" FontFamily="Segoe UI"
                               FontSize="11" Margin="0,0,0,4"/>
                    <ComboBox x:Name="cmbFormat" Background="#3C3C3C" Foreground="#CCCCCC"
                              FontFamily="Segoe UI" FontSize="12"
                              SelectedIndex="0" Margin="0,0,0,8">
                        <ComboBoxItem Content="TXT"/>
                        <ComboBoxItem Content="CSV"/>
                        <ComboBoxItem Content="Both"/>
                    </ComboBox>

                    <TextBlock Text="Output Location" Foreground="#AAAAAA" FontFamily="Segoe UI"
                               FontSize="11" Margin="0,0,0,4"/>
                    <TextBox x:Name="txtOutputPath" Background="#333333" Foreground="#CCCCCC"
                             FontFamily="Segoe UI" FontSize="11"
                             BorderBrush="#555555" Padding="4,3" Margin="0,0,0,4"/>
                    <Button x:Name="btnBrowse" Content="Browse..."
                            Style="{StaticResource StdButton}" Margin="0,0,0,8"
                            HorizontalAlignment="Stretch"/>

                    <Button x:Name="btnExport" Content="Export"
                            Style="{StaticResource StdButton}" Margin="0,0,0,16"
                            HorizontalAlignment="Stretch"/>

                    <Separator Background="#3C3C3C" Margin="0,0,0,12"/>

                    <Button x:Name="btnConsole" Content="View Console Table"
                            Style="{StaticResource StdButton}" Margin="0,0,0,0"
                            HorizontalAlignment="Stretch"/>
                </StackPanel>
            </Border>
        </Grid>

        <!-- Status bar -->
        <Border Grid.Row="2" Background="#007ACC" Padding="10,6">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="200"/>
                </Grid.ColumnDefinitions>
                <TextBlock x:Name="StatusText" Grid.Column="0"
                           Text="Ready" Foreground="White"
                           FontFamily="Segoe UI" FontSize="12"
                           VerticalAlignment="Center"/>
                <ProgressBar x:Name="ProgressBar" Grid.Column="1"
                             Minimum="0" Maximum="100" Value="0"
                             Height="16" Foreground="#0078D4" Background="#005A9E"/>
            </Grid>
        </Border>
    </Grid>
</Window>
"@

# ── Build Window ───────────────────────────────────────────────────────────────

# Remove the x:Class attribute if present (not needed for dynamic parsing)
$xaml.Window.RemoveAttribute('x:Class') 2>$null

$reader = [System.Xml.XmlNodeReader]::new($xaml)
$window = [System.Windows.Markup.XamlReader]::Load($reader)

# ── Resolve Named Elements ─────────────────────────────────────────────────────

$ui = @{}
$namedElements = @(
    'ImagePlaceholder', 'HardwareImage', 'MakeLabel', 'ModelLabel',
    'MainTabs',
    # System
    'valComputerName', 'valCurrentUser', 'valDomain', 'valMake', 'valModel',
    'valSystemType', 'valSerialNumber', 'valAssetTag', 'valChassisType',
    # OS
    'valOSName', 'valOSVersion', 'valOSBuild', 'valOSArch',
    'valInstallDate', 'valLastBoot', 'valUptime', 'valRegOwner', 'valProductID',
    # Processor
    'valCPUName', 'valCPUMfg', 'valCPUCores', 'valCPULP',
    'valCPUMaxClk', 'valCPUCurClk', 'valCPUArch', 'valCPUL2', 'valCPUL3', 'valCPUSocket',
    # Memory
    'valMemTotal', 'valMemAvail', 'valMemUsed', 'valMemPct',
    'valMemSlots', 'valMemUsedSlots', 'gridDIMMs',
    # Storage
    'gridPhysDisks', 'gridLogVols',
    # Graphics
    'gridGPU',
    # Network
    'gridNetwork',
    # BIOS / Board
    'valBIOSMfg', 'valBIOSVer', 'valBIOSDate', 'valSMBIOS',
    'valMBMfg', 'valMBProduct', 'valMBVersion', 'valMBSerial',
    # Battery
    'NoBatteryText', 'BatteryGrid',
    'valBatStatus', 'valBatCharge', 'valBatRuntime',
    'valBatDesign', 'valBatFullCharge', 'valBatHealth',
    # Other
    'gridHotfixes', 'gridStartup',
    # Actions
    'btnScan', 'btnBrowse', 'btnExport', 'btnConsole',
    'cmbFormat', 'txtOutputPath',
    # Status
    'StatusText', 'ProgressBar'
)

foreach ($name in $namedElements) {
    $ui[$name] = $window.FindName($name)
}

# Set default output path
$ui['txtOutputPath'].Text = [Environment]::GetFolderPath('Desktop')

# ── State ──────────────────────────────────────────────────────────────────────

$script:ScanData = $null

# ── Helper: safe string conversion ─────────────────────────────────────────────

function Get-SafeValue {
    param($Value, [string]$Default = 'Unknown')
    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace("$Value")) { return $Default }
    return "$Value"
}

# ── Populate UI from scan data ─────────────────────────────────────────────────

function Update-UIFromData {
    param($Data)

    $d = $Data

    # System overview
    $so = $d.SystemOverview
    if ($so) {
        $ui['valComputerName'].Text = Get-SafeValue $so.ComputerName
        $ui['valCurrentUser'].Text  = Get-SafeValue $so.CurrentUser
        $ui['valDomain'].Text       = Get-SafeValue $so.Domain
        $ui['valMake'].Text         = Get-SafeValue $so.SystemManufacturer
        $ui['valModel'].Text        = Get-SafeValue $so.SystemModel
        $ui['valSystemType'].Text   = Get-SafeValue $so.SystemType
        $ui['valSerialNumber'].Text = Get-SafeValue $so.SerialNumber
        $ui['valAssetTag'].Text     = Get-SafeValue $so.AssetTag
        $ui['valChassisType'].Text  = Get-SafeValue $so.ChassisType

        $ui['MakeLabel'].Text  = "Make: $(Get-SafeValue $so.SystemManufacturer)"
        $ui['ModelLabel'].Text = "Model: $(Get-SafeValue $so.SystemModel)"
    }

    # OS
    $os = $d.OperatingSystem
    if ($os) {
        $ui['valOSName'].Text      = Get-SafeValue $os.OSName
        $ui['valOSVersion'].Text   = Get-SafeValue $os.OSVersion
        $ui['valOSBuild'].Text     = Get-SafeValue $os.OSBuild
        $ui['valOSArch'].Text      = Get-SafeValue $os.OSArchitecture
        $ui['valInstallDate'].Text = Get-SafeValue $os.InstallDate
        $ui['valLastBoot'].Text    = Get-SafeValue $os.LastBootTime
        $ui['valUptime'].Text      = Get-SafeValue $os.Uptime
        $ui['valRegOwner'].Text    = Get-SafeValue $os.RegisteredOwner
        $ui['valProductID'].Text   = Get-SafeValue $os.ProductID
    }

    # Processor
    $cpu = $d.Processor
    if ($cpu) {
        $ui['valCPUName'].Text   = Get-SafeValue $cpu.ProcessorName
        $ui['valCPUMfg'].Text    = Get-SafeValue $cpu.Manufacturer
        $ui['valCPUCores'].Text  = Get-SafeValue $cpu.NumberOfCores
        $ui['valCPULP'].Text     = Get-SafeValue $cpu.NumberOfLogicalProcessors
        $ui['valCPUMaxClk'].Text = Get-SafeValue $cpu.MaxClockSpeedMHz
        $ui['valCPUCurClk'].Text = Get-SafeValue $cpu.CurrentClockSpeedMHz
        $ui['valCPUArch'].Text   = Get-SafeValue $cpu.Architecture
        $ui['valCPUL2'].Text     = Get-SafeValue $cpu.L2CacheSizeKB
        $ui['valCPUL3'].Text     = Get-SafeValue $cpu.L3CacheSizeKB
        $ui['valCPUSocket'].Text = Get-SafeValue $cpu.SocketDesignation
    }

    # Memory
    $mem = $d.Memory
    if ($mem) {
        $ui['valMemTotal'].Text     = Get-SafeValue $mem.TotalPhysicalMemoryGB
        $ui['valMemAvail'].Text     = Get-SafeValue $mem.AvailableMemoryGB
        $ui['valMemUsed'].Text      = Get-SafeValue $mem.UsedMemoryGB
        $ui['valMemPct'].Text       = Get-SafeValue $mem.MemoryUsagePercent
        $ui['valMemSlots'].Text     = Get-SafeValue $mem.TotalSlots
        $ui['valMemUsedSlots'].Text = Get-SafeValue $mem.UsedSlots
        if ($mem.DIMMs) {
            $ui['gridDIMMs'].ItemsSource = [System.Collections.ArrayList]@($mem.DIMMs)
        }
    }

    # Storage
    $stor = $d.Storage
    if ($stor) {
        if ($stor.PhysicalDisks)  { $ui['gridPhysDisks'].ItemsSource = [System.Collections.ArrayList]@($stor.PhysicalDisks) }
        if ($stor.LogicalVolumes) { $ui['gridLogVols'].ItemsSource   = [System.Collections.ArrayList]@($stor.LogicalVolumes) }
    }

    # Graphics
    $gfx = $d.Graphics
    if ($gfx) {
        $ui['gridGPU'].ItemsSource = [System.Collections.ArrayList]@($gfx)
    }

    # Network
    $net = $d.NetworkAdapters
    if ($net) {
        $ui['gridNetwork'].ItemsSource = [System.Collections.ArrayList]@($net)
    }

    # BIOS
    $bios = $d.BIOS
    if ($bios) {
        $ui['valBIOSMfg'].Text  = Get-SafeValue $bios.BIOSManufacturer
        $ui['valBIOSVer'].Text  = Get-SafeValue $bios.BIOSVersion
        $ui['valBIOSDate'].Text = Get-SafeValue $bios.BIOSDate
        $ui['valSMBIOS'].Text   = Get-SafeValue $bios.SMBIOSVersion
    }

    # Motherboard
    $mb = $d.Motherboard
    if ($mb) {
        $ui['valMBMfg'].Text     = Get-SafeValue $mb.Manufacturer
        $ui['valMBProduct'].Text = Get-SafeValue $mb.Product
        $ui['valMBVersion'].Text = Get-SafeValue $mb.Version
        $ui['valMBSerial'].Text  = Get-SafeValue $mb.SerialNumber
    }

    # Battery
    $bat = $d.Battery
    if ($bat) {
        if ($bat.HasBattery) {
            $ui['NoBatteryText'].Visibility = 'Collapsed'
            $ui['BatteryGrid'].Visibility   = 'Visible'
            $ui['valBatStatus'].Text      = Get-SafeValue $bat.Status
            $ui['valBatCharge'].Text       = Get-SafeValue $bat.ChargePercent
            $ui['valBatRuntime'].Text      = Get-SafeValue $bat.EstimatedRuntime
            $ui['valBatDesign'].Text       = Get-SafeValue $bat.DesignCapacity
            $ui['valBatFullCharge'].Text   = Get-SafeValue $bat.FullChargeCapacity
            $ui['valBatHealth'].Text       = Get-SafeValue $bat.BatteryHealth
        }
        else {
            $ui['NoBatteryText'].Visibility = 'Visible'
            $ui['BatteryGrid'].Visibility   = 'Collapsed'
        }
    }

    # Hotfixes & Startup
    if ($d.Hotfixes)        { $ui['gridHotfixes'].ItemsSource = [System.Collections.ArrayList]@($d.Hotfixes) }
    if ($d.StartupPrograms) { $ui['gridStartup'].ItemsSource  = [System.Collections.ArrayList]@($d.StartupPrograms) }

    # Hardware image
    try {
        if ($so -and $so.SystemManufacturer -and $so.SystemModel) {
            $chassisCat = if ($so.SystemType) { $so.SystemType } else { 'Desktop' }
            $imgPath = Get-HardwareImagePath -Manufacturer $so.SystemManufacturer `
                                             -Model $so.SystemModel `
                                             -ChassisCategory $chassisCat
            if ($imgPath -and (Test-Path $imgPath)) {
                $bitmap = [System.Windows.Media.Imaging.BitmapImage]::new()
                $bitmap.BeginInit()
                $bitmap.UriSource = [Uri]::new($imgPath)
                $bitmap.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
                $bitmap.EndInit()
                $ui['HardwareImage'].Source     = $bitmap
                $ui['HardwareImage'].Visibility = 'Visible'
                $ui['ImagePlaceholder'].Visibility = 'Collapsed'
            }
        }
    }
    catch {
        # Keep placeholder visible on error
    }
}

# ── Run Scan (background runspace + DispatcherTimer) ───────────────────────────

$ui['btnScan'].Add_Click({
    $ui['btnScan'].IsEnabled = $false
    $ui['StatusText'].Text   = 'Scanning...'
    $ui['ProgressBar'].Value = 0

    $dispatcher = $window.Dispatcher

    # Shared synchronized hashtable for cross-thread communication
    $syncHash = [hashtable]::Synchronized(@{
        Completed = $false
        Result    = $null
        Error     = $null
        Progress  = 0
        Status    = 'Starting scan...'
    })

    # Background runspace
    $runspace = [RunspaceFactory]::CreateRunspace()
    $runspace.ApartmentState = 'STA'
    $runspace.Open()

    $psCmd = [powershell]::Create()
    $psCmd.Runspace = $runspace

    [void]$psCmd.AddScript({
        param($SyncHash, $ScriptRoot)

        # Dot-source the core module inside the runspace
        . "$ScriptRoot\SysInfo-Core.ps1"

        try {
            $data = Get-SystemInfoData -ProgressCallback {
                param([int]$Pct, [string]$Msg)
                $SyncHash.Progress = $Pct
                $SyncHash.Status   = $Msg
            }
            $SyncHash.Result = $data
        }
        catch {
            $SyncHash.Error = $_.Exception.Message
        }
        finally {
            $SyncHash.Completed = $true
        }
    })

    [void]$psCmd.AddArgument($syncHash)
    [void]$psCmd.AddArgument($PSScriptRoot)
    $psCmd.BeginInvoke() | Out-Null

    # DispatcherTimer polls the runspace for progress and completion
    $timer = [System.Windows.Threading.DispatcherTimer]::new()
    $timer.Interval = [TimeSpan]::FromMilliseconds(200)

    $timer.Add_Tick({
        $ui['ProgressBar'].Value = $syncHash.Progress
        $ui['StatusText'].Text   = $syncHash.Status

        if ($syncHash.Completed) {
            $timer.Stop()

            if ($syncHash.Error) {
                $ui['StatusText'].Text   = "Scan failed: $($syncHash.Error)"
                $ui['ProgressBar'].Value = 0
            }
            else {
                $script:ScanData = $syncHash.Result
                Update-UIFromData -Data $script:ScanData

                # Count successfully collected categories
                $categories = @(
                    'SystemOverview','OperatingSystem','Processor','Memory',
                    'Storage','Graphics','NetworkAdapters','BIOS',
                    'Motherboard','Battery','Hotfixes','StartupPrograms'
                )
                $collected = 0
                foreach ($cat in $categories) {
                    $val = $script:ScanData.$cat
                    if ($null -ne $val) {
                        $hasError = $false
                        try { $hasError = [bool]$val.Error } catch {}
                        if (-not $hasError) { $collected++ }
                    }
                }
                $ui['StatusText'].Text   = "Scan complete - $collected of $($categories.Count) categories collected"
                $ui['ProgressBar'].Value = 100
            }

            $ui['btnScan'].IsEnabled = $true

            # Cleanup
            try { $psCmd.Dispose() } catch {}
            try { $runspace.Dispose() } catch {}
        }
    }.GetNewClosure())

    $timer.Start()
})

# ── Browse Button ──────────────────────────────────────────────────────────────

$ui['btnBrowse'].Add_Click({
    $dialog = [System.Windows.Forms.FolderBrowserDialog]::new()
    $dialog.Description = 'Select output folder for export'
    $dialog.SelectedPath = $ui['txtOutputPath'].Text

    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $ui['txtOutputPath'].Text = $dialog.SelectedPath
    }
})

# ── Export Button ──────────────────────────────────────────────────────────────

$ui['btnExport'].Add_Click({
    if ($null -eq $script:ScanData) {
        [System.Windows.MessageBox]::Show(
            'No scan data available. Please run a scan first.',
            'Export', 'OK', 'Warning')
        return
    }

    $outDir = $ui['txtOutputPath'].Text
    if (-not (Test-Path $outDir)) {
        [System.Windows.MessageBox]::Show(
            "Output directory does not exist:`n$outDir",
            'Export', 'OK', 'Warning')
        return
    }

    $format = $ui['cmbFormat'].SelectedItem.Content
    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $computerName = if ($script:ScanData.SystemOverview.ComputerName) {
        $script:ScanData.SystemOverview.ComputerName
    } else { 'UNKNOWN' }

    try {
        $exported = @()

        if ($format -eq 'TXT' -or $format -eq 'Both') {
            $txtPath = Join-Path $outDir "SysInfo_${computerName}_${timestamp}.txt"
            Export-SystemInfoTXT -Data $script:ScanData -Path $txtPath
            $exported += $txtPath
        }

        if ($format -eq 'CSV' -or $format -eq 'Both') {
            $csvPath = Join-Path $outDir "SysInfo_${computerName}_${timestamp}.csv"
            Export-SystemInfoCSV -Data $script:ScanData -Path $csvPath
            $exported += $csvPath
        }

        $fileList = ($exported | ForEach-Object { "  - $_" }) -join "`n"
        [System.Windows.MessageBox]::Show(
            "Export successful!`n`n$fileList",
            'Export', 'OK', 'Information')
        $ui['StatusText'].Text = "Exported to $outDir"
    }
    catch {
        [System.Windows.MessageBox]::Show(
            "Export failed:`n$($_.Exception.Message)",
            'Export Error', 'OK', 'Error')
    }
})

# ── View Console Table Button ──────────────────────────────────────────────────

$ui['btnConsole'].Add_Click({
    if ($null -eq $script:ScanData) {
        [System.Windows.MessageBox]::Show(
            'No scan data available. Please run a scan first.',
            'Console Table', 'OK', 'Warning')
        return
    }

    $tableText = Format-SystemInfoTable -Data $script:ScanData

    [xml]$consoleXaml = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    Title="System Info - Console View"
    Width="820" Height="620"
    WindowStartupLocation="CenterOwner"
    Background="#1E1E1E">
    <ScrollViewer HorizontalScrollBarVisibility="Auto" VerticalScrollBarVisibility="Auto" Margin="8">
        <TextBlock x:Name="ConsoleOutput"
                   FontFamily="Consolas"
                   FontSize="12"
                   Foreground="#CCCCCC"
                   Background="#1E1E1E"
                   TextWrapping="NoWrap"
                   xml:space="preserve"/>
    </ScrollViewer>
</Window>
"@

    $consoleReader = [System.Xml.XmlNodeReader]::new($consoleXaml)
    $consoleWindow = [System.Windows.Markup.XamlReader]::Load($consoleReader)
    $consoleWindow.FindName('ConsoleOutput').Text = $tableText
    $consoleWindow.Owner = $window
    $consoleWindow.ShowDialog() | Out-Null
})

# ── Show Window ────────────────────────────────────────────────────────────────

$window.ShowDialog() | Out-Null
