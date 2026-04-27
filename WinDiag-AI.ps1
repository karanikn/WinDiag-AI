<#
.SYNOPSIS
    WinDiag-AI - Windows Diagnostics with AI Analysis (GUI)
    by Nikolaos Karanikolas
.DESCRIPTION
    WPF GUI - collects diagnostics, sends to local AI (Ollama).
    Background threading via [PowerShell]::Create() + Runspace.
    Features: SMART details, PDF export, custom checks,
          driver check, dark/light theme, expanded settings, tooltips.
.NOTES
    Run as Administrator for full diagnostics.
#>
#Requires -Version 5.1

#region Admin + STA Guard
# Debug: show what PS sees (visible in console before GUI opens)
Write-Host "WinDiag-AI starting..." -ForegroundColor Cyan
Write-Host "  PSCommandPath    : $PSCommandPath" -ForegroundColor DarkGray
Write-Host "  MyCommand.Path   : $($MyInvocation.MyCommand.Path)" -ForegroundColor DarkGray
Write-Host "  PSScriptRoot     : $PSScriptRoot" -ForegroundColor DarkGray
Write-Host "  PWD              : $($PWD.Path)" -ForegroundColor DarkGray
$_ScriptPath = $null
# Method 1: PSCommandPath — set by PowerShell when launched with -File or Run with PowerShell
if(-not $_ScriptPath -and $PSCommandPath -and $PSCommandPath -ne '' -and (Test-Path $PSCommandPath -EA SilentlyContinue)){
    $_ScriptPath = $PSCommandPath
}
# Method 2: MyInvocation.MyCommand.Path — reliable in PS5.1 -File launches
if(-not $_ScriptPath -and $MyInvocation.MyCommand.Path -and (Test-Path $MyInvocation.MyCommand.Path -EA SilentlyContinue)){
    $_ScriptPath = $MyInvocation.MyCommand.Path
}
# Method 3: MyInvocation.MyCommand.Source
if(-not $_ScriptPath -and $MyInvocation.MyCommand.Source -and (Test-Path $MyInvocation.MyCommand.Source -EA SilentlyContinue)){
    $_ScriptPath = $MyInvocation.MyCommand.Source
}
# Method 4: PSScriptRoot + script name
if(-not $_ScriptPath -and $PSScriptRoot -and $PSScriptRoot -ne ''){
    $t = Join-Path $PSScriptRoot "WinDiag-AI.ps1"
    if(Test-Path $t -EA SilentlyContinue){ $_ScriptPath = $t }
}
# Method 5: Search PWD and parent
if(-not $_ScriptPath){
    foreach($sp in @($PWD.Path, (Split-Path $PWD.Path -Parent))){
        if($sp){ $t = Join-Path $sp "WinDiag-AI.ps1"; if(Test-Path $t -EA SilentlyContinue){ $_ScriptPath = $t; break } }
    }
}

$_ScriptDir = if($_ScriptPath){ Split-Path $_ScriptPath -Parent } else { $PWD.Path }
$_PSExe = (Get-Command powershell.exe -EA SilentlyContinue).Source
if(-not $_PSExe){ $_PSExe = (Get-Command pwsh -EA SilentlyContinue).Source }

# Elevate to Admin if needed
try {
    $cp = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
    if(-not $cp.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)){
        if($_ScriptPath -and $_PSExe){
            # Use -File with quoted path (handles spaces). -File preserves $PSCommandPath/$PSScriptRoot in child.
            Start-Process $_PSExe `
                -ArgumentList "-NoProfile","-STA","-ExecutionPolicy","Bypass","-File",('"{0}"' -f $_ScriptPath) `
                -Verb RunAs -WorkingDirectory $_ScriptDir
            exit
        }
    }
} catch {}
# Ensure STA apartment state (only needed if already elevated but wrong apartment)
try {
    if([Threading.Thread]::CurrentThread.ApartmentState -ne "STA"){
        if($_ScriptPath -and $_PSExe){
            Start-Process $_PSExe `
                -ArgumentList "-NoProfile","-STA","-ExecutionPolicy","Bypass","-File",('"{0}"' -f $_ScriptPath) `
                -Verb RunAs -WorkingDirectory $_ScriptDir
            exit
        }
    }
} catch {}
try { if($_ScriptPath){ Unblock-File $_ScriptPath -EA SilentlyContinue } } catch {}
#endregion

#region Globals
$ErrorActionPreference = "Continue"
$AppName = "WinDiag-AI"; $AppVer = "3.0"

# ScriptDir — reuse what we already resolved above
$Global:ScriptDir = $_ScriptDir

$Global:OllamaUrl = "http://localhost:11434"
$Script:Verbose = $false
$Script:SelectedModel = ""
$Global:DiagData = $null
$Global:AiAnalysis = ""
$Script:WorkerPS = $null
$Script:WorkerHandle = $null
$Global:ExtData = $null
$Global:ExtDiagData = $null
$Global:ExtJson = $null
$Global:ReportPath = $Global:ScriptDir
$Global:LogFilePath = Join-Path $Global:ScriptDir "WinDiag-AI.log"
$Global:AiTemp = 0.3
$Global:AiMaxTokens = 4096
$Global:EventHours = 24
$Global:AutoSaveReport = $false
$Global:IsDarkTheme = $false
$Global:CustomChecksDir = Join-Path $Global:ScriptDir "custom_checks"

$RecommendedModels = @(
    [PSCustomObject]@{Name="llama3.2:1b"; Size="~1.3 GB"; Quality="Basic - fast"; MinRAM=4; Info="Meta Llama 3.2 1B`n`nSpeed: Very Fast (5-15 sec)`nUse: Quick triage, simple checks`nAccuracy: Basic - may miss complex issues`nUpdated: Sep 2024`nURL: ollama.com/library/llama3.2:1b`n`nBest for low-RAM systems. Good for fast overviews but limited reasoning."}
    [PSCustomObject]@{Name="llama3.2";    Size="~2.0 GB"; Quality="Good balance";  MinRAM=6; Info="Meta Llama 3.2 3B`n`nSpeed: Fast (15-30 sec)`nUse: General diagnostics, good all-rounder`nAccuracy: Good - handles most scenarios`nUpdated: Sep 2024`nURL: ollama.com/library/llama3.2`n`nRecommended starting point. Balances speed with quality for everyday diagnostics."}
    [PSCustomObject]@{Name="phi3:mini";   Size="~2.3 GB"; Quality="Very good";     MinRAM=6; Info="Microsoft Phi-3 Mini 3.8B`n`nSpeed: Fast (15-40 sec)`nUse: Structured analysis, technical tasks`nAccuracy: Very Good - strong reasoning`nUpdated: Apr 2024`nURL: ollama.com/library/phi3:mini`n`nExcellent at following structured formats. Great for detailed repair commands."}
    [PSCustomObject]@{Name="phi3.5";      Size="~2.2 GB"; Quality="Very good+";    MinRAM=6; Info="Microsoft Phi-3.5 Mini 3.8B`n`nSpeed: Fast (15-40 sec)`nUse: Improved reasoning, multilingual`nAccuracy: Very Good - better than phi3`nUpdated: Aug 2024`nURL: ollama.com/library/phi3.5`n`nUpgraded Phi-3 with better multilingual support and improved reasoning."}
    [PSCustomObject]@{Name="gemma2:2b";   Size="~1.6 GB"; Quality="Compact";       MinRAM=4; Info="Google Gemma 2 2B`n`nSpeed: Very Fast (10-20 sec)`nUse: Quick diagnostics, lightweight`nAccuracy: Good for size - compact model`nUpdated: Jun 2024`nURL: ollama.com/library/gemma2:2b`n`nGoogle efficient small model. Good general knowledge with minimal resources."}
    [PSCustomObject]@{Name="qwen2.5:3b";  Size="~1.9 GB"; Quality="Multilingual";  MinRAM=6; Info="Alibaba Qwen 2.5 3B`n`nSpeed: Fast (15-30 sec)`nUse: Multilingual diagnostics, code analysis`nAccuracy: Good - strong at code/technical`nUpdated: Sep 2024`nURL: ollama.com/library/qwen2.5:3b`n`nStrong multilingual support. Good at analyzing scripts and technical output."}
    [PSCustomObject]@{Name="qwen2.5:7b";  Size="~4.7 GB"; Quality="Strong multi";  MinRAM=8; Info="Alibaba Qwen 2.5 7B`n`nSpeed: Medium (30-90 sec)`nUse: Detailed multilingual analysis`nAccuracy: High - excellent at code`nUpdated: Sep 2024`nURL: ollama.com/library/qwen2.5:7b`n`nLarger Qwen with significantly better analysis. Great for code-heavy diagnostics."}
    [PSCustomObject]@{Name="mistral";     Size="~4.1 GB"; Quality="Strong";        MinRAM=8; Info="Mistral 7B v0.3`n`nSpeed: Medium (30-90 sec)`nUse: Detailed analysis, complex issues`nAccuracy: High - catches subtle problems`nUpdated: May 2024`nURL: ollama.com/library/mistral`n`nHigh quality analysis with detailed explanations. Requires 8GB+ RAM."}
    [PSCustomObject]@{Name="llama3.1:8b"; Size="~4.7 GB"; Quality="Detailed";      MinRAM=10; Info="Meta Llama 3.1 8B`n`nSpeed: Slow (60-180 sec)`nUse: Deep analysis, comprehensive reports`nAccuracy: Highest - most thorough`nUpdated: Jul 2024`nURL: ollama.com/library/llama3.1:8b`n`nMost detailed and accurate. Best for thorough diagnostics. Needs 10GB+ RAM."}
    [PSCustomObject]@{Name="deepseek-r1:8b"; Size="~4.9 GB"; Quality="Reasoning";  MinRAM=10; Info="DeepSeek R1 8B`n`nSpeed: Slow (60-180 sec)`nUse: Complex reasoning, root cause analysis`nAccuracy: Very High - chain-of-thought`nUpdated: Jan 2025`nURL: ollama.com/library/deepseek-r1:8b`n`nExcels at step-by-step reasoning. Great for tracing complex issue chains."}
    [PSCustomObject]@{Name="granite3.2:8b"; Size="~4.9 GB"; Quality="Enterprise";  MinRAM=10; Info="IBM Granite 3.2 8B`n`nSpeed: Slow (60-180 sec)`nUse: Enterprise IT, structured output`nAccuracy: High - built for business`nUpdated: Mar 2025`nURL: ollama.com/library/granite3.2:8b`n`nIBM enterprise model. Strong at structured analysis and IT diagnostics."}
)
#endregion

#region Ollama Helpers (UI thread - lightweight)
function Test-OllamaInstalled {
    if(Get-Command "ollama" -EA SilentlyContinue){return $true}
    foreach($p in @("$env:LOCALAPPDATA\Programs\Ollama\ollama.exe","$env:ProgramFiles\Ollama\ollama.exe")){if(Test-Path $p){return $true}}; $false
}
function Test-OllamaRunning { try{$null=Invoke-RestMethod "$($Global:OllamaUrl)/api/tags" -Method Get -TimeoutSec 3 -EA Stop;$true}catch{$false} }
function Get-OllamaModels { try{$r=Invoke-RestMethod "$($Global:OllamaUrl)/api/tags" -Method Get -TimeoutSec 3 -EA Stop;@($r.models|ForEach-Object{$_.name})}catch{@()} }
function Get-OllamaStoragePath { $c=[Environment]::GetEnvironmentVariable("OLLAMA_MODELS","User"); if($c -and (Test-Path $c)){$c}else{"$env:USERPROFILE\.ollama\models"} }
function Get-OllamaStorageSize { $p=Get-OllamaStoragePath; if(Test-Path $p){[math]::Round((Get-ChildItem $p -Recurse -EA SilentlyContinue|Measure-Object Length -Sum).Sum/1GB,2)}else{0} }
#endregion


#region File Logger
function Write-LogFile {
    param([string]$Message, [string]$Level = "INFO")
    if(-not $Global:LogFilePath){ return }
    try {
        $logLine = "[$(Get-Date -F 'yyyy-MM-dd HH:mm:ss')][$Level] $Message"
        Add-Content -Path $Global:LogFilePath -Value $logLine -Encoding UTF8 -EA SilentlyContinue
    } catch {}
}
#endregion

#region Streaming Download Helper (shared scriptblock)
# Cancel flag - shared between UI and background thread
$Global:CancelDownload = [System.Collections.Concurrent.ConcurrentDictionary[string,bool]]::new()

$Global:StreamingPullScript = {
    function Q([string]$m){ $MsgQueue.Enqueue($m) }
    Q("[$(Get-Date -F 'yyyy-MM-dd HH:mm:ss')][INFO] Pull started for '$ModelName' from $OllamaUrl/api/pull")
    try {
        $body = "{`"name`":`"$ModelName`"}"
        $req = [System.Net.HttpWebRequest]::Create("$OllamaUrl/api/pull")
        $req.Method = "POST"; $req.ContentType = "application/json"; $req.Timeout = 1800000
        $bodyBytes = [System.Text.Encoding]::UTF8.GetBytes($body)
        $req.ContentLength = $bodyBytes.Length
        $reqStream = $req.GetRequestStream(); $reqStream.Write($bodyBytes, 0, $bodyBytes.Length); $reqStream.Close()

        $resp = $req.GetResponse(); $stream = $resp.GetResponseStream()
        $reader = [System.IO.StreamReader]::new($stream)
        $lastPct = -1; $lastLog = [DateTime]::MinValue; $startTime = Get-Date
        $cancelled = $false

        while(-not $reader.EndOfStream){
            # Check cancel flag
            $cancelVal = $false
            if($CancelFlag.TryGetValue("cancel", [ref]$cancelVal) -and $cancelVal){
                $cancelled = $true
                Q("[$(Get-Date -F 'yyyy-MM-dd HH:mm:ss')][WARN] Download cancelled by user")
                break
            }

            $line = $reader.ReadLine()
            if(-not $line){continue}
            try{
                $json = $line | ConvertFrom-Json -EA SilentlyContinue
                if($json.status -eq "success"){
                    $elapsed = (Get-Date) - $startTime
                    Q("[$(Get-Date -F 'yyyy-MM-dd HH:mm:ss')][OK] '$ModelName' downloaded! ($('{0:mm\:ss}' -f $elapsed))")
                }
                elseif($json.total -and $json.completed){
                    $pct = [math]::Round(($json.completed / $json.total) * 100, 0)
                    $now = Get-Date
                    if($pct -ne $lastPct -and ($now - $lastLog).TotalSeconds -ge 2){
                        $sizeMB = [math]::Round($json.total / 1MB, 0)
                        $dlMB = [math]::Round($json.completed / 1MB, 0)
                        $elapsed = ($now - $startTime).TotalSeconds
                        $speed = if($elapsed -gt 0){[math]::Round($json.completed / 1MB / $elapsed, 1)}else{0}
                        $digest = if($json.digest){$json.digest.Substring(0,[Math]::Min(19,$json.digest.Length))}else{""}
                        Q("[$(Get-Date -F 'yyyy-MM-dd HH:mm:ss')][INFO] $($json.status) $digest`: $pct% ($dlMB/$sizeMB MB) @ $speed MB/s")
                        Q("PROGRESS:$pct")
                        $lastPct = $pct; $lastLog = $now
                    }
                }
                elseif($json.status){
                    $now = Get-Date
                    if(($now - $lastLog).TotalSeconds -ge 3){
                        Q("[$(Get-Date -F 'yyyy-MM-dd HH:mm:ss')][INFO] $($json.status)")
                        $lastLog = $now
                    }
                }
            }catch{}
        }
        try{ $reader.Close() }catch{}
        try{ $resp.Close() }catch{}
        if($cancelled){
            Q("DONE:PULL:CANCELLED")
        } else {
            Q("DONE:PULL:OK")
        }
    } catch {
        Q("[$(Get-Date -F 'yyyy-MM-dd HH:mm:ss')][ERROR] Download failed: $($_.Exception.Message)")
        Q("DONE:PULL:FAIL")
    }
}
#endregion

#region WPF
Add-Type -AssemblyName PresentationFramework,PresentationCore,WindowsBase,System.Windows.Forms

[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="WinDiag-AI" Height="780" Width="1200" MinHeight="500" MinWidth="750" WindowStartupLocation="CenterScreen">
<Window.Resources>
    <Style x:Key="Btn" TargetType="Button">
        <Setter Property="Height" Value="28"/><Setter Property="Padding" Value="12,0"/><Setter Property="FontSize" Value="12"/><Setter Property="Cursor" Value="Hand"/>
        <Setter Property="Background" Value="Transparent"/><Setter Property="BorderBrush" Value="#CCC"/><Setter Property="BorderThickness" Value="1"/>
        <Setter Property="Template"><Setter.Value><ControlTemplate TargetType="Button">
            <Border x:Name="bd" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="5" Padding="{TemplateBinding Padding}">
                <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/></Border>
            <ControlTemplate.Triggers>
                <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="bd" Property="Background" Value="#F0F0F0"/></Trigger>
                <Trigger Property="IsPressed" Value="True"><Setter TargetName="bd" Property="Background" Value="#E0E0E0"/></Trigger>
                <Trigger Property="IsEnabled" Value="False"><Setter Property="Opacity" Value="0.45"/></Trigger>
            </ControlTemplate.Triggers></ControlTemplate></Setter.Value></Setter>
    </Style>
    <Style x:Key="RunBtn" TargetType="Button" BasedOn="{StaticResource Btn}">
        <Setter Property="Background" Value="#1E6EB5"/><Setter Property="Foreground" Value="White"/><Setter Property="BorderBrush" Value="#1558A0"/><Setter Property="FontWeight" Value="SemiBold"/>
        <Setter Property="Template"><Setter.Value><ControlTemplate TargetType="Button">
            <Border x:Name="bd" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="5" Padding="{TemplateBinding Padding}">
                <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/></Border>
            <ControlTemplate.Triggers>
                <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="bd" Property="Background" Value="#1558A0"/></Trigger>
                <Trigger Property="IsPressed" Value="True"><Setter TargetName="bd" Property="Background" Value="#0F4580"/></Trigger>
                <Trigger Property="IsEnabled" Value="False"><Setter Property="Opacity" Value="0.45"/></Trigger>
            </ControlTemplate.Triggers></ControlTemplate></Setter.Value></Setter>
    </Style>
    <Style x:Key="AiBtn" TargetType="Button" BasedOn="{StaticResource Btn}">
        <Setter Property="Background" Value="#7C3AED"/><Setter Property="Foreground" Value="White"/><Setter Property="BorderBrush" Value="#6D28D9"/><Setter Property="FontWeight" Value="SemiBold"/>
        <Setter Property="Template"><Setter.Value><ControlTemplate TargetType="Button">
            <Border x:Name="bd" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="5" Padding="{TemplateBinding Padding}">
                <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/></Border>
            <ControlTemplate.Triggers>
                <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="bd" Property="Background" Value="#6D28D9"/></Trigger>
                <Trigger Property="IsPressed" Value="True"><Setter TargetName="bd" Property="Background" Value="#5B21B6"/></Trigger>
                <Trigger Property="IsEnabled" Value="False"><Setter Property="Opacity" Value="0.45"/></Trigger>
            </ControlTemplate.Triggers></ControlTemplate></Setter.Value></Setter>
    </Style>
    <Style x:Key="TglBtn" TargetType="ToggleButton">
        <Setter Property="Height" Value="28"/><Setter Property="Padding" Value="12,0"/><Setter Property="FontSize" Value="12"/><Setter Property="Cursor" Value="Hand"/>
        <Setter Property="Background" Value="Transparent"/><Setter Property="BorderBrush" Value="#CCC"/><Setter Property="BorderThickness" Value="1"/>
        <Setter Property="Template"><Setter.Value><ControlTemplate TargetType="ToggleButton">
            <Border x:Name="bd" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="5" Padding="{TemplateBinding Padding}">
                <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/></Border>
            <ControlTemplate.Triggers>
                <Trigger Property="IsChecked" Value="True"><Setter TargetName="bd" Property="Background" Value="#E8F0FA"/><Setter TargetName="bd" Property="BorderBrush" Value="#1E6EB5"/></Trigger>
                <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="bd" Property="Background" Value="#F0F0F0"/></Trigger>
            </ControlTemplate.Triggers></ControlTemplate></Setter.Value></Setter>
    </Style>
</Window.Resources>
<Grid>
    <Grid.RowDefinitions>
        <RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="*"/><RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>
    <!-- Menu -->
    <Menu Grid.Row="0" Background="#F8F8F8" BorderBrush="#E8E8E8" BorderThickness="0,0,0,1">
        <MenuItem Header="_File">
            <MenuItem x:Name="mnuSettings" Header="_Settings"/>
            <MenuItem x:Name="mnuCustomChecks" Header="_Custom Checks"/>
            <Separator/>
            <MenuItem x:Name="mnuExportPdf" Header="Export to _PDF" IsEnabled="False"/>
            <Separator/>
            <MenuItem x:Name="mnuExit" Header="E_xit"/>
        </MenuItem>
        <MenuItem Header="_Info">
            <MenuItem x:Name="mnuAbout" Header="_About"/>
        </MenuItem>
    </Menu>
    <!-- Header -->
    <Border Grid.Row="1" Background="#F8F8F8" BorderBrush="#E8E8E8" BorderThickness="0,0,0,1" Padding="12,8">
        <DockPanel>
            <StackPanel DockPanel.Dock="Left" VerticalAlignment="Center">
                <TextBlock FontSize="16" FontWeight="SemiBold" Text="WinDiag-AI" Foreground="#1A1A1A"/>
                <TextBlock FontSize="11" Foreground="#888" Text="Windows Diagnostics with AI Analysis"/>
            </StackPanel>
            <StackPanel DockPanel.Dock="Right" Orientation="Horizontal" VerticalAlignment="Center" HorizontalAlignment="Right">
                <ComboBox x:Name="cmbTheme" Width="75" Height="22" FontSize="10" VerticalContentAlignment="Center" Margin="0,0,12,0">
                    <ComboBox.ToolTip>
                        <ToolTip><TextBlock Text="UI Theme: Auto follows Windows setting" TextWrapping="Wrap"/></ToolTip>
                    </ComboBox.ToolTip>
                </ComboBox>
                <TextBlock x:Name="txtOllamaStatus" VerticalAlignment="Center" FontSize="11" Foreground="#888"/>
            </StackPanel>
            <Rectangle/>
        </DockPanel>
    </Border>
    <!-- Toolbar -->
    <Border Grid.Row="2" Background="#FAFAFA" BorderBrush="#E8E8E8" BorderThickness="0,0,0,1" Padding="12,6">
        <DockPanel>
            <Button x:Name="btnScan" DockPanel.Dock="Left" Style="{StaticResource RunBtn}" Content="&#x1F50D; Scan System" Padding="14,0" Margin="0,0,8,0"/>
            <Button x:Name="btnAI" DockPanel.Dock="Left" Style="{StaticResource AiBtn}" Content="&#x1F916; AI Analysis" Padding="14,0" Margin="0,0,8,0" IsEnabled="False"/>
            <Button x:Name="btnReport" DockPanel.Dock="Left" Style="{StaticResource Btn}" Content="&#x1F4C4; Save Report" Padding="14,0" Margin="0,0,8,0" IsEnabled="False"/>
            <Button x:Name="btnOllama" DockPanel.Dock="Left" Style="{StaticResource Btn}" Content="&#x2699; Ollama" Padding="14,0" Margin="0,0,8,0"/>
            <Button x:Name="btnCleanup" DockPanel.Dock="Left" Style="{StaticResource Btn}" Content="&#x1F5D1; Cleanup" Padding="14,0" Margin="0,0,8,0" IsEnabled="False"/>
            <ToggleButton x:Name="tglVerbose" DockPanel.Dock="Left" Style="{StaticResource TglBtn}" Content="Verbose" Margin="0,0,8,0"/>
            <Button x:Name="btnClearLog" DockPanel.Dock="Right" Style="{StaticResource Btn}" Content="Clear log"/>
            <TextBlock x:Name="txtStatus" DockPanel.Dock="Right" VerticalAlignment="Center" FontSize="12" Foreground="#555" Margin="0,0,12,0" HorizontalAlignment="Right"/>
        </DockPanel>
    </Border>
    <!-- Model bar -->
    <Border Grid.Row="3" Background="#F2F0FF" BorderBrush="#E0DCFF" BorderThickness="0,0,0,1" Padding="12,6">
        <DockPanel>
            <TextBlock DockPanel.Dock="Left" Text="AI Model:" VerticalAlignment="Center" FontSize="12" FontWeight="SemiBold" Foreground="#5B21B6" Margin="0,0,8,0"/>
            <ComboBox x:Name="cmbModel" DockPanel.Dock="Left" Width="280" Height="26" FontSize="12" VerticalContentAlignment="Center"/>
            <TextBlock x:Name="txtModelTip" DockPanel.Dock="Left" VerticalAlignment="Center" FontSize="13" Foreground="#7C3AED" Margin="6,0,0,0" Cursor="Help" Text="&#x24D8;">
                <TextBlock.ToolTip>
                    <ToolTip MaxWidth="500" Placement="Bottom">
                        <TextBlock x:Name="txtModelTipContent" TextWrapping="Wrap" FontFamily="Consolas" FontSize="11.5" Text="Select a model to see info"/>
                    </ToolTip>
                </TextBlock.ToolTip>
            </TextBlock>
            <Button x:Name="btnPull" DockPanel.Dock="Left" Style="{StaticResource Btn}" Content="&#x2B07; Download" Padding="10,0" Margin="8,0,0,0" Height="26"/>
            <TextBlock x:Name="txtModelInfo" DockPanel.Dock="Right" VerticalAlignment="Center" FontSize="11" Foreground="#7C3AED" HorizontalAlignment="Right" Cursor="Hand" TextDecorations="Underline"/>
        </DockPanel>
    </Border>
    <!-- Content -->
    <Grid Grid.Row="4" Margin="12,10,12,0">
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="250" MinWidth="150"/><ColumnDefinition Width="Auto"/><ColumnDefinition Width="*" MinWidth="200"/>
        </Grid.ColumnDefinitions>
        <!-- Left: checks -->
        <Border Grid.Column="0" BorderBrush="#DCDCDC" BorderThickness="1" CornerRadius="8">
            <DockPanel>
                <Border DockPanel.Dock="Top" Background="#F2F4F7" Padding="10,6" BorderBrush="#ECECEC" BorderThickness="0,0,0,1">
                    <DockPanel>
                        <TextBlock DockPanel.Dock="Left" Text="Diagnostic Checks" FontSize="12" FontWeight="SemiBold" Foreground="#555" VerticalAlignment="Center"/>
                        <StackPanel DockPanel.Dock="Right" Orientation="Horizontal" HorizontalAlignment="Right">
                            <Button x:Name="btnCheckAll" Content="All" FontSize="10" Padding="6,1" Margin="0,0,4,0" Cursor="Hand" Background="Transparent" BorderBrush="#CCC" BorderThickness="1"/>
                            <Button x:Name="btnUncheckAll" Content="None" FontSize="10" Padding="6,1" Cursor="Hand" Background="Transparent" BorderBrush="#CCC" BorderThickness="1"/>
                        </StackPanel>
                    </DockPanel>
                </Border>
                <ScrollViewer VerticalScrollBarVisibility="Auto" Padding="8,4">
                    <StackPanel x:Name="pnlChecks">
                        <CheckBox x:Name="chkSysInfo"   Content="System Information"       IsChecked="True" Margin="0,3" FontSize="11.5"/>
                        <CheckBox x:Name="chkHardware"   Content="Hardware Details"          IsChecked="True" Margin="0,3" FontSize="11.5"/>
                        <CheckBox x:Name="chkBattery"    Content="Battery Info"              IsChecked="True" Margin="0,3" FontSize="11.5"/>
                        <CheckBox x:Name="chkDisks"      Content="Disk Health + S.M.A.R.T"   IsChecked="True" Margin="0,3" FontSize="11.5"/>
                        <CheckBox x:Name="chkSmartDetails" Content="  &gt; SMART Details"    IsChecked="True" Margin="16,3,0,3" FontSize="11.5"/>
                        <CheckBox x:Name="chkEvents"     Content="Event Logs"                IsChecked="True" Margin="0,3" FontSize="11.5"/>
                        <CheckBox x:Name="chkServices"   Content="Services (Stopped Auto)"   IsChecked="True" Margin="0,3" FontSize="11.5"/>
                        <CheckBox x:Name="chkNetwork"    Content="Network + DNS"             IsChecked="True" Margin="0,3" FontSize="11.5"/>
                        <CheckBox x:Name="chkSecurity"   Content="Security Status"           IsChecked="True" Margin="0,3" FontSize="11.5"/>
                        <CheckBox x:Name="chkUpdates"    Content="Windows Updates"           IsChecked="True" Margin="0,3" FontSize="11.5"/>
                        <CheckBox x:Name="chkProcesses"  Content="Top Processes"             IsChecked="True" Margin="0,3" FontSize="11.5"/>
                        <CheckBox x:Name="chkStartup"    Content="Startup Programs"          IsChecked="True" Margin="0,3" FontSize="11.5"/>
                        <CheckBox x:Name="chkScheduled"  Content="Scheduled Tasks"           IsChecked="True" Margin="0,3" FontSize="11.5"/>
                        <CheckBox x:Name="chkAutoruns"   Content="Autoruns (Registry/Shell)"  IsChecked="True" Margin="0,3" FontSize="11.5"/>
                        <CheckBox x:Name="chkIntegrity"  Content="System Integrity (SFC)"    IsChecked="True" Margin="0,3" FontSize="11.5"/>
                        <CheckBox x:Name="chkIndexing"   Content="Windows Search/Indexing"   IsChecked="True" Margin="0,3" FontSize="11.5"/>
                        <CheckBox x:Name="chkSoftware"   Content="Installed Software"        IsChecked="True" Margin="0,3" FontSize="11.5"/>
                        <CheckBox x:Name="chkUserInfo"   Content="User Accounts"             IsChecked="True" Margin="0,3" FontSize="11.5"/>
                        <CheckBox x:Name="chkHosts"      Content="Hosts File"                IsChecked="True" Margin="0,3" FontSize="11.5"/>
                        <CheckBox x:Name="chkExternalIP" Content="External IP"               IsChecked="True" Margin="0,3" FontSize="11.5"/>
                        <CheckBox x:Name="chkHotfixes"   Content="Installed Hotfixes"        IsChecked="True" Margin="0,3" FontSize="11.5"/>
                        <CheckBox x:Name="chkRemoteTools" Content="Remote Tools"             IsChecked="True" Margin="0,3" FontSize="11.5"/>
                        <CheckBox x:Name="chkChkdskLogs" Content="Chkdsk Logs"               IsChecked="True" Margin="0,3" FontSize="11.5"/>
                        <CheckBox x:Name="chkBatteryReport" Content="Battery Report (powercfg)" IsChecked="True" Margin="0,3" FontSize="11.5"/>
                        <CheckBox x:Name="chkDrivers"   Content="Driver Check"              IsChecked="True" Margin="0,3" FontSize="11.5"/>
                        <CheckBox x:Name="chkBSOD"      Content="BSOD / Crash Logs"         IsChecked="True" Margin="0,3" FontSize="11.5"/>
                        <CheckBox x:Name="chkRAMTest"    Content="RAM Test Logs (mdsched)"   IsChecked="True" Margin="0,3" FontSize="11.5"/>
                        <CheckBox x:Name="chkPerfmon"    Content="Performance Report"        IsChecked="False" Margin="0,3" FontSize="11.5"/>
                    </StackPanel>
                </ScrollViewer>
            </DockPanel>
        </Border>
        <GridSplitter Grid.Column="1" Width="7" HorizontalAlignment="Center" VerticalAlignment="Stretch" Background="Transparent" Margin="2,0" ResizeBehavior="PreviousAndNext" ResizeDirection="Columns" Cursor="SizeWE"/>
        <!-- Right: output -->
        <Grid Grid.Column="2">
            <Grid.RowDefinitions><RowDefinition Height="2*" MinHeight="80"/><RowDefinition Height="Auto"/><RowDefinition Height="*" MinHeight="60"/></Grid.RowDefinitions>
            <Border Grid.Row="0" BorderBrush="#DCDCDC" BorderThickness="1" CornerRadius="8">
                <RichTextBox x:Name="txtLog" FontFamily="Consolas" FontSize="11" IsReadOnly="True" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled" Padding="10,6" BorderThickness="0" Background="Transparent"><FlowDocument/></RichTextBox>
            </Border>
            <GridSplitter Grid.Row="1" Height="7" HorizontalAlignment="Stretch" Background="Transparent" Margin="0,2" ResizeBehavior="PreviousAndNext" ResizeDirection="Rows" Cursor="SizeNS"/>
            <Border Grid.Row="2" BorderBrush="#DCDCDC" BorderThickness="1" CornerRadius="8">
                <Grid><Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/></Grid.RowDefinitions>
                    <Border Grid.Row="0" Background="#F8F8F8" BorderBrush="#ECECEC" BorderThickness="0,0,0,1" Padding="10,4" CornerRadius="8,8,0,0">
                        <TextBlock Text="AI Output" FontSize="10" FontWeight="SemiBold" Foreground="#888"/>
                    </Border>
                    <RichTextBox Grid.Row="1" x:Name="txtAI" FontFamily="Consolas" FontSize="11" IsReadOnly="True" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled" Padding="10,6" BorderThickness="0" Background="Transparent"><FlowDocument/></RichTextBox>
                </Grid>
            </Border>
        </Grid>
    </Grid>
    <TextBlock Grid.Row="5" Margin="12,6,12,8" Foreground="#AAA" FontSize="11" Text="WinDiag-AI  |  Powered by Ollama  |  karanik.gr"/>
</Grid>
</Window>
"@

$rd = New-Object System.Xml.XmlNodeReader $xaml
$W  = [Windows.Markup.XamlReader]::Load($rd)

# Controls
$txtLog    = $W.FindName("txtLog");    $txtAI     = $W.FindName("txtAI")
$txtStatus = $W.FindName("txtStatus"); $txtOSt    = $W.FindName("txtOllamaStatus"); $txtMI = $W.FindName("txtModelInfo")
$btnScan   = $W.FindName("btnScan");   $btnAI     = $W.FindName("btnAI");     $btnReport = $W.FindName("btnReport")
$btnOllama = $W.FindName("btnOllama"); $btnCleanup= $W.FindName("btnCleanup"); $btnClearLog=$W.FindName("btnClearLog")
$btnPull   = $W.FindName("btnPull");   $cmbModel  = $W.FindName("cmbModel");  $tglVerbose= $W.FindName("tglVerbose")
$mnuExit   = $W.FindName("mnuExit");  $mnuAbout  = $W.FindName("mnuAbout"); $mnuSettings = $W.FindName("mnuSettings")
$mnuCustomChecks = $W.FindName("mnuCustomChecks"); $mnuExportPdf = $W.FindName("mnuExportPdf")
$cmbTheme = $W.FindName("cmbTheme")
$txtModelTip = $W.FindName("txtModelTip"); $txtModelTipContent = $W.FindName("txtModelTipContent")
$btnCheckAll = $W.FindName("btnCheckAll"); $btnUncheckAll = $W.FindName("btnUncheckAll")
$chks = @{}
foreach($n in @("SysInfo","Hardware","Battery","Disks","SmartDetails","Events","Services","Network","Security","Updates","Processes","Startup","Scheduled","Autoruns","Integrity","Indexing","Software","Hosts","UserInfo","ExternalIP","Hotfixes","RemoteTools","ChkdskLogs","BatteryReport","Drivers","BSOD","RAMTest","Perfmon")){
    $chks[$n] = $W.FindName("chk$n")
}
# Check All / Uncheck All
$btnCheckAll.Add_Click({ foreach($c in $chks.Values){ $c.IsChecked = $true } })
$btnUncheckAll.Add_Click({ foreach($c in $chks.Values){ $c.IsChecked = $false } })
# Verbose ON by default
$tglVerbose.IsChecked = $true; $Script:Verbose = $true
# Theme selector
foreach($t in @("Auto","Light","Dark")){ $cmbTheme.Items.Add($t) | Out-Null }
$cmbTheme.SelectedIndex = 0
$cmbTheme.Add_SelectionChanged({
    $theme = $cmbTheme.SelectedItem
    if(-not $theme){ return }
    $applyDark = $false
    if($theme -eq "Auto"){
        try {
            $regVal = (Get-ItemProperty "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize" -Name AppsUseLightTheme -EA SilentlyContinue).AppsUseLightTheme
            $applyDark = ($regVal -eq 0)
        } catch { $applyDark = $false }
    } elseif($theme -eq "Dark"){ $applyDark = $true }

    $Global:IsDarkTheme = $applyDark

    $bc = [Windows.Media.BrushConverter]::new()
    if($applyDark){
        $bg1 = $bc.ConvertFrom("#1E1E2E");  $bg2 = $bc.ConvertFrom("#181825")
        $bg3 = $bc.ConvertFrom("#252535");   $bgLog = $bc.ConvertFrom("#11111B")
        $bgModel = $bc.ConvertFrom("#2A2040"); $fg = $bc.ConvertFrom("#CDD6F4")
        $fgDim = $bc.ConvertFrom("#A6ADC8"); $fgAccent = $bc.ConvertFrom("#B4BEFE")
        $brd = $bc.ConvertFrom("#45475A");   $btnBg = $bc.ConvertFrom("#313244")
        $btnFg = $bc.ConvertFrom("#CDD6F4"); $chkFg = $bc.ConvertFrom("#CDD6F4")
    } else {
        $bg1 = [Windows.Media.Brushes]::White; $bg2 = $bc.ConvertFrom("#F8F8F8")
        $bg3 = $bc.ConvertFrom("#F2F4F7");   $bgLog = [Windows.Media.Brushes]::White
        $bgModel = $bc.ConvertFrom("#F2F0FF"); $fg = $bc.ConvertFrom("#1A1A1A")
        $fgDim = $bc.ConvertFrom("#888888"); $fgAccent = $bc.ConvertFrom("#5B21B6")
        $brd = $bc.ConvertFrom("#DCDCDC");   $btnBg = [Windows.Media.Brushes]::Transparent
        $btnFg = $bc.ConvertFrom("#1A1A1A"); $chkFg = $bc.ConvertFrom("#1A1A1A")
    }

    # Recursive function to theme all children
    $applyToVisual = {
        param($element)
        if($null -eq $element){ return }

        # Apply foreground to text-bearing controls
        if($element -is [Windows.Controls.TextBlock]){ $element.Foreground = $fg }
        if($element -is [Windows.Controls.CheckBox]){ $element.Foreground = $chkFg }
        if($element -is [Windows.Controls.Label]){ $element.Foreground = $fg }
        if($element -is [Windows.Controls.MenuItem]){ $element.Foreground = $fg }

        # Recurse into children
        if($element -is [Windows.Controls.Panel]){
            foreach($ch in $element.Children){ & $applyToVisual $ch }
        }
        if($element -is [Windows.Controls.ContentControl] -and $element.Content -is [Windows.UIElement]){
            & $applyToVisual $element.Content
        }
        if($element -is [Windows.Controls.Decorator] -and $element.Child){
            & $applyToVisual $element.Child
        }
        if($element -is [Windows.Controls.ItemsControl]){
            foreach($item in $element.Items){
                if($item -is [Windows.UIElement]){ & $applyToVisual $item }
            }
        }
    }

    # Main window
    $W.Background = $bg1

    # Apply to all grid children by row
    $grid = $W.Content
    foreach($child in $grid.Children){
        $row = [Windows.Controls.Grid]::GetRow($child)
        if($child -is [Windows.Controls.Menu]){
            $child.Background = $bg2; $child.Foreground = $fg
            foreach($mi in $child.Items){ if($mi -is [Windows.Controls.MenuItem]){ $mi.Foreground = $fg } }
        }
        if($child -is [Windows.Controls.Border]){
            $child.BorderBrush = $brd
            switch($row){
                1 { $child.Background = $bg2 }
                2 { $child.Background = if($applyDark){$bg1}else{$bc.ConvertFrom("#FAFAFA")} }
                3 { $child.Background = $bgModel }
            }
            # Recurse into border children
            & $applyToVisual $child.Child
        }
        if($child -is [Windows.Controls.TextBlock]){
            if($row -eq 5){ $child.Foreground = $fgDim }
        }
        if($child -is [Windows.Controls.Grid] -and $row -eq 4){
            # Content area - left panel + right log/AI
            foreach($sub in $child.Children){
                if($sub -is [Windows.Controls.Border]){
                    $sub.BorderBrush = $brd
                    $sub.Background = if($applyDark){$bg1}else{[Windows.Media.Brushes]::Transparent}
                    & $applyToVisual $sub.Child
                }
                if($sub -is [Windows.Controls.Grid]){
                    foreach($ssub in $sub.Children){
                        if($ssub -is [Windows.Controls.Border]){
                            $ssub.BorderBrush = $brd
                            & $applyToVisual $ssub.Child
                        }
                    }
                }
            }
        }
    }
    # Named controls - explicit override
    $txtLog.Background = $bgLog; $txtLog.Foreground = $fg
    $txtAI.Background = $bgLog; $txtAI.Foreground = $fg
    $txtStatus.Foreground = if($applyDark){$fgDim}else{$bc.ConvertFrom("#555555")}
    $txtOSt.Foreground = $txtOSt.Foreground # keep Ollama status color
    $txtMI.Foreground = $fgAccent
    # Checkboxes in checks panel
    foreach($c in $chks.Values){ $c.Foreground = $chkFg }

    # Toolbar buttons - override foreground and background
    $toolbarBtns = @($btnScan, $btnAI, $btnReport, $btnOllama, $btnCleanup, $btnClearLog, $btnPull)
    foreach($b in $toolbarBtns){
        if($b -eq $btnScan -or $b -eq $btnAI){ continue } # keep styled colors for Scan/AI
        $b.Foreground = $btnFg
        $b.Background = $btnBg
        $b.BorderBrush = $brd
    }
    # Verbose toggle
    $tglVerbose.Foreground = $btnFg
    $tglVerbose.BorderBrush = $brd

    # Menu bar and items - WPF menus need explicit foreground
    $menu = $W.Content.Children | Where-Object { $_ -is [Windows.Controls.Menu] } | Select-Object -First 1
    if($menu){
        $menu.Background = $bg2
        $menu.Foreground = $fg
        # Recursively set menu item colors
        $setMenuColors = {
            param($items, $fgColor, $bgColor)
            foreach($mi in $items){
                if($mi -is [Windows.Controls.MenuItem]){
                    $mi.Foreground = $fgColor
                    $mi.Background = $bgColor
                    if($mi.Items.Count -gt 0){ & $setMenuColors $mi.Items $fgColor $bgColor }
                }
            }
        }
        & $setMenuColors $menu.Items $fg $(if($applyDark){$bg2}else{$bc.ConvertFrom("#F8F8F8")})
    }

    # ComboBoxes
    $cmbModel.Foreground = $fg
    $cmbModel.Background = if($applyDark){$btnBg}else{[Windows.Media.Brushes]::White}
    $cmbTheme.Foreground = $fg
    $cmbTheme.Background = if($applyDark){$btnBg}else{[Windows.Media.Brushes]::White}

    # Model bar text
    $txtModelTip.Foreground = $fgAccent

    # Header texts
    $grid = $W.Content
    foreach($child in $grid.Children){
        if($child -is [Windows.Controls.Border]){
            $row = [Windows.Controls.Grid]::GetRow($child)
            if($row -eq 1 -and $child.Child -is [Windows.Controls.DockPanel]){
                foreach($dp in $child.Child.Children){
                    if($dp -is [Windows.Controls.StackPanel]){
                        foreach($tb in $dp.Children){
                            if($tb -is [Windows.Controls.TextBlock]){
                                if($tb.FontWeight -eq [Windows.FontWeights]::SemiBold){ $tb.Foreground = $fg }
                                else{ $tb.Foreground = $fgDim }
                            }
                        }
                    }
                }
            }
            # Model bar texts
            if($row -eq 3 -and $child.Child -is [Windows.Controls.DockPanel]){
                foreach($dp in $child.Child.Children){
                    if($dp -is [Windows.Controls.TextBlock]){ $dp.Foreground = $fgAccent }
                }
            }
        }
    }

    # AI Output header
    foreach($child in $grid.Children){
        if($child -is [Windows.Controls.Grid]){
            $row = [Windows.Controls.Grid]::GetRow($child)
            if($row -eq 4){
                # Find AI Output header border
                foreach($sub in $child.Children){
                    if($sub -is [Windows.Controls.Grid]){
                        foreach($ssub in $sub.Children){
                            if($ssub -is [Windows.Controls.Border] -and $ssub.Child -is [Windows.Controls.Grid]){
                                foreach($gs in $ssub.Child.Children){
                                    if($gs -is [Windows.Controls.Border]){
                                        $gs.Background = if($applyDark){$bc.ConvertFrom("#252535")}else{$bc.ConvertFrom("#F8F8F8")}
                                        $gs.BorderBrush = $brd
                                        if($gs.Child -is [Windows.Controls.TextBlock]){ $gs.Child.Foreground = $fgDim }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
    }

    # Checks panel header
    foreach($child in $grid.Children){
        if($child -is [Windows.Controls.Grid]){
            $row = [Windows.Controls.Grid]::GetRow($child)
            if($row -eq 4){
                foreach($sub in $child.Children){
                    if($sub -is [Windows.Controls.Border] -and $sub.Child -is [Windows.Controls.DockPanel]){
                        foreach($dp in $sub.Child.Children){
                            if($dp -is [Windows.Controls.Border]){
                                $dp.Background = if($applyDark){$bc.ConvertFrom("#252535")}else{$bc.ConvertFrom("#F2F4F7")}
                                $dp.BorderBrush = $brd
                                & $applyToVisual $dp.Child
                            }
                            if($dp -is [Windows.Controls.ScrollViewer]){
                                $dp.Background = if($applyDark){$bg1}else{[Windows.Media.Brushes]::White}
                            }
                        }
                    }
                }
            }
        }
    }

    # All/None buttons in checks panel
    $btnCheckAll.Foreground = $btnFg; $btnCheckAll.Background = $btnBg; $btnCheckAll.BorderBrush = $brd
    $btnUncheckAll.Foreground = $btnFg; $btnUncheckAll.Background = $btnBg; $btnUncheckAll.BorderBrush = $brd

    # Footer text
    $footer = $grid.Children | Where-Object { $_ -is [Windows.Controls.TextBlock] -and [Windows.Controls.Grid]::GetRow($_) -eq 5 } | Select-Object -First 1
    if($footer){ $footer.Foreground = $fgDim }

    Write-Host "[Theme] Applied: $theme $(if($applyDark){'(dark)'}else{'(light)'})" -ForegroundColor Cyan
})
# Helper: apply current theme to a dialog window
function Apply-DialogTheme([Windows.Window]$dlg){
    if($Global:IsDarkTheme){
        $bc = [Windows.Media.BrushConverter]::new()
        $dlg.Background = $bc.ConvertFrom("#1E1E2E")
        $dlg.Foreground = $bc.ConvertFrom("#CDD6F4")
        # Recurse all children
        $themeChild = {
            param($el)
            if($null -eq $el){ return }
            if($el -is [Windows.Controls.TextBlock]){ $el.Foreground = $bc.ConvertFrom("#CDD6F4") }
            if($el -is [Windows.Controls.Label]){ $el.Foreground = $bc.ConvertFrom("#CDD6F4") }
            if($el -is [Windows.Controls.CheckBox]){ $el.Foreground = $bc.ConvertFrom("#CDD6F4") }
            if($el -is [Windows.Controls.Button]){ $el.Foreground = $bc.ConvertFrom("#CDD6F4"); $el.Background = $bc.ConvertFrom("#313244"); $el.BorderBrush = $bc.ConvertFrom("#45475A") }
            if($el -is [Windows.Controls.TextBox]){ $el.Background = $bc.ConvertFrom("#11111B"); $el.Foreground = $bc.ConvertFrom("#CDD6F4"); $el.BorderBrush = $bc.ConvertFrom("#45475A") }
            if($el -is [Windows.Controls.ComboBox]){ $el.Background = $bc.ConvertFrom("#313244"); $el.Foreground = $bc.ConvertFrom("#CDD6F4"); $el.BorderBrush = $bc.ConvertFrom("#45475A") }
            if($el -is [Windows.Controls.Slider]){ $el.Foreground = $bc.ConvertFrom("#CDD6F4") }
            if($el -is [Windows.Controls.TabControl]){ $el.Background = $bc.ConvertFrom("#1E1E2E") }
            if($el -is [Windows.Controls.TabItem]){ $el.Background = $bc.ConvertFrom("#252535"); $el.Foreground = $bc.ConvertFrom("#CDD6F4") }
            if($el -is [Windows.Controls.ListView]){ $el.Background = $bc.ConvertFrom("#11111B"); $el.Foreground = $bc.ConvertFrom("#CDD6F4"); $el.BorderBrush = $bc.ConvertFrom("#45475A") }
            if($el -is [Windows.Controls.ScrollViewer]){ $el.Background = $bc.ConvertFrom("#1E1E2E") }
            if($el -is [Windows.Controls.Panel]){ foreach($ch in $el.Children){ & $themeChild $ch } }
            if($el -is [Windows.Controls.ContentControl] -and $el.Content -is [Windows.UIElement]){ & $themeChild $el.Content }
            if($el -is [Windows.Controls.Decorator] -and $el.Child){ & $themeChild $el.Child }
            if($el -is [Windows.Controls.ItemsControl]){
                foreach($item in $el.Items){ if($item -is [Windows.UIElement]){ & $themeChild $item } }
            }
        }
        if($dlg.Content -is [Windows.UIElement]){ & $themeChild $dlg.Content }
    }
}
# Click on storage path opens the folder
$txtMI.Add_MouseLeftButtonUp({ $p = Get-OllamaStoragePath; if(Test-Path $p){Start-Process "explorer.exe" -ArgumentList $p} })
#endregion

#region UI Helpers
function Apply-DialogTheme([Windows.Window]$Dialog) {
    if(-not $Global:IsDarkTheme){ return }
    $bc = [Windows.Media.BrushConverter]::new()
    $Dialog.Background = $bc.ConvertFrom("#1E1E2E")
    $Dialog.Foreground = $bc.ConvertFrom("#CDD6F4")
    # Recursively theme all children
    $themeChild = {
        param($el)
        if($null -eq $el){ return }
        if($el -is [Windows.Controls.TextBlock]){ $el.Foreground = $bc.ConvertFrom("#CDD6F4") }
        if($el -is [Windows.Controls.TextBox]){ $el.Background = $bc.ConvertFrom("#313244"); $el.Foreground = $bc.ConvertFrom("#CDD6F4"); $el.BorderBrush = $bc.ConvertFrom("#45475A") }
        if($el -is [Windows.Controls.Button]){ $el.Background = $bc.ConvertFrom("#313244"); $el.Foreground = $bc.ConvertFrom("#CDD6F4"); $el.BorderBrush = $bc.ConvertFrom("#45475A") }
        if($el -is [Windows.Controls.CheckBox]){ $el.Foreground = $bc.ConvertFrom("#CDD6F4") }
        if($el -is [Windows.Controls.ComboBox]){ $el.Background = $bc.ConvertFrom("#313244"); $el.Foreground = $bc.ConvertFrom("#CDD6F4") }
        if($el -is [Windows.Controls.Slider]){ $el.Foreground = $bc.ConvertFrom("#CDD6F4") }
        if($el -is [Windows.Controls.TabControl]){ $el.Background = $bc.ConvertFrom("#1E1E2E") }
        if($el -is [Windows.Controls.TabItem]){ $el.Background = $bc.ConvertFrom("#252535"); $el.Foreground = $bc.ConvertFrom("#CDD6F4") }
        if($el -is [Windows.Controls.ListView]){ $el.Background = $bc.ConvertFrom("#11111B"); $el.Foreground = $bc.ConvertFrom("#CDD6F4") }
        if($el -is [Windows.Controls.Label]){ $el.Foreground = $bc.ConvertFrom("#CDD6F4") }
        if($el -is [Windows.Controls.ScrollViewer]){ $el.Background = $bc.ConvertFrom("#1E1E2E") }
        # Recurse
        if($el -is [Windows.Controls.Panel]){ foreach($ch in $el.Children){ & $themeChild $ch } }
        if($el -is [Windows.Controls.ContentControl] -and $el.Content -is [Windows.UIElement]){ & $themeChild $el.Content }
        if($el -is [Windows.Controls.Decorator] -and $el.Child){ & $themeChild $el.Child }
        if($el -is [Windows.Controls.ItemsControl]){ foreach($item in $el.Items){ if($item -is [Windows.UIElement]){ & $themeChild $item } } }
    }
    & $themeChild $Dialog.Content
}
function Ui-Append([string]$Txt,[string]$Clr=$null,[System.Windows.Controls.RichTextBox]$Target=$txtLog) {
    if(-not $Clr){
        if($Txt-match'\[ERROR\]'){$Clr="Red"}elseif($Txt-match'\[WARN\]'){$Clr="DarkOrange"}
        elseif($Txt-match'\[OK\]'){$Clr="Green"}elseif($Txt-match'\[AI\]'){$Clr="MediumOrchid"}
        elseif($Txt-match'\[SCAN\]'){$Clr="DodgerBlue"}elseif($Txt-match'\[DEBUG\]'){$Clr="MediumPurple"}
        elseif($Txt-match'\[INFO\]'){$Clr="CornflowerBlue"}
    }
    $p=$Target.Document.Blocks|Select-Object -Last 1
    if(-not $p -or $p -isnot [Windows.Documents.Paragraph]){$p=[Windows.Documents.Paragraph]::new();$p.Margin=[Windows.Thickness]::new(0);$Target.Document.Blocks.Add($p)}
    $r=[Windows.Documents.Run]::new($Txt+[Environment]::NewLine)
    if($Clr){try{$r.Foreground=[Windows.Media.Brushes]::$Clr}catch{}}
    $p.Inlines.Add($r); $Target.ScrollToEnd()
    # Console output (visible in PS window behind GUI)
    $hostClr = switch($Clr){ "Red"{"Red"} "DarkOrange"{"Yellow"} "Green"{"Green"} "MediumOrchid"{"Magenta"} "DodgerBlue"{"Cyan"} "MediumPurple"{"DarkMagenta"} "CornflowerBlue"{"Gray"} default{"White"} }
    Write-Host $Txt -ForegroundColor $hostClr
    # Also write to log file
    Write-LogFile $Txt
}
function Ui-Log([string]$m,[string]$l="INFO"){Ui-Append "[$(Get-Date -F 'yyyy-MM-dd HH:mm:ss')][$l] $m"}
function Ui-SetStatus([string]$s){$txtStatus.Text=$s}
function Set-UiBusy([bool]$b){
    $btnScan.IsEnabled=-not $b; $btnAI.IsEnabled=(-not $b)-and($null-ne $Global:DiagData)
    $btnReport.IsEnabled=(-not $b)-and($null-ne $Global:DiagData); $btnOllama.IsEnabled=-not $b
    $btnPull.IsEnabled=-not $b; $btnCleanup.IsEnabled=(-not $b)-and((Get-OllamaStorageSize)-gt 0)
    $mnuExportPdf.IsEnabled=(-not $b)-and($null-ne $Global:DiagData)
}
function Refresh-Ollama {
    $inst=Test-OllamaInstalled; $run=if($inst){Test-OllamaRunning}else{$false}
    if(-not $inst){$txtOSt.Text="Ollama: NOT INSTALLED";$txtOSt.Foreground=[Windows.Media.Brushes]::Red}
    elseif(-not $run){$txtOSt.Text="Ollama: NOT RUNNING";$txtOSt.Foreground=[Windows.Media.Brushes]::DarkOrange}
    else{$txtOSt.Text="Ollama: Running";$txtOSt.Foreground=[Windows.Media.Brushes]::Green}
    $cmbModel.Items.Clear()
    if($run){
        $ms=Get-OllamaModels
        if($ms.Count-gt 0){foreach($m in $ms){$cmbModel.Items.Add($m)|Out-Null}; $cmbModel.Items.Add("---")|Out-Null}
        foreach($rm in $RecommendedModels){if(-not($ms|Where-Object{$_-like "$($rm.Name)*"})){$cmbModel.Items.Add("[DL] $($rm.Name) ($($rm.Size))")|Out-Null}}
        $cmbModel.Items.Add("[DL] Custom...")|Out-Null
        if($ms.Count-gt 0){$cmbModel.SelectedIndex=0; Ui-Log "Ollama models: $($ms -join ', ')" "INFO"}
    }else{foreach($rm in $RecommendedModels){$cmbModel.Items.Add("[DL] $($rm.Name) ($($rm.Size))")|Out-Null}}
    $txtMI.Text="Storage: $(Get-OllamaStoragePath) ($(Get-OllamaStorageSize) GB)"
    $btnCleanup.IsEnabled=(Get-OllamaStorageSize)-gt 0
}
function Get-SelModel {
    $s=$cmbModel.SelectedItem; if(-not $s -or $s-eq "---"){return ""}
    if($s -match '^\[DL\]'){
        $clean = $s -replace '^\[DL\]\s*',''
        $clean = $clean -replace '\s*\(.*$',''
        return $clean.Trim()
    }
    ($s-replace':latest$','')
}
function Update-ModelTip {
    $m = Get-SelModel
    if(-not $m){$txtModelTipContent.Text="Select a model to see info";return}
    $mClean = $m.Trim()
    $rm = $null
    $rm = $RecommendedModels | Where-Object { $mClean -eq $_.Name } | Select-Object -First 1
    if(-not $rm){ $rm = $RecommendedModels | Where-Object { $mClean -eq "$($_.Name):latest" -or "$($mClean):latest" -eq $_.Name } | Select-Object -First 1 }
    if(-not $rm){
        $mBase = $mClean -replace ':[^:]+$',''
        $rm = $RecommendedModels | Where-Object { $mBase -eq ($_.Name -replace ':[^:]+$','') } | Select-Object -First 1
    }
    if(-not $rm){ $rm = $RecommendedModels | Where-Object { $mClean -like "$($_.Name)*" -or $_.Name -like "$mClean*" -or $mClean -like "*$($_.Name)*" } | Select-Object -First 1 }
    if($rm){
        $txtModelTipContent.Text = "$($rm.Name) ($($rm.Size))`nMin RAM: $($rm.MinRAM) GB | Quality: $($rm.Quality)`n`n$($rm.Info)"
    }else{
        $mLib = $mClean -replace ':[^:]+$',''
        $txtModelTipContent.Text = "$mClean`n(Custom or unlisted model)`n`nCheck: ollama.com/library/$mLib"
    }
}
$cmbModel.Add_SelectionChanged({ Update-ModelTip })
#endregion

#region Background Worker Infrastructure
$Script:MsgQueue = [System.Collections.Concurrent.ConcurrentQueue[string]]::new()
$Script:WorkerTimer = $null

function Start-BackgroundJob {
    param([ScriptBlock]$Job, [hashtable]$Params = @{})

    $Script:WorkerPS = [PowerShell]::Create()
    $runspace = [RunspaceFactory]::CreateRunspace()
    $runspace.ApartmentState = "MTA"
    $runspace.Open()
    $runspace.SessionStateProxy.SetVariable("MsgQueue", $Script:MsgQueue)
    $runspace.SessionStateProxy.SetVariable("OllamaUrl", $Global:OllamaUrl)
    $runspace.SessionStateProxy.SetVariable("CancelFlag", $Global:CancelDownload)
    foreach($k in $Params.Keys){ $runspace.SessionStateProxy.SetVariable($k, $Params[$k]) }
    $Script:WorkerPS.Runspace = $runspace
    $Script:WorkerPS.AddScript($Job) | Out-Null
    $Script:WorkerHandle = $Script:WorkerPS.BeginInvoke()

    $Script:WorkerTimer = [Windows.Threading.DispatcherTimer]::new()
    $Script:WorkerTimer.Interval = [TimeSpan]::FromMilliseconds(200)
    $Script:WorkerTimer.Add_Tick({
        $msg = $null
        while($Script:MsgQueue.TryDequeue([ref]$msg)){
            if($msg -like "AI:*")   { Ui-Append ($msg.Substring(3)) $null $txtAI }
            elseif($msg -like "PROGRESS:*"){
                # Update status bar with percentage
                $pct = $msg.Substring(9)
                Ui-SetStatus "Downloading... $pct%"
            }
            elseif($msg -like "EXT:*"){
                $Global:ExtJson = $msg.Substring(4)
                Ui-Log "Extended JSON received ($($Global:ExtJson.Length) chars)" "DEBUG"
            }
            elseif($msg -like "DONE:*"){
                $Script:WorkerTimer.Stop()
                try{$Script:WorkerPS.EndInvoke($Script:WorkerHandle)}catch{}
                try{$Script:WorkerPS.Runspace.Close()}catch{}
                try{$Script:WorkerPS.Dispose()}catch{}
                $payload = $msg.Substring(5)
                if($payload -like "SCAN:*"){
                    $json = $payload.Substring(5)
                    try{ $Global:DiagData = $json | ConvertFrom-Json -EA Stop
                         if($Global:ExtJson){
                             try{
                                 $Global:ExtDiagData = $Global:ExtJson | ConvertFrom-Json -EA Stop
                                 Ui-Log "Extended data parsed: $($Global:ExtDiagData.PSObject.Properties.Name -join ', ')" "DEBUG"
                             }catch{ Ui-Log "Failed to parse extended data: $($_.Exception.Message)" "ERROR" }
                         }
                         $D = $Global:DiagData
                         Ui-Log "=== SCAN COMPLETE ===" "OK"
                         Ui-Log "" "INFO"
                         $cc=@($D.Events|Where-Object{$_.Level-eq"Critical"}).Count
                         $ec=@($D.Events|Where-Object{$_.Level-eq"Error"}).Count
                         $wc=@($D.Events|Where-Object{$_.Level-eq"Warning"}).Count
                         $sc2=@($D.Services).Count; $pu=$D.Updates.PendingCount
                         Ui-Log "--- SUMMARY ---" "INFO"
                         if($D.SystemInfo.OS)      { Ui-Log "OS: $($D.SystemInfo.OS)" "INFO" }
                         if($D.SystemInfo.CPU)      { Ui-Log "CPU: $($D.SystemInfo.CPU)" "INFO" }
                         if($D.SystemInfo.RAMTotal) { Ui-Log "RAM: $($D.SystemInfo.RAMTotal) (Used: $($D.SystemInfo.RAMUsage))" "INFO" }
                         if($D.SystemInfo.Uptime)   { Ui-Log "Uptime: $($D.SystemInfo.Uptime)" "INFO" }
                         foreach($dk in @($D.Disks)){ if($dk.Drive){ $sl = if($dk.Status -eq "CRITICAL"){"ERROR"}elseif($dk.Status -eq "WARNING"){"WARN"}else{"OK"}; Ui-Log "Disk $($dk.Drive) $($dk.Label): $($dk.FreeGB)GB free / $($dk.TotalGB)GB ($($dk.Status))" $sl } }
                         if($cc -gt 0 -or $ec -gt 0){ Ui-Log "Events: $cc critical, $ec errors, $wc warnings" "WARN" }else{ Ui-Log "Events: $wc warnings, no critical/errors" "OK" }
                         if($sc2 -gt 0){ Ui-Log "Stopped auto-start services: $sc2" "WARN"; foreach($sv in @($D.Services)){if($sv.DisplayName){Ui-Log "  - $($sv.DisplayName) ($($sv.Name))" "WARN"}} }else{ Ui-Log "All auto-start services running" "OK" }
                         foreach($ad in @($D.Network.Adapters)){if($ad.Adapter){Ui-Log "Network: $($ad.Adapter) IP=$($ad.IP) GW=$($ad.Gateway)" "INFO"}}
                         $EE = $null; if($Global:ExtJson){try{$EE=$Global:ExtJson|ConvertFrom-Json -EA Stop}catch{}}
                         if($EE -and $EE.ExternalIP.IP){Ui-Log "External IP: $($EE.ExternalIP.IP)" "INFO"}
                         if($D.Security.Defender){Ui-Log "Defender: $($D.Security.Defender) | RealTime: $($D.Security.RealTime)" $(if($D.Security.RealTime -eq "OFF"){"WARN"}else{"OK"})}
                         if($D.Security.Firewall){Ui-Log "Firewall: $($D.Security.Firewall)" "INFO"}
                         if($pu -gt 0){Ui-Log "Pending updates: $pu" "WARN";foreach($u in @($D.Updates.Pending)){if($u.Title){Ui-Log "  - [$($u.Severity)] $($u.Title)" "WARN"}}}
                         elseif($pu -eq 0){Ui-Log "Windows Updates: up to date" "OK"}
                         else{Ui-Log "Windows Updates: could not check" "WARN"}
                         if($EE -and $EE.Hosts.TotalCount){Ui-Log "Hosts file: $($EE.Hosts.TotalCount) active entries" "INFO"}
                         if($EE -and $EE.RemoteTools.TeamViewer){Ui-Log "TeamViewer: $($EE.RemoteTools.TeamViewer)" "INFO"}
                         if($EE -and $EE.RemoteTools.AnyDesk){Ui-Log "AnyDesk: $($EE.RemoteTools.AnyDesk)" "INFO"}
                         if($EE -and ($EE.BatteryReport.Available -eq $true -or $EE.BatteryReport.Available -eq "True")){
                             $bri = "Battery: Design=$($EE.BatteryReport.DesignCapacity)"
                             if($EE.BatteryReport.FullChargeCapacity){$bri += " FullCharge=$($EE.BatteryReport.FullChargeCapacity)"}
                             if($EE.BatteryReport.CycleCount){$bri += " Cycles=$($EE.BatteryReport.CycleCount)"}
                             Ui-Log $bri "INFO"
                         }
                         # Drivers summary
                         if($EE -and $EE.Drivers){
                             $drvArr = @($EE.Drivers)
                             $probDrivers = @($drvArr | Where-Object { $_.Problem -and $_.Problem -ne "None" -and $_.Problem -ne "0" })
                             if($probDrivers.Count -gt 0){ Ui-Log "Problem drivers: $($probDrivers.Count)" "WARN"; foreach($pd in $probDrivers){Ui-Log "  - $($pd.Name) [$($pd.Problem)]" "WARN"} }
                             else{ Ui-Log "Drivers: $($drvArr.Count) checked, no problems detected" "OK" }
                         }
                         # SMART summary
                         if($EE -and $EE.SmartDetails){
                             $smartArr = @($EE.SmartDetails)
                             foreach($sd in $smartArr){
                                 $sh = if($sd.Health){"$($sd.Health)"}else{"Unknown"}
                                 Ui-Log "SMART $($sd.Model): Health=$sh" $(if($sh -ne "Healthy"){"WARN"}else{"OK"})
                             }
                         }
                         if($D.Software){Ui-Log "Installed software: $(@($D.Software).Count)" "INFO"}
                         # BSOD summary
                         if($EE -and $EE.BSOD){
                             $bsodArr = @($EE.BSOD)
                             $bsodReal = @($bsodArr | Where-Object { $_.Type -eq "BSOD" -or $_.Type -eq "MiniDump" -or $_.Type -eq "FullDump" })
                             if($bsodReal.Count -gt 0){ Ui-Log "BSOD crashes found: $($bsodReal.Count)" "WARN"; foreach($b in ($bsodReal|Select-Object -First 5)){Ui-Log "  - [$($b.Time)] $($b.Type): $($b.Info.Substring(0,[Math]::Min(80,$b.Info.Length)))" "WARN"} }
                             else{ Ui-Log "No BSOD crashes detected" "OK" }
                         }
                         # RAM test summary
                         if($EE -and $EE.RAMTest){
                             if($EE.RAMTest.HasResults -eq $true){ Ui-Log "RAM diagnostics: results available" "INFO" }
                             if($EE.RAMTest.WHEAErrors){ Ui-Log "WHEA memory errors: $($EE.RAMTest.WHEAErrors)" "WARN" }
                         }
                         Ui-Log "--- END SUMMARY ---" "INFO"
                         Ui-Log "" "INFO"
                         Ui-Log "Ready for AI Analysis or Save Report." "INFO"
                         # Auto-save report if enabled
                         if($Global:AutoSaveReport){
                             $btnReport.RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Primitives.ButtonBase]::ClickEvent))
                         }
                    }catch{ Ui-Log "Failed to parse scan results: $($_.Exception.Message)" "ERROR" }
                }
                elseif($payload -like "AI:*"){
                    $Global:AiAnalysis = $payload.Substring(3)
                }
                elseif($payload -like "PULL:*"){
                    if($payload -eq "PULL:CANCELLED"){
                        Ui-Log "Download cancelled" "WARN"
                    }
                    # Reset download button state
                    $btnPull.Content = [char]0x2B07 + " Download"
                    $btnPull.Tag = $null
                    Refresh-Ollama
                }
                Ui-SetStatus "Done."
                Set-UiBusy $false
            }
            elseif($msg -like "ERROR:*"){ Ui-Log $msg.Substring(6) "ERROR" }
            else{ Ui-Append $msg }
        }
    })
    $Script:WorkerTimer.Start()
}
#endregion

#region Scan Worker ScriptBlock
$ScanWorkerScript = {
    function Q([string]$m){ $MsgQueue.Enqueue($m) }
    function QLog([string]$m,[string]$l="INFO"){ Q("[$(Get-Date -F 'yyyy-MM-dd HH:mm:ss')][$l] $m") }
    function QDbg([string]$m){ if($VerboseMode){ Q("[$(Get-Date -F 'yyyy-MM-dd HH:mm:ss')][DEBUG] $m") } }

    $diag = @{SystemInfo=@{};Hardware=@{};Battery=@{};Disks=@();Events=@();Services=@();AllServices=@();Network=@{Adapters=@();DNSTests=@()};Security=@{};Updates=@{PendingCount=0;Pending=@()};Processes=@{TopCPU=@();TopRAM=@()};Startup=@();ScheduledTasks=@();Autoruns=@();Integrity=@{};SearchIndex=@{};Software=@();Hosts=@{Entries=@();Raw=""};UserInfo=@{LocalAccounts=@();ActiveSessions=@();UserFolders=@()};ExternalIP=@{};Hotfixes=@();RemoteTools=@{};ChkdskLogs=@();BatteryReport=@{};NetConnections=@();NetRoutes=@();ArpTable=@();MappedDrives=@();Drivers=@();SmartDetails=@();BSOD=@();RAMTest=@{};Perfmon=@{}}

    try {
    # ── System Info ──
    if($DoSysInfo){
        QLog "System information..." "SCAN"
        $os=Get-CimInstance Win32_OperatingSystem; $cs=Get-CimInstance Win32_ComputerSystem
        $cpu=Get-CimInstance Win32_Processor|Select-Object -First 1; $bios=Get-CimInstance Win32_BIOS
        $up=(Get-Date)-$os.LastBootUpTime
        $diag.SystemInfo = @{
            ComputerName=$env:COMPUTERNAME; User="$env:USERDOMAIN\$env:USERNAME"
            OS="$($os.Caption) $($os.Version) Build $($os.BuildNumber)"; Arch=$os.OSArchitecture
            CPU="$($cpu.Name) ($($cpu.NumberOfCores)c/$($cpu.NumberOfLogicalProcessors)t)"
            RAMTotal="{0:N1} GB"-f($cs.TotalPhysicalMemory/1GB); RAMFree="{0:N1} GB"-f($os.FreePhysicalMemory/1MB)
            RAMUsage="{0:N0}%"-f((1-($os.FreePhysicalMemory*1KB/$cs.TotalPhysicalMemory))*100)
            Uptime="{0}d {1}h {2}m"-f $up.Days,$up.Hours,$up.Minutes
            LastBoot=$os.LastBootUpTime.ToString("yyyy-MM-dd HH:mm:ss"); Domain=$cs.Domain
            BIOS="$($bios.Manufacturer) $($bios.SMBIOSBIOSVersion)"; Manufacturer=$cs.Manufacturer; Model=$cs.Model
            SerialNumber=$bios.SerialNumber; InstallDate=$os.InstallDate.ToString("yyyy-MM-dd")
        }
        QDbg "OS: $($diag.SystemInfo.OS)"
        QDbg "CPU: $($diag.SystemInfo.CPU)"
        QDbg "RAM: $($diag.SystemInfo.RAMTotal) | Free: $($diag.SystemInfo.RAMFree) | Usage: $($diag.SystemInfo.RAMUsage)"
        QDbg "Uptime: $($diag.SystemInfo.Uptime) | Last Boot: $($diag.SystemInfo.LastBoot)"
        QDbg "Manufacturer: $($diag.SystemInfo.Manufacturer) | Model: $($diag.SystemInfo.Model)"
        QDbg "Serial: $($diag.SystemInfo.SerialNumber) | BIOS: $($diag.SystemInfo.BIOS)"
        QLog "Done" "OK"
    }

    # ── Hardware ──
    if($DoHardware){
        QLog "Hardware details..." "SCAN"
        $ram = Get-CimInstance Win32_PhysicalMemory | ForEach-Object { "$($_.BankLabel) $([math]::Round($_.Capacity/1MB))MB $($_.PartNumber)" }
        $gpu = Get-CimInstance Win32_VideoController | ForEach-Object { "$($_.Name) $([math]::Round($_.AdapterRAM/1MB))MB" }
        $diag.Hardware = @{ RAM=($ram-join " | "); GPU=($gpu-join " | ") }
        QDbg "RAM: $($diag.Hardware.RAM)"
        QLog "Done" "OK"
    }

    # ── Battery ──
    if($DoBattery){
        QLog "Battery info..." "SCAN"
        try {
            $bat = Get-CimInstance Win32_Battery -EA Stop
            if($bat){
                $diag.Battery = @{
                    Status=$bat.Status; EstCharge="$($bat.EstimatedChargeRemaining)%"
                    RunTime=if($bat.EstimatedRunTime -and $bat.EstimatedRunTime -lt 71582788){"$($bat.EstimatedRunTime) min"}else{"On AC"}
                    Chemistry=switch($bat.Chemistry){1{"Other"};2{"Unknown"};3{"Lead Acid"};4{"Nickel Cadmium"};5{"Nickel Metal Hydride"};6{"Lithium-ion"};default{"N/A"}}
                    DesignCapacity=if($bat.DesignCapacity){"$($bat.DesignCapacity) mWh"}else{"N/A"}
                    FullCharge=if($bat.FullChargeCapacity){"$($bat.FullChargeCapacity) mWh"}else{"N/A"}
                }
                if($bat.DesignCapacity -and $bat.FullChargeCapacity -and $bat.DesignCapacity -gt 0){
                    $health = [math]::Round(($bat.FullChargeCapacity / $bat.DesignCapacity) * 100, 1)
                    $diag.Battery.Health = "$health%"
                }
                QDbg "Battery: $($diag.Battery.EstCharge) charge, health=$($diag.Battery.Health)"
            } else { $diag.Battery = @{Status="No battery detected"} }
        } catch { $diag.Battery = @{Status="N/A"} }
        QLog "Done" "OK"
    }

    # ── Disks ──
    if($DoDisks){
        QLog "Disk health..." "SCAN"
        $diag.Disks = @()
        Get-CimInstance Win32_LogicalDisk -Filter "DriveType=3" | ForEach-Object {
            if($_.Size -and $_.Size -gt 0){
                $fp=[math]::Round(($_.FreeSpace/$_.Size)*100,1)
                $diag.Disks += @{Drive=$_.DeviceID;Label=$_.VolumeName;TotalGB="{0:N1}"-f($_.Size/1GB);FreeGB="{0:N1}"-f($_.FreeSpace/1GB);FreePercent="$fp%";Status=if($fp-lt 10){"CRITICAL"}elseif($fp-lt 20){"WARNING"}else{"OK"}}
                QDbg "Drive $($_.DeviceID) $($_.VolumeName): $("{0:N1}"-f($_.FreeSpace/1GB))GB free / $("{0:N1}"-f($_.Size/1GB))GB ($fp%)"
            }
        }
        try{Get-CimInstance -Namespace root\Microsoft\Windows\Storage -ClassName MSFT_PhysicalDisk -EA Stop|ForEach-Object{
            $hm=@{0="Healthy";1="Warning";2="Unhealthy"};$h=$hm[[int]$_.HealthStatus];if(-not $h){$h="Unknown"}
            $diag.Disks+=@{Drive="Disk $($_.DeviceId)";Label=$_.FriendlyName;TotalGB="{0:N1}"-f($_.Size/1GB);FreeGB="-";FreePercent="-";Status=$h}
            QDbg "Physical disk $($_.FriendlyName): $h"
        }}catch{}
        QLog "Done" "OK"
    }

    # ── SMART Details ──
    if($DoSmartDetails){
        QLog "SMART details..." "SCAN"
        $diag.SmartDetails = @()
        try{
            # Try WMI SMART data
            $smartData = Get-CimInstance -Namespace root\WMI -ClassName MSStorageDriver_FailurePredictStatus -EA SilentlyContinue
            $smartThresh = Get-CimInstance -Namespace root\WMI -ClassName MSStorageDriver_FailurePredictData -EA SilentlyContinue
            $physDisks = Get-CimInstance -Namespace root\Microsoft\Windows\Storage -ClassName MSFT_PhysicalDisk -EA SilentlyContinue
            if($physDisks){
                foreach($pd in $physDisks){
                    $smartEntry = @{
                        Model=$pd.FriendlyName; MediaType=switch($pd.MediaType){3{"HDD"};4{"SSD"};5{"SCM"};default{"Unknown"}}
                        BusType=switch($pd.BusType){3{"ATA"};11{"SATA"};17{"NVMe"};default{"Other ($($pd.BusType))"}}
                        Health=@{0="Healthy";1="Warning";2="Unhealthy";5="Unknown"}[[int]$pd.HealthStatus]
                        Size="{0:N1} GB"-f($pd.Size/1GB)
                        FirmwareVersion=$pd.FirmwareVersion
                        SerialNumber=$pd.SerialNumber
                    }
                    # Try to get reliability counters
                    try{
                        $rel = Get-CimInstance -Namespace root\Microsoft\Windows\Storage -ClassName MSFT_StorageReliabilityCounter -EA Stop |
                            Where-Object { $_.DeviceId -eq $pd.DeviceId } | Select-Object -First 1
                        if($rel){
                            $smartEntry.Temperature = if($rel.Temperature){"$($rel.Temperature)°C"}else{"N/A"}
                            $smartEntry.ReadErrors = if($rel.ReadErrorsTotal){$rel.ReadErrorsTotal}else{0}
                            $smartEntry.WriteErrors = if($rel.WriteErrorsTotal){$rel.WriteErrorsTotal}else{0}
                            $smartEntry.PowerOnHours = if($rel.PowerOnHours){$rel.PowerOnHours}else{"N/A"}
                            $smartEntry.Wear = if($rel.Wear){"$($rel.Wear)%"}else{"N/A"}
                        }
                    }catch{}
                    $diag.SmartDetails += $smartEntry
                    QDbg "SMART: $($pd.FriendlyName) Health=$($smartEntry.Health) Temp=$($smartEntry.Temperature)"
                }
            }
            # Fallback: FailurePredictStatus
            if($smartData){
                foreach($sd in $smartData){
                    $predicted = $sd.PredictFailure
                    if($predicted){
                        QLog "SMART FAILURE PREDICTED on a disk!" "WARN"
                    }
                }
            }
        }catch{ QDbg "SMART query error: $($_.Exception.Message)" }
        QLog "$(@($diag.SmartDetails).Count) disks with SMART data" "OK"
    }

    # ── Events ──
    if($DoEvents){
        $evtHrs = if($EventHours){$EventHours}else{24}
        QLog "Event logs (${evtHrs}h)..." "SCAN"
        $after=(Get-Date).AddHours(-$evtHrs); $diag.Events=@()
        foreach($ln in @("System","Application")){
            try{
                $evts=Get-WinEvent -FilterHashtable @{LogName=$ln;Level=@(1,2,3);StartTime=$after} -MaxEvents 50 -EA Stop
                foreach($e in $evts){
                    $lv=switch($e.Level){1{"Critical"};2{"Error"};3{"Warning"};default{"Info"}}
                    $msg=if($e.Message){($e.Message-replace'\r?\n',' ')}else{"(no message)"}; if($msg.Length-gt 300){$msg=$msg.Substring(0,300)}
                    $diag.Events+=@{Time=$e.TimeCreated.ToString("yyyy-MM-dd HH:mm:ss");Log=$ln;Level=$lv;Source=$e.ProviderName;EventID=$e.Id;Message=$msg}
                }
                QDbg "$ln log: $($evts.Count) events"
            }catch{ if($_.Exception.Message -notlike "*No events*"){QDbg "Could not read $ln`: $($_.Exception.Message)"} }
        }
        QLog "$(@($diag.Events).Count) issues" "OK"
    }

    # ── Services ──
    if($DoServices){
        QLog "Services..." "SCAN"
        $diag.Services=@()
        $diag.AllServices=@()
        Get-CimInstance Win32_Service | ForEach-Object{
            $svc = @{Name=$_.Name;DisplayName=$_.DisplayName;State=$_.State;StartMode=$_.StartMode;RunAs=$_.StartName}
            $diag.AllServices += $svc
            if($_.StartMode -eq "Auto" -and $_.State -ne "Running" -and $_.Name -notmatch "sppsvc|TrustedInstaller|SysMain|MapsBroker|wuauserv|SecurityHealthService|edgeupdate|gupdate|GoogleUpdate|OneSyncSvc|tiledatamodelsvc"){
                $diag.Services += $svc
                QDbg "Stopped: $($_.DisplayName)"
            }
        }
        QLog "$(@($diag.Services).Count) stopped auto-start, $(@($diag.AllServices).Count) total" "OK"
    }

    # ── Network ──
    if($DoNetwork){
        QLog "Network..." "SCAN"
        $diag.Network=@{Adapters=@();DNSTests=@()}
        Get-NetAdapter -EA SilentlyContinue|Where-Object Status -eq "Up"|ForEach-Object{
            $ip=(Get-NetIPAddress -InterfaceIndex $_.ifIndex -EA SilentlyContinue|Where-Object{$_.AddressFamily-eq"IPv4"-and$_.IPAddress-ne"127.0.0.1"}).IPAddress -join ", "
            $gw=(Get-NetRoute -InterfaceIndex $_.ifIndex -DestinationPrefix "0.0.0.0/0" -EA SilentlyContinue).NextHop
            $dns=(Get-DnsClientServerAddress -InterfaceIndex $_.ifIndex -AddressFamily IPv4 -EA SilentlyContinue).ServerAddresses -join ", "
            $spd="$($_.LinkSpeed)"
            if($spd-match'^\d+$'){$spd="{0:N0} Mbps"-f([double]$spd/1e6)}
            $diag.Network.Adapters+=@{Adapter=$_.Name;IP=$ip;Gateway=$gw;DNS=$dns;Speed=$spd;MAC=$_.MacAddress}
            QDbg "$($_.Name): IP=$ip Speed=$spd"
        }
        foreach($h in @("google.com","microsoft.com")){
            try{$r=Resolve-DnsName $h -Type A -DnsOnly -EA Stop|Select-Object -First 1;$diag.Network.DNSTests+="$h = $($r.IPAddress) [OK]"}
            catch{$diag.Network.DNSTests+="$h = FAILED"}
        }
        QLog "$(@($diag.Network.Adapters).Count) adapters" "OK"
    }

    # ── Security ──
    if($DoSecurity){
        QLog "Security..." "SCAN"
        $diag.Security=@{}
        try{$fw=Get-NetFirewallProfile -EA Stop;$diag.Security.Firewall=($fw|ForEach-Object{"$($_.Name):$(if($_.Enabled){'ON'}else{'OFF'})"})-join " | "}catch{$diag.Security.Firewall="N/A"}
        try{$d=Get-MpComputerStatus -EA Stop;$diag.Security.Defender=if($d.AntivirusEnabled){"Enabled"}else{"DISABLED"};$diag.Security.RealTime=if($d.RealTimeProtectionEnabled){"ON"}else{"OFF"};$diag.Security.Definitions=$d.AntivirusSignatureLastUpdated.ToString("yyyy-MM-dd HH:mm");$diag.Security.EngineVersion=$d.AMEngineVersion}catch{$diag.Security.Defender="N/A"}
        try{$u=(Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" -EA Stop).EnableLUA;$diag.Security.UAC=if($u-eq 1){"Enabled"}else{"DISABLED"}}catch{$diag.Security.UAC="N/A"}
        QDbg "Firewall: $($diag.Security.Firewall)"
        QDbg "Defender: $($diag.Security.Defender) | RealTime: $($diag.Security.RealTime) | Defs: $($diag.Security.Definitions)"
        QDbg "UAC: $($diag.Security.UAC)"
        QLog "Done" "OK"
    }

    # ── Updates ──
    if($DoUpdates){
        QLog "Windows Updates..." "SCAN"
        try{$sess=New-Object -ComObject Microsoft.Update.Session -EA Stop;$sr=$sess.CreateUpdateSearcher();$pd=$sr.Search("IsInstalled=0 AND IsHidden=0")
            $diag.Updates=@{PendingCount=$pd.Updates.Count;Pending=@()}
            foreach($u in $pd.Updates){
                $diag.Updates.Pending+=@{Title=$u.Title;KB=($u.KBArticleIDs-join",");Severity=if($u.MsrcSeverity){$u.MsrcSeverity}else{"N/A"}}
                QDbg "  Pending: [$($u.MsrcSeverity)] $($u.Title)"
            }
            QLog "$($diag.Updates.PendingCount) pending" "OK"
        }catch{$diag.Updates=@{PendingCount=-1;Pending=@()};QLog "Could not check updates" "WARN"}
    }

    # ── Processes ──
    if($DoProcesses){
        QLog "Processes..." "SCAN"
        $diag.Processes=@{TopCPU=@();TopRAM=@()}
        Get-Process|Sort-Object CPU -Descending|Select-Object -First 15|ForEach-Object{
            $diag.Processes.TopCPU+=@{Name=$_.ProcessName;PID=$_.Id;CPU="{0:N1}"-f $_.CPU;RAM="{0:N0}"-f($_.WorkingSet64/1MB)}
        }
        Get-Process|Sort-Object WorkingSet64 -Descending|Select-Object -First 15|ForEach-Object{
            $diag.Processes.TopRAM+=@{Name=$_.ProcessName;PID=$_.Id;RAM="{0:N0}"-f($_.WorkingSet64/1MB);CPU="{0:N1}"-f $_.CPU}
        }
        QLog "Done" "OK"
    }

    # ── Startup ──
    if($DoStartup){
        QLog "Startup programs..." "SCAN"
        $diag.Startup=@()
        foreach($p in @("HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run","HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run")){
            try{$items=Get-ItemProperty $p -EA Stop;$items.PSObject.Properties|Where-Object{$_.Name-notmatch'^PS'}|ForEach-Object{
                $cmd=$_.Value-replace'"','';if($cmd.Length-gt 150){$cmd=$cmd.Substring(0,150)}
                $diag.Startup+=@{Name=$_.Name;Command=$cmd;Location=$p-replace'HKLM:\\','HKLM\'-replace'HKCU:\\','HKCU\'}
            }}catch{}
        }
        $sf="$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Startup"
        if(Test-Path $sf){Get-ChildItem $sf -EA SilentlyContinue|ForEach-Object{$diag.Startup+=@{Name=$_.Name;Command=$_.FullName;Location="Startup Folder"}}}
        $sf2="$env:ProgramData\Microsoft\Windows\Start Menu\Programs\Startup"
        if(Test-Path $sf2){Get-ChildItem $sf2 -EA SilentlyContinue|ForEach-Object{$diag.Startup+=@{Name=$_.Name;Command=$_.FullName;Location="Common Startup Folder"}}}
        QLog "$(@($diag.Startup).Count) items" "OK"
    }

    # ── Scheduled Tasks ──
    if($DoScheduled){
        QLog "Scheduled Tasks..." "SCAN"
        $diag.ScheduledTasks=@()
        Get-ScheduledTask -EA SilentlyContinue|Where-Object{$_.State-ne"Disabled"-and $_.TaskPath-notmatch'^\\Microsoft\\'}|Select-Object -First 40|ForEach-Object{
            $info=Get-ScheduledTaskInfo $_ -EA SilentlyContinue
            $diag.ScheduledTasks+=@{Name=$_.TaskName;Path=$_.TaskPath;State=$_.State.ToString();LastRun=if($info.LastRunTime){"$($info.LastRunTime)"}else{"Never"};Author=$_.Author}
        }
        QLog "$(@($diag.ScheduledTasks).Count) active non-Microsoft tasks" "OK"
    }

    # ── Autoruns ──
    if($DoAutoruns){
        QLog "Autoruns (registry/shell/drivers)..." "SCAN"
        $diag.Autoruns=@()
        $paths=@(
            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\ShellIconOverlayIdentifiers",
            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved",
            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Browser Helper Objects"
        )
        foreach($rp in $paths){
            try{if(Test-Path $rp){
                Get-ChildItem $rp -EA SilentlyContinue|Select-Object -First 20|ForEach-Object{
                    $diag.Autoruns+=@{Type="Shell";Name=$_.PSChildName;Location=$rp-replace'HKLM:\\SOFTWARE\\',''}
                }
            }}catch{}
        }
        try{Get-CimInstance Win32_SystemDriver|Where-Object{$_.PathName-and $_.PathName-notmatch'\\Windows\\System32\\drivers\\'}|Select-Object -First 20|ForEach-Object{
            $diag.Autoruns+=@{Type="Driver";Name=$_.DisplayName;Location=$_.PathName}
        }}catch{}
        QLog "$(@($diag.Autoruns).Count) items" "OK"
    }

    # ── Integrity ──
    if($DoIntegrity){
        QLog "System integrity..." "SCAN"
        $diag.Integrity=@{}
        # SFC results - multiple detection methods
        try{
            $sfcFound = $false
            # Method 1: CBS.log (primary source)
            $cbsLog="$env:SystemRoot\Logs\CBS\CBS.log"
            if(Test-Path $cbsLog){
                $last=Get-Content $cbsLog -Tail 1000 -EA SilentlyContinue
                # SFC writes "Windows Resource Protection" lines
                $sfcResults = $last | Where-Object { $_ -match 'Windows Resource Protection (did not find|found corrupt|could not perform|found integrity)' } | Select-Object -Last 1
                if(-not $sfcResults){
                    # Broader patterns
                    $sfcResults = $last | Where-Object { $_ -match 'Verification\s+(100%\s+)?complete|integrity violations|successfully repaired|Cannot repair member file|Hashes for member' } | Select-Object -Last 1
                }
                if($sfcResults){
                    $sfcClean = $sfcResults.Trim() -replace '^\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2},\s*','' -replace '^\[SR\]\s*',''
                    $diag.Integrity.LastSFC = $sfcClean
                    $sfcFound = $true
                    QDbg "SFC from CBS.log: $sfcClean"
                }
            }
            # Method 2: SFC log in CBS folder
            if(-not $sfcFound){
                $sfcLog = "$env:SystemRoot\Logs\CBS\SFC.log"
                if(Test-Path $sfcLog){
                    $sfcLast = Get-Content $sfcLog -Tail 50 -EA SilentlyContinue
                    $sfcLine = $sfcLast | Where-Object { $_ -match 'Windows Resource Protection|Verification|integrity' } | Select-Object -Last 1
                    if($sfcLine){ $diag.Integrity.LastSFC = $sfcLine.Trim(); $sfcFound = $true }
                }
            }
            # Method 3: Event Log (Application, source=SFC)
            if(-not $sfcFound){
                try{
                    $sfcEvent = Get-WinEvent -FilterHashtable @{LogName='Application';ProviderName='Microsoft-Windows-WindowsUpdateClient','SFC'} -MaxEvents 5 -EA SilentlyContinue |
                        Where-Object { $_.Message -match 'Resource Protection|sfc|integrity' } | Select-Object -First 1
                    if($sfcEvent){ $diag.Integrity.LastSFC = "[$($sfcEvent.TimeCreated.ToString('yyyy-MM-dd'))] $($sfcEvent.Message.Substring(0,[Math]::Min(200,$sfcEvent.Message.Length)))"; $sfcFound = $true }
                }catch{}
            }
            if(-not $sfcFound){ $diag.Integrity.LastSFC = "No SFC results found (run: sfc /scannow)" }
        }catch{$diag.Integrity.LastSFC="Error reading SFC results: $($_.Exception.Message)"}
        # DISM health status
        try{
            $dismLog = "$env:SystemRoot\Logs\DISM\dism.log"
            if(Test-Path $dismLog){
                $dismLast = Get-Content $dismLog -Tail 200 -EA SilentlyContinue
                $dismResult = $dismLast | Where-Object { $_ -match 'The component store is|repairable|no component store corruption|RestoreHealth' } | Select-Object -Last 1
                if($dismResult){ $diag.Integrity.LastDISM = $dismResult.Trim() -replace '^\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2},\s*','' }
                else{ $diag.Integrity.LastDISM = "No recent DISM results" }
            }else{ $diag.Integrity.LastDISM = "DISM log not found" }
        }catch{ $diag.Integrity.LastDISM = "Error reading DISM log" }
        # Pending reboot
        try{$cbsReg=Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing" -EA SilentlyContinue
            $diag.Integrity.RebootPending=if($cbsReg.RebootPending){"YES"}else{"No"}
        }catch{$diag.Integrity.RebootPending="N/A"}
        # Component store health
        try{
            $compStore = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing" -EA SilentlyContinue
            $diag.Integrity.LastCleanup = if($compStore.LastScavengeDateTime){"$($compStore.LastScavengeDateTime)"}else{"N/A"}
        }catch{}
        QDbg "SFC: $($diag.Integrity.LastSFC)"
        QDbg "DISM: $($diag.Integrity.LastDISM)"
        QLog "Done" "OK"
    }

    # ── Search/Indexing ──
    if($DoIndexing){
        QLog "Windows Search/Indexing..." "SCAN"
        $diag.SearchIndex=@{}
        try{$svc=Get-Service "WSearch" -EA SilentlyContinue
            $diag.SearchIndex.Service=if($svc){$svc.Status.ToString()}else{"Not installed"}
        }catch{$diag.SearchIndex.Service="N/A"}
        try{$sm=New-Object -ComObject Microsoft.Search.Interop.CSearchManager -EA Stop;$cat=$sm.GetCatalog("SystemIndex")
            $stMap=@{0="Idle";1="Paused";2="Recovering";3="Full Crawl";4="Incremental Crawl";6="Not Running"}
            $diag.SearchIndex.Status=$stMap[[int]$cat.GetCatalogStatus()]; $diag.SearchIndex.Items="{0:N0}"-f $cat.NumberOfItems()
        }catch{$diag.SearchIndex.Status="Could not query"}
        $idxP="$env:ProgramData\Microsoft\Search\Data\Applications\Windows"
        if(Test-Path $idxP){$diag.SearchIndex.Size="{0:N1} MB"-f((Get-ChildItem $idxP -Recurse -EA SilentlyContinue|Measure-Object Length -Sum).Sum/1MB)}
        QLog "Done" "OK"
    }

    # ── Installed Software ──
    if($DoSoftware){
        QLog "Installed software..." "SCAN"
        $diag.Software=@()
        $regPaths=@("HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*","HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*")
        foreach($rp in $regPaths){
            Get-ItemProperty $rp -EA SilentlyContinue|Where-Object{$_.DisplayName-and $_.DisplayName.Trim()}|ForEach-Object{
                $diag.Software+=@{Name=$_.DisplayName;Version=$_.DisplayVersion;Publisher=$_.Publisher}
            }
        }
        $diag.Software=$diag.Software|Group-Object{$_.Name}|ForEach-Object{$_.Group[0]}|Sort-Object{$_.Name}|Select-Object -First 100
        QLog "$(@($diag.Software).Count) applications" "OK"
    }

    # ── Hosts File ──
    if($DoHosts){
        QLog "Hosts file..." "SCAN"
        $hostsPath = "$env:SystemRoot\System32\drivers\etc\hosts"
        try{
            if(Test-Path $hostsPath){
                $raw = Get-Content $hostsPath -EA Stop
                $allEntries = @()
                foreach($line in $raw){
                    $l = $line.Trim()
                    if($l -and -not $l.StartsWith('#')){
                        $parts = $l -split '\s+',2
                        if($parts.Count -ge 2){ $allEntries += @{IP=$parts[0];Hostname=$parts[1]} }
                    }
                }
                $totalHostEntries = $allEntries.Count
                $diag.Hosts.TotalCount = $totalHostEntries
                if($totalHostEntries -gt 200){
                    $diag.Hosts.Entries = @($allEntries | Select-Object -First 200)
                    $diag.Hosts.Truncated = $true
                }else{
                    $diag.Hosts.Entries = $allEntries
                    $diag.Hosts.Truncated = $false
                }
                if($raw.Count -le 200){ $diag.Hosts.Raw = ($raw -join "`n") }
                else{ $diag.Hosts.Raw = "(File too large: $($raw.Count) lines)" }
                QLog "$totalHostEntries active entries" "OK"
            }else{QLog "Hosts file not found" "WARN"}
        }catch{QLog "Error reading hosts: $($_.Exception.Message)" "WARN"}
    }

    # ── User Information ──
    if($DoUserInfo){
        QLog "User information..." "SCAN"
        $diag.UserInfo = @{LocalAccounts=@();ActiveSessions=@();UserFolders=@()}
        try{
            Get-LocalUser -EA Stop | ForEach-Object {
                $diag.UserInfo.LocalAccounts += @{Name=$_.Name;Enabled=$_.Enabled;PasswordRequired=$_.PasswordRequired;LastLogon=if($_.LastLogon){"$($_.LastLogon)"}else{"Never"}}
            }
        }catch{}
        try{
            $q = & quser.exe 2>&1
            if($LASTEXITCODE -eq 0 -and $q){ foreach($line in $q){if($line -match '\w'){$diag.UserInfo.ActiveSessions += $line.Trim()}} }
        }catch{}
        try{ Get-ChildItem "C:\Users" -Directory -Force -EA SilentlyContinue | ForEach-Object {$diag.UserInfo.UserFolders += $_.Name} }catch{}
        QLog "$(@($diag.UserInfo.LocalAccounts).Count) accounts" "OK"
    }

    # ── External IP ──
    if($DoExternalIP){
        QLog "External IP..." "SCAN"
        $diag.ExternalIP = @{}
        try{
            $ip = Invoke-RestMethod -Uri "https://ifconfig.me/ip" -TimeoutSec 8 -EA Stop
            $diag.ExternalIP.IP = $ip.Trim()
            QLog "External IP: $($diag.ExternalIP.IP)" "OK"
        }catch{
            try{$ip = Invoke-RestMethod -Uri "https://api.ipify.org" -TimeoutSec 8 -EA Stop; $diag.ExternalIP.IP=$ip.Trim(); QLog "External IP: $ip" "OK"}
            catch{$diag.ExternalIP.IP="Could not determine";QLog "External IP: failed" "WARN"}
        }
    }

    # ── Installed Hotfixes ──
    if($DoHotfixes){
        QLog "Installed hotfixes..." "SCAN"
        $diag.Hotfixes = @()
        try{
            Get-CimInstance Win32_QuickFixEngineering -EA Stop | Sort-Object InstalledOn -Descending | Select-Object -First 30 | ForEach-Object {
                $diag.Hotfixes += @{HotFixID=$_.HotFixID;InstalledOn=if($_.InstalledOn){"$($_.InstalledOn.ToString('yyyy-MM-dd'))"}else{"Unknown"};Description=$_.Description}
            }
            QLog "$(@($diag.Hotfixes).Count) hotfixes" "OK"
        }catch{QLog "Could not get hotfixes" "WARN"}
    }

    # ── Remote Tools ──
    if($DoRemoteTools){
        QLog "Remote tools..." "SCAN"
        $diag.RemoteTools = @{}
        try{
            $tvId = (Get-ItemProperty "HKLM:\SOFTWARE\WOW6432Node\TeamViewer" -Name ClientID -EA Stop).ClientID
            $diag.RemoteTools.TeamViewer = "Installed (ID: $tvId)"
        }catch{
            $tvSvc = Get-Service "TeamViewer" -EA SilentlyContinue
            if($tvSvc){$diag.RemoteTools.TeamViewer = "Installed (Service: $($tvSvc.Status))"}
            else{$diag.RemoteTools.TeamViewer = "Not detected"}
        }
        $adSvc = Get-Service "AnyDesk*" -EA SilentlyContinue
        if($adSvc){$diag.RemoteTools.AnyDesk = "Installed (Service: $($adSvc.Status))"}
        else{
            $adPath = Get-Command "anydesk.exe" -EA SilentlyContinue
            if($adPath){$diag.RemoteTools.AnyDesk = "Installed"}else{$diag.RemoteTools.AnyDesk = "Not detected"}
        }
        $rdSvc = Get-Service "RustDesk" -EA SilentlyContinue
        if($rdSvc){$diag.RemoteTools.RustDesk = "Installed (Service: $($rdSvc.Status))"}
        else{$diag.RemoteTools.RustDesk = "Not detected"}
        QLog "Done" "OK"
    }

    # ── Chkdsk Logs ──
    if($DoChkdskLogs){
        QLog "Chkdsk logs..." "SCAN"
        $diag.ChkdskLogs = @()
        try{
            $chkEntries = @()
            try{$chkEntries += Get-WinEvent -LogName "Microsoft-Windows-Chkdsk/Operational" -MaxEvents 10 -EA Stop}catch{}
            try{$chkEntries += Get-WinEvent -FilterHashtable @{LogName='Application';Id=1001} -MaxEvents 10 -EA Stop | Where-Object{$_.ProviderName -match 'wininit|Wininit'}}catch{}
            if($chkEntries.Count -gt 0){
                $chkEntries | Sort-Object TimeCreated -Descending | Select-Object -First 10 | ForEach-Object {
                    $msg = if($_.Message){($_.Message -replace '\r?\n',' ')}else{"(no message)"}
                    if($msg.Length -gt 300){$msg=$msg.Substring(0,300)}
                    $diag.ChkdskLogs += @{Time=$_.TimeCreated.ToString("yyyy-MM-dd HH:mm");Message=$msg}
                }
            }
            QLog "$(@($diag.ChkdskLogs).Count) entries" "OK"
        }catch{QLog "Could not read chkdsk logs" "WARN"}
    }

    # ── Battery Report (powercfg) ──
    if($DoBatteryReport){
        QLog "Battery report (powercfg)..." "SCAN"
        $diag.BatteryReport = @{}
        try{
            $bat = Get-CimInstance Win32_Battery -EA SilentlyContinue
            if($bat){
                $tmpHtml = "$env:TEMP\WinDiag_BatteryReport.html"
                $null = & powercfg.exe /batteryreport /output $tmpHtml 2>&1
                if(Test-Path $tmpHtml){
                    $diag.BatteryReport.Available = $true
                    $diag.BatteryReport.ReportPath = $tmpHtml
                    $html = Get-Content $tmpHtml -Raw -EA SilentlyContinue
                    if($html -match 'DESIGN CAPACITY.*?(\d[\d,]*)\s*mWh'){$diag.BatteryReport.DesignCapacity = $matches[1] + " mWh"}
                    if($html -match 'FULL CHARGE CAPACITY.*?(\d[\d,]*)\s*mWh'){$diag.BatteryReport.FullChargeCapacity = $matches[1] + " mWh"}
                    if($html -match 'CYCLE COUNT.*?(\d+)'){$diag.BatteryReport.CycleCount = $matches[1]}
                    QLog "Battery report generated" "OK"
                }else{$diag.BatteryReport.Available = $false; QLog "powercfg failed" "WARN"}
            }else{$diag.BatteryReport.Available = $false; QLog "No battery" "INFO"}
        }catch{$diag.BatteryReport.Available = $false; QLog "Battery report error: $($_.Exception.Message)" "WARN"}
    }

    # ── Driver Check ──
    if($DoDrivers){
        QLog "Driver check..." "SCAN"
        $diag.Drivers = @()
        try{
            # Get all PnP devices with problems or all drivers
            Get-CimInstance Win32_PnPSignedDriver -EA SilentlyContinue | Where-Object { $_.DeviceName } | ForEach-Object {
                $drv = @{
                    Name=$_.DeviceName
                    Manufacturer=if($_.Manufacturer){$_.Manufacturer}else{"Unknown"}
                    DriverVersion=if($_.DriverVersion){$_.DriverVersion}else{"N/A"}
                    DriverDate=if($_.DriverDate){$_.DriverDate.ToString("yyyy-MM-dd")}else{"N/A"}
                    InfName=$_.InfName
                    IsSigned=if($_.IsSigned){"Yes"}else{"NO"}
                    DeviceClass=$_.DeviceClass
                }
                $diag.Drivers += $drv
            }
            # Check for problem devices
            $probDevices = Get-CimInstance Win32_PnPEntity -EA SilentlyContinue | Where-Object { $_.ConfigManagerErrorCode -ne 0 }
            foreach($pd in $probDevices){
                $errMap = @{1="Not configured";3="Corrupted driver";10="Cannot start";22="Disabled";28="No driver installed";31="Not working properly";43="Stopped responding"}
                $errDesc = $errMap[[int]$pd.ConfigManagerErrorCode]
                if(-not $errDesc){ $errDesc = "Error code $($pd.ConfigManagerErrorCode)" }
                # Find matching driver entry and add problem info
                $existing = $diag.Drivers | Where-Object { $_.Name -eq $pd.Name } | Select-Object -First 1
                if($existing){
                    $existing.Problem = $errDesc
                }else{
                    $diag.Drivers += @{
                        Name=$pd.Name; Manufacturer=if($pd.Manufacturer){$pd.Manufacturer}else{"Unknown"}
                        DriverVersion="N/A"; DriverDate="N/A"; InfName=""; IsSigned="N/A"
                        DeviceClass=$pd.PNPClass; Problem=$errDesc
                    }
                }
                QDbg "Problem device: $($pd.Name) - $errDesc"
            }
            QLog "$(@($diag.Drivers).Count) drivers, $(@($probDevices).Count) problems" "OK"
        }catch{QLog "Driver check error: $($_.Exception.Message)" "WARN"}
    }

    # ── BSOD / Crash Logs ──
    if($DoBSOD){
        QLog "BSOD / Crash logs..." "SCAN"
        $diag.BSOD = @()
        try{
            # Check for BugCheck events (Event ID 1001 from BugCheck, 6008 for unexpected shutdown)
            $bsodEvents = @()
            try{
                $bsodEvents += Get-WinEvent -FilterHashtable @{LogName='System';ProviderName='Microsoft-Windows-WER-SystemErrorReporting';Id=1001} -MaxEvents 20 -EA Stop
            }catch{}
            try{
                $bsodEvents += Get-WinEvent -FilterHashtable @{LogName='System';Id=6008} -MaxEvents 10 -EA Stop
            }catch{}
            # Also check for BugCheck registry
            try{
                $bc = Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control\CrashControl" -EA SilentlyContinue
                if($bc){
                    $diag.BSOD += @{Type="CrashDump";Info="DumpType=$($bc.CrashDumpEnabled) AutoReboot=$($bc.AutoReboot)";Time="Config"}
                    if($bc.LastCrashTime){ $diag.BSOD += @{Type="LastCrash";Info="Registry LastCrashTime: $($bc.LastCrashTime)";Time="Registry"} }
                }
            }catch{}
            # Parse events
            foreach($e in ($bsodEvents | Sort-Object TimeCreated -Descending | Select-Object -First 15)){
                $msg = if($e.Message){($e.Message -replace '\r?\n',' ')}else{"(no details)"}
                if($msg.Length -gt 300){$msg=$msg.Substring(0,300)}
                $diag.BSOD += @{Type=if($e.Id -eq 1001){"BSOD"}else{"UnexpectedShutdown"};Info=$msg;Time=$e.TimeCreated.ToString("yyyy-MM-dd HH:mm:ss")}
                QDbg "BSOD: [$($e.TimeCreated)] ID=$($e.Id)"
            }
            # Check minidump folder
            $miniDumpPath = "$env:SystemRoot\Minidump"
            if(Test-Path $miniDumpPath){
                $dumps = Get-ChildItem $miniDumpPath -Filter "*.dmp" -EA SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 10
                foreach($d in $dumps){
                    $diag.BSOD += @{Type="MiniDump";Info="$($d.Name) ($([math]::Round($d.Length/1KB,0)) KB)";Time=$d.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss")}
                }
                QDbg "Minidumps: $($dumps.Count) files"
            }
            # MEMORY.DMP
            $memDmp = "$env:SystemRoot\MEMORY.DMP"
            if(Test-Path $memDmp){
                $mf = Get-Item $memDmp
                $diag.BSOD += @{Type="FullDump";Info="MEMORY.DMP ($([math]::Round($mf.Length/1MB,0)) MB)";Time=$mf.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss")}
            }
            QLog "$(@($diag.BSOD).Count) crash entries found" $(if($diag.BSOD.Count -gt 1){"WARN"}else{"OK"})
        }catch{QLog "BSOD check error: $($_.Exception.Message)" "WARN"}
    }

    # ── RAM Test Logs (mdsched / MemoryDiagnostics) ──
    if($DoRAMTest){
        QLog "RAM test logs..." "SCAN"
        $diag.RAMTest = @{}
        try{
            # Check Windows Memory Diagnostic results
            $memEvents = @()
            try{
                $memEvents = Get-WinEvent -FilterHashtable @{LogName='System';ProviderName='Microsoft-Windows-MemoryDiagnostics-Results'} -MaxEvents 5 -EA Stop
            }catch{}
            if($memEvents.Count -gt 0){
                $diag.RAMTest.HasResults = $true
                $diag.RAMTest.Results = @()
                foreach($e in $memEvents){
                    $msg = if($e.Message){($e.Message -replace '\r?\n',' ')}else{"(no details)"}
                    if($msg.Length -gt 300){$msg=$msg.Substring(0,300)}
                    $diag.RAMTest.Results += @{Time=$e.TimeCreated.ToString("yyyy-MM-dd HH:mm:ss");Message=$msg;Level=$e.LevelDisplayName}
                    QDbg "RAM test: $($e.TimeCreated) - $($msg.Substring(0,[Math]::Min(80,$msg.Length)))"
                }
                QLog "$($memEvents.Count) RAM test results" "OK"
            } else {
                $diag.RAMTest.HasResults = $false
                $diag.RAMTest.Status = "No memory diagnostic results found (run: mdsched.exe)"
                QLog "No RAM test results found" "INFO"
            }
            # Also check for WHEA memory errors
            try{
                $wheaMem = Get-WinEvent -FilterHashtable @{LogName='System';ProviderName='Microsoft-Windows-WHEA-Logger'} -MaxEvents 10 -EA Stop |
                    Where-Object { $_.Message -match 'memory|corrected|uncorrected' }
                if($wheaMem.Count -gt 0){
                    $diag.RAMTest.WHEAErrors = $wheaMem.Count
                    QLog "WHEA memory errors: $($wheaMem.Count)" "WARN"
                }
            }catch{}
        }catch{QLog "RAM test check error: $($_.Exception.Message)" "WARN"}
    }

    # ── Performance Report ──
    if($DoPerfmon){
        QLog "Performance report..." "SCAN"
        $diag.Perfmon = @{Status="Collecting"; Warnings=@()}
        try{
            # Collect live performance counters (fast, no waiting)
            QDbg "Sampling CPU usage..."
            $cpu1 = (Get-CimInstance Win32_Processor | Measure-Object LoadPercentage -Average).Average
            Start-Sleep -Seconds 3
            $cpu2 = (Get-CimInstance Win32_Processor | Measure-Object LoadPercentage -Average).Average
            $cpuAvg = [math]::Round(($cpu1 + $cpu2) / 2, 1)
            $diag.Perfmon.CPUAvg = "$cpuAvg%"
            $diag.Perfmon.CPUPeak = "$([math]::Max($cpu1, $cpu2))%"
            QDbg "CPU avg: $cpuAvg%"

            # RAM
            $os = Get-CimInstance Win32_OperatingSystem
            $ramUsed = [math]::Round(($os.TotalVisibleMemorySize - $os.FreePhysicalMemory) / 1MB, 1)
            $ramFree = [math]::Round($os.FreePhysicalMemory / 1MB, 1)
            $diag.Perfmon.RAMUsedGB = "${ramUsed} GB"
            $diag.Perfmon.RAMFreeGB = "${ramFree} GB"
            $ramPct = [math]::Round(($ramUsed / ($os.TotalVisibleMemorySize / 1MB)) * 100, 0)
            QDbg "RAM used: $ramUsed GB ($ramPct%)"

            # Disk I/O via performance counters
            try{
                $diskRead = (Get-Counter '\PhysicalDisk(_Total)\Disk Read Bytes/sec' -SampleInterval 2 -MaxSamples 1 -EA Stop).CounterSamples[0].CookedValue
                $diskWrite = (Get-Counter '\PhysicalDisk(_Total)\Disk Write Bytes/sec' -SampleInterval 2 -MaxSamples 1 -EA Stop).CounterSamples[0].CookedValue
                $diag.Perfmon.DiskReadMBs = "$([math]::Round($diskRead/1MB, 2)) MB/s"
                $diag.Perfmon.DiskWriteMBs = "$([math]::Round($diskWrite/1MB, 2)) MB/s"
                QDbg "Disk: Read=$([math]::Round($diskRead/1MB,2)) MB/s Write=$([math]::Round($diskWrite/1MB,2)) MB/s"
            }catch{ $diag.Perfmon.DiskReadMBs = "N/A"; $diag.Perfmon.DiskWriteMBs = "N/A" }

            # Network I/O
            try{
                $netSent = (Get-Counter '\Network Interface(*)\Bytes Sent/sec' -SampleInterval 2 -MaxSamples 1 -EA Stop).CounterSamples | Measure-Object CookedValue -Sum
                $netRecv = (Get-Counter '\Network Interface(*)\Bytes Received/sec' -SampleInterval 2 -MaxSamples 1 -EA Stop).CounterSamples | Measure-Object CookedValue -Sum
                $diag.Perfmon.NetworkSentMBs = "$([math]::Round($netSent.Sum/1MB, 3)) MB/s"
                $diag.Perfmon.NetworkRecvMBs = "$([math]::Round($netRecv.Sum/1MB, 3)) MB/s"
                QDbg "Network: Sent=$([math]::Round($netSent.Sum/1MB,3)) MB/s Recv=$([math]::Round($netRecv.Sum/1MB,3)) MB/s"
            }catch{ $diag.Perfmon.NetworkSentMBs = "N/A"; $diag.Perfmon.NetworkRecvMBs = "N/A" }

            # Page faults
            try{
                $pageFaults = (Get-Counter '\Memory\Page Faults/sec' -SampleInterval 2 -MaxSamples 1 -EA Stop).CounterSamples[0].CookedValue
                $diag.Perfmon.PageFaultsPerSec = [math]::Round($pageFaults, 0)
                QDbg "Page faults/sec: $([math]::Round($pageFaults,0))"
            }catch{}

            # Context switches
            try{
                $ctxSwitches = (Get-Counter '\System\Context Switches/sec' -SampleInterval 2 -MaxSamples 1 -EA Stop).CounterSamples[0].CookedValue
                $diag.Perfmon.ContextSwitchesPerSec = [math]::Round($ctxSwitches, 0)
                QDbg "Context switches/sec: $([math]::Round($ctxSwitches,0))"
            }catch{}

            # Top CPU processes at snapshot
            $topCpuProcs = Get-Process | Sort-Object CPU -Descending | Select-Object -First 5 | ForEach-Object { "$($_.ProcessName) ($([math]::Round($_.CPU,1))s)" }
            $diag.Perfmon.TopCPUProcesses = $topCpuProcs -join ", "
            $topRAMProcs = Get-Process | Sort-Object WorkingSet64 -Descending | Select-Object -First 5 | ForEach-Object { "$($_.ProcessName) ($([math]::Round($_.WorkingSet64/1MB,0)) MB)" }
            $diag.Perfmon.TopRAMProcesses = $topRAMProcs -join ", "
            QDbg "Top CPU: $($diag.Perfmon.TopCPUProcesses)"
            QDbg "Top RAM: $($diag.Perfmon.TopRAMProcesses)"

            # Warnings
            if($cpuAvg -gt 80){ $diag.Perfmon.Warnings += "High CPU usage: $cpuAvg%" }
            if($ramPct -gt 85){ $diag.Perfmon.Warnings += "High RAM usage: $ramPct% ($ramUsed GB used)" }
            if($ramFree -lt 1){ $diag.Perfmon.Warnings += "Critical: less than 1 GB RAM free" }

            $diag.Perfmon.Status = "OK"
            $diag.Perfmon.Available = $true
            QLog "Performance data collected (CPU=$cpuAvg% RAM=$ramUsed GB)" "OK"
        }catch{
            $diag.Perfmon.Available = $false
            $diag.Perfmon.Status = "Error: $($_.Exception.Message)"
            QLog "Perfmon error: $($_.Exception.Message)" "WARN"
        }
    }

    # ── Established Connections ──
    if($DoNetwork){
        QLog "Established connections..." "SCAN"
        $diag.NetConnections = @()
        try{
            Get-NetTCPConnection -State Established -EA SilentlyContinue | Select-Object -First 50 | ForEach-Object{
                $pName = ""; try{$pName = (Get-Process -Id $_.OwningProcess -EA SilentlyContinue).ProcessName}catch{}
                $diag.NetConnections += @{LocalAddr="$($_.LocalAddress):$($_.LocalPort)";RemoteAddr="$($_.RemoteAddress):$($_.RemotePort)";Process=$pName;PID=$_.OwningProcess}
            }
        }catch{}
        QLog "Done" "OK"
    }

    # ── Routing Table ──
    if($DoNetwork){
        QLog "Routing table..." "SCAN"
        $diag.NetRoutes = @()
        try{
            Get-NetRoute -EA SilentlyContinue | Where-Object{$_.DestinationPrefix -ne "ff00::/8" -and $_.DestinationPrefix -ne "255.255.255.255/32"} | Select-Object -First 30 | ForEach-Object{
                $diag.NetRoutes += @{Destination=$_.DestinationPrefix;NextHop=$_.NextHop;Metric=$_.RouteMetric;Interface=$_.InterfaceAlias}
            }
        }catch{}
        QLog "Done" "OK"
    }

    # ── ARP Table ──
    if($DoNetwork){
        QLog "ARP table..." "SCAN"
        $diag.ArpTable = @()
        try{
            Get-NetNeighbor -EA SilentlyContinue | Where-Object{$_.State -ne "Unreachable" -and $_.IPAddress -notmatch "^ff|^224"} | Select-Object -First 30 | ForEach-Object{
                $diag.ArpTable += @{IP=$_.IPAddress;MAC=$_.LinkLayerAddress;State=$_.State;Interface=$_.InterfaceAlias}
            }
        }catch{}
        QLog "Done" "OK"
    }

    # ── Mapped Drives ──
    if($DoNetwork){
        QLog "Mapped drives..." "SCAN"
        $diag.MappedDrives = @()
        try{
            Get-PSDrive -PSProvider FileSystem -EA SilentlyContinue | Where-Object{$_.DisplayRoot} | ForEach-Object{
                $diag.MappedDrives += @{Drive="$($_.Name):";Path=$_.DisplayRoot;Free="{0:N1} GB"-f($_.Free/1GB);Used="{0:N1} GB"-f($_.Used/1GB)}
            }
            if($diag.MappedDrives.Count -eq 0){
                $nu = & net.exe use 2>&1 | Where-Object{$_ -match '^\s*(OK|Disconnected|Unavailable)\s'}
                foreach($line in $nu){ if($line -match '(\w:)\s+(.+)'){ $diag.MappedDrives += @{Drive=$matches[1];Path=$matches[2].Trim();Free="";Used=""} } }
            }
        }catch{}
        QLog "Done" "OK"
    }

    # ── Custom Checks ──
    if($DoCustomChecks -and $CustomChecksDir -and (Test-Path $CustomChecksDir)){
        QLog "Running custom checks from $CustomChecksDir..." "SCAN"
        $diag.CustomChecks = @()
        $scripts = Get-ChildItem $CustomChecksDir -Filter "*.ps1" -EA SilentlyContinue
        foreach($script in $scripts){
            QLog "  Running: $($script.Name)" "SCAN"
            try{
                $output = & $script.FullName 2>&1 | Out-String
                $diag.CustomChecks += @{Script=$script.Name;Output=$output.Trim();Status="OK"}
                QDbg "Custom check $($script.Name): OK"
            }catch{
                $diag.CustomChecks += @{Script=$script.Name;Output=$_.Exception.Message;Status="ERROR"}
                QLog "  Error in $($script.Name): $($_.Exception.Message)" "WARN"
            }
        }
        QLog "$(@($diag.CustomChecks).Count) custom checks completed" "OK"
    }

    } catch { Q("ERROR:Scan exception: $($_.Exception.Message)") }

    # Send results as JSON — split into base + extended
    $extData = @{}
    foreach($k in @("Hosts","UserInfo","ExternalIP","Hotfixes","RemoteTools","ChkdskLogs","BatteryReport","AllServices","NetConnections","NetRoutes","ArpTable","MappedDrives","Drivers","SmartDetails","CustomChecks","BSOD","RAMTest","Perfmon")){
        if($diag.ContainsKey($k)){ $extData[$k] = $diag[$k]; $diag.Remove($k) }
    }
    $extJson = $extData | ConvertTo-Json -Depth 5 -Compress
    Q("EXT:$extJson")
    $json = $diag | ConvertTo-Json -Depth 5 -Compress
    Q("DONE:SCAN:$json")
}
#endregion

#region AI Worker ScriptBlock
$AiWorkerScript = {
    function Q([string]$m){ $MsgQueue.Enqueue($m) }

    Q("AI:[AI] Connecting to $Model...")
    try {
        $body = @{
            model=$Model; stream=$false
            messages=@(@{role="system";content=$SysPrompt},@{role="user";content=$DiagText})
            options=@{temperature=$AiTemp;num_predict=$AiMaxTokens}
        } | ConvertTo-Json -Depth 5

        Q("AI:[AI] Sending to $Model... (this may take several minutes)")

        $r = Invoke-RestMethod -Uri "$OllamaUrl/api/chat" -Method Post `
            -Body ([System.Text.Encoding]::UTF8.GetBytes($body)) `
            -ContentType "application/json; charset=utf-8" -TimeoutSec 900 -EA Stop

        $result = $r.message.content
        foreach($line in ($result -split '\r?\n')){ Q("AI:$line") }
        Q("[$(Get-Date -F 'yyyy-MM-dd HH:mm:ss')][OK] AI analysis complete!")
        Q("DONE:AI:$result")
    } catch {
        Q("AI:[ERROR] $($_.Exception.Message)")
        Q("ERROR:AI failed: $($_.Exception.Message)")
        Q("DONE:AI:")
    }
}
#endregion

#region Event Handlers
$btnScan.Add_Click({
    Set-UiBusy $true; $txtLog.Document.Blocks.Clear()
    Ui-SetStatus "Scanning..."; $Global:DiagData = $null; $Global:ExtJson = $null
    $Script:Verbose = $tglVerbose.IsChecked

    $evtHrsVal = if($Global:EventHours){$Global:EventHours}else{24}
    $params = @{ VerboseMode = $Script:Verbose; EventHours = $evtHrsVal; CustomChecksDir = $Global:CustomChecksDir; DoCustomChecks = (Test-Path $Global:CustomChecksDir) }
    foreach($n in $chks.Keys){ $params["Do$n"] = [bool]$chks[$n].IsChecked }

    Start-BackgroundJob -Job $ScanWorkerScript -Params $params
})

$btnAI.Add_Click({
    $model = Get-SelModel
    if(-not $model){[Windows.MessageBox]::Show("Select a model.","No Model","OK","Warning");return}
    if(-not $Global:DiagData){[Windows.MessageBox]::Show("Run scan first.","No Data","OK","Warning");return}
    if(-not(Test-OllamaRunning)){[Windows.MessageBox]::Show("Ollama not running.","Error","OK","Warning");return}

    # Check if needs download
    $ms=Get-OllamaModels
    if(-not($ms|Where-Object{$_-like "$model*"})){
        $a=[Windows.MessageBox]::Show("Model '$model' not installed. Download first?`n`nThis may take several minutes.","Download Required","YesNo","Question")
        if($a-ne"Yes"){return}
        Set-UiBusy $true; Ui-SetStatus "Downloading $model..."
        Ui-Log "Downloading '$model' before AI analysis from $($Global:OllamaUrl)/api/pull" "AI"
        # Streaming download with progress
        try{
            $body = "{`"name`":`"$model`"}"
            $req = [System.Net.HttpWebRequest]::Create("$($Global:OllamaUrl)/api/pull")
            $req.Method = "POST"; $req.ContentType = "application/json"; $req.Timeout = 1800000
            $bb = [System.Text.Encoding]::UTF8.GetBytes($body); $req.ContentLength = $bb.Length
            $rs = $req.GetRequestStream(); $rs.Write($bb,0,$bb.Length); $rs.Close()
            $resp = $req.GetResponse(); $rd2 = [System.IO.StreamReader]::new($resp.GetResponseStream())
            $ok = $false; $lastPctAi = -1; $t0Ai = Get-Date
            while(-not $rd2.EndOfStream){
                $ln = $rd2.ReadLine(); if(-not $ln){continue}
                try{$j = $ln | ConvertFrom-Json -EA SilentlyContinue
                    if($j.status -eq "success"){$ok=$true; Ui-Log "Downloaded '$model'! ($('{0:mm\:ss}' -f ((Get-Date)-$t0Ai)))" "OK"}
                    elseif($j.total -and $j.completed){
                        $pct=[math]::Round(($j.completed/$j.total)*100,0)
                        if($pct -ne $lastPctAi){
                            $spd=[math]::Round($j.completed/1MB/[Math]::Max(1,((Get-Date)-$t0Ai).TotalSeconds),1)
                            $digest = if($j.digest){$j.digest.Substring(0,[Math]::Min(19,$j.digest.Length))}else{""}
                            Ui-SetStatus "Downloading $model... $pct% @ $spd MB/s"
                            if($pct % 10 -eq 0 -or $pct -eq 1){
                                Ui-Log "$($j.status) $digest`: $pct% ($([math]::Round($j.completed/1MB,0))/$([math]::Round($j.total/1MB,0)) MB) @ $spd MB/s" "INFO"
                            }
                            $lastPctAi=$pct
                        }
                    }
                    elseif($j.status){ Ui-Log "$($j.status)" "INFO" }
                }catch{}
                [Windows.Forms.Application]::DoEvents()
            }
            $rd2.Close(); $resp.Close()
            if(-not $ok){Ui-Log "Download may have failed" "WARN"}
            Refresh-Ollama
        }catch{Ui-Log "Download failed: $($_.Exception.Message)" "ERROR";Set-UiBusy $false;return}
    }

    $Script:SelectedModel = $model
    Set-UiBusy $true; $txtAI.Document.Blocks.Clear()
    Ui-SetStatus "AI analyzing... (may take several minutes)"

    # Build diagnostic text
    $D = $Global:DiagData
    $sb = [System.Text.StringBuilder]::new()
    if($D.SystemInfo.ComputerName){[void]$sb.AppendLine("=== SYSTEM ==="); foreach($k in @("ComputerName","OS","CPU","RAMTotal","RAMFree","RAMUsage","Uptime","Manufacturer","Model")){if($D.SystemInfo.$k){[void]$sb.AppendLine("$k`: $($D.SystemInfo.$k)")}}}
    if($D.Battery.Status){[void]$sb.AppendLine("`n=== BATTERY ==="); foreach($k in @("Status","EstCharge","Health","Chemistry")){if($D.Battery.$k){[void]$sb.AppendLine("$k`: $($D.Battery.$k)")}}}
    if($D.Disks){[void]$sb.AppendLine("`n=== DISKS ==="); foreach($d in $D.Disks){[void]$sb.AppendLine("$($d.Drive) $($d.Label): $($d.FreeGB)GB free/$($d.TotalGB)GB - $($d.Status)")}}
    if($D.Events){
        $cc=@($D.Events|Where-Object{$_.Level-eq"Critical"}).Count;$ec=@($D.Events|Where-Object{$_.Level-eq"Error"}).Count;$wc=@($D.Events|Where-Object{$_.Level-eq"Warning"}).Count
        [void]$sb.AppendLine("`n=== EVENTS === Critical:$cc Error:$ec Warn:$wc")
        $D.Events|Group-Object{$_.Source+"|"+$_.EventID}|Sort-Object Count -Descending|Select-Object -First 20|ForEach-Object{$s=$_.Group[0];[void]$sb.AppendLine("[$($s.Level)] $($s.Source) ID:$($s.EventID) x$($_.Count) - $($s.Message)")}
    }
    if($D.Services.Count-gt 0){[void]$sb.AppendLine("`n=== STOPPED SERVICES ===");foreach($s in $D.Services){[void]$sb.AppendLine("$($s.DisplayName) ($($s.Name))")}}
    if($D.Network.Adapters){[void]$sb.AppendLine("`n=== NETWORK ===");foreach($a in $D.Network.Adapters){[void]$sb.AppendLine("$($a.Adapter) IP=$($a.IP) GW=$($a.Gateway) DNS=$($a.DNS) $($a.Speed)")};[void]$sb.AppendLine("DNS: $($D.Network.DNSTests-join' | ')")}
    if($D.Security.Firewall){[void]$sb.AppendLine("`n=== SECURITY ===");foreach($k in @("Firewall","Defender","RealTime","Definitions","UAC")){if($D.Security.$k){[void]$sb.AppendLine("$k`: $($D.Security.$k)")}}}
    if($D.Updates){[void]$sb.AppendLine("`n=== UPDATES === Pending: $($D.Updates.PendingCount)");foreach($u in $D.Updates.Pending){[void]$sb.AppendLine("  [$($u.Severity)] $($u.Title)")}}
    if($D.Processes.TopRAM){[void]$sb.AppendLine("`n=== TOP RAM PROCESSES ===");foreach($p in $D.Processes.TopRAM){[void]$sb.AppendLine("$($p.Name) PID:$($p.PID) $($p.RAM)MB RAM $($p.CPU)s CPU")}}
    if($D.Startup){[void]$sb.AppendLine("`n=== STARTUP ===");foreach($s in $D.Startup){[void]$sb.AppendLine("$($s.Name) - $($s.Command) [$($s.Location)]")}}
    if($D.ScheduledTasks){[void]$sb.AppendLine("`n=== SCHEDULED TASKS ===");foreach($t in $D.ScheduledTasks){[void]$sb.AppendLine("$($t.Name) [$($t.State)] Last:$($t.LastRun)")}}
    if($D.Autoruns){[void]$sb.AppendLine("`n=== AUTORUNS ===");foreach($a in $D.Autoruns){[void]$sb.AppendLine("[$($a.Type)] $($a.Name) - $($a.Location)")}}
    if($D.Integrity.LastSFC){[void]$sb.AppendLine("`n=== INTEGRITY ===");[void]$sb.AppendLine("SFC: $($D.Integrity.LastSFC)");[void]$sb.AppendLine("Reboot: $($D.Integrity.RebootPending)")}
    if($D.SearchIndex.Service){[void]$sb.AppendLine("`n=== SEARCH INDEX ===");foreach($k in @("Service","Status","Items","Size")){if($D.SearchIndex.$k){[void]$sb.AppendLine("$k`: $($D.SearchIndex.$k)")}}}
    # Extended data
    $EX = $null
    if($Global:ExtJson){ try{ $EX = $Global:ExtJson | ConvertFrom-Json -EA Stop }catch{} }
    if($EX -and $EX.Hosts.Entries){$he=@($EX.Hosts.Entries);if($he.Count -gt 0){[void]$sb.AppendLine("`n=== HOSTS FILE ($($he.Count) entries) ===");foreach($h in $he){[void]$sb.AppendLine("$($h.IP) $($h.Hostname)")}}}
    if($EX -and $EX.UserInfo.LocalAccounts){[void]$sb.AppendLine("`n=== USER ACCOUNTS ===");foreach($u in $EX.UserInfo.LocalAccounts){[void]$sb.AppendLine("$($u.Name) Enabled=$($u.Enabled) PwdReq=$($u.PasswordRequired) LastLogon=$($u.LastLogon)")}}
    if($EX -and $EX.ExternalIP.IP){[void]$sb.AppendLine("`n=== EXTERNAL IP ===");[void]$sb.AppendLine($EX.ExternalIP.IP)}
    if($EX -and $EX.Hotfixes){$hf=@($EX.Hotfixes);if($hf.Count -gt 0){[void]$sb.AppendLine("`n=== INSTALLED HOTFIXES ($($hf.Count)) ===");foreach($h in $hf){[void]$sb.AppendLine("$($h.HotFixID) $($h.InstalledOn) $($h.Description)")}}}
    if($EX -and $EX.RemoteTools){[void]$sb.AppendLine("`n=== REMOTE TOOLS ===");foreach($k in @("TeamViewer","AnyDesk","RustDesk")){if($EX.RemoteTools.$k){[void]$sb.AppendLine("$k`: $($EX.RemoteTools.$k)")}}}
    if($EX -and $EX.ChkdskLogs){$cl=@($EX.ChkdskLogs);if($cl.Count -gt 0){[void]$sb.AppendLine("`n=== CHKDSK LOGS ===");foreach($c in $cl){[void]$sb.AppendLine("[$($c.Time)] $($c.Message)")}}}
    if($EX -and ($EX.BatteryReport.Available -eq $true -or $EX.BatteryReport.Available -eq "True")){[void]$sb.AppendLine("`n=== BATTERY REPORT ===");foreach($k in @("DesignCapacity","FullChargeCapacity","CycleCount")){if($EX.BatteryReport.$k){[void]$sb.AppendLine("$k`: $($EX.BatteryReport.$k)")}}}
    # SMART details
    if($EX -and $EX.SmartDetails){$sda=@($EX.SmartDetails);if($sda.Count -gt 0){[void]$sb.AppendLine("`n=== SMART DETAILS ===");foreach($sd in $sda){[void]$sb.AppendLine("$($sd.Model) | Type=$($sd.MediaType) Bus=$($sd.BusType) Health=$($sd.Health) Temp=$($sd.Temperature) PowerOn=$($sd.PowerOnHours)h ReadErr=$($sd.ReadErrors) WriteErr=$($sd.WriteErrors) Wear=$($sd.Wear)")}}}
    # Drivers
    if($EX -and $EX.Drivers){$dra=@($EX.Drivers);$probDrv=@($dra|Where-Object{$_.Problem -and $_.Problem -ne "None" -and $_.Problem -ne "0"});if($probDrv.Count -gt 0){[void]$sb.AppendLine("`n=== PROBLEM DRIVERS ===");foreach($pd in $probDrv){[void]$sb.AppendLine("$($pd.Name) - $($pd.Problem) (Driver: $($pd.DriverVersion) Date: $($pd.DriverDate))")}}}
    # Custom checks
    if($EX -and $EX.CustomChecks){$cca=@($EX.CustomChecks);if($cca.Count -gt 0){[void]$sb.AppendLine("`n=== CUSTOM CHECKS ===");foreach($cc2 in $cca){[void]$sb.AppendLine("--- $($cc2.Script) [$($cc2.Status)] ---");[void]$sb.AppendLine($cc2.Output)}}}
    # BSOD
    if($EX -and $EX.BSOD){$ba=@($EX.BSOD);if($ba.Count -gt 0){[void]$sb.AppendLine("`n=== BSOD / CRASH LOGS ===");foreach($b in $ba){[void]$sb.AppendLine("[$($b.Time)] $($b.Type): $($b.Info)")}}}
    # RAM tests
    if($EX -and $EX.RAMTest){
        [void]$sb.AppendLine("`n=== RAM DIAGNOSTICS ===")
        if($EX.RAMTest.HasResults -eq $true){foreach($r in @($EX.RAMTest.Results)){[void]$sb.AppendLine("[$($r.Time)] $($r.Level): $($r.Message)")}}
        else{[void]$sb.AppendLine($EX.RAMTest.Status)}
        if($EX.RAMTest.WHEAErrors){[void]$sb.AppendLine("WHEA Memory Errors: $($EX.RAMTest.WHEAErrors)")}
    }
    if($EX -and $EX.Perfmon -and ($EX.Perfmon.Available -eq $true)){
        [void]$sb.AppendLine("`n=== PERFORMANCE SNAPSHOT ===")
        foreach($prop in @("CPUAvg","CPUPeak","RAMUsedGB","RAMFreeGB","DiskReadMBs","DiskWriteMBs","PageFaultsPerSec","ContextSwitchesPerSec")){
            $val = $EX.Perfmon.$prop; if($null -ne $val){[void]$sb.AppendLine("$prop`: $val")}
        }
        if($EX.Perfmon.TopCPUProcesses){[void]$sb.AppendLine("Top CPU: $($EX.Perfmon.TopCPUProcesses)")}
        if($EX.Perfmon.TopRAMProcesses){[void]$sb.AppendLine("Top RAM: $($EX.Perfmon.TopRAMProcesses)")}
        if($EX.Perfmon.Warnings){foreach($w in @($EX.Perfmon.Warnings)){[void]$sb.AppendLine("WARNING: $w")}}
    }

    [void]$sb.AppendLine("`n=== SUGGEST ===")
    [void]$sb.AppendLine("If integrity issues: DISM /Online /Cleanup-Image /RestoreHealth then sfc /scannow")
    [void]$sb.AppendLine("If search issues: net stop WSearch && rebuild index")

    $sysPrompt = @"
You are an expert Windows system administrator. Analyze the diagnostics data and provide a structured actionable report in English.

FORMAT YOUR RESPONSE EXACTLY AS FOLLOWS:

## EXECUTIVE SUMMARY
(1-2 sentences overall system health assessment)

## CRITICAL ISSUES
(Issues requiring immediate attention. For each: describe the problem, impact, and exact repair command.)

## WARNINGS
(Non-critical but notable findings. For each: describe and suggest remediation.)

## REPAIR COMMANDS
(Consolidated list of PowerShell/CMD commands to fix identified issues, ready to copy-paste.)

## OPTIMIZATION SUGGESTIONS
(Performance improvements, cleanup, best practices.)

## PREVENTIVE MEASURES
(Recommendations to prevent future issues.)

RULES:
- Be specific with commands - provide exact PowerShell or CMD syntax.
- Group related event log entries by source/ID instead of listing each one.
- Prioritize findings by severity (critical > error > warning).
- If system integrity issues exist, always suggest: DISM /Online /Cleanup-Image /RestoreHealth followed by sfc /scannow
- If search/indexing issues exist, suggest rebuild steps.
- If SMART issues detected, flag them prominently.
- If driver problems found, provide update/reinstall commands.
- Keep it concise and actionable. No filler text.
"@

    $aiTempVal = if($Global:AiTemp){$Global:AiTemp}else{0.3}
    $aiMaxTokVal = if($Global:AiMaxTokens){$Global:AiMaxTokens}else{4096}
    $params = @{ Model=$model; SysPrompt=$sysPrompt; DiagText=$sb.ToString(); AiTemp=$aiTempVal; AiMaxTokens=$aiMaxTokVal }
    Start-BackgroundJob -Job $AiWorkerScript -Params $params
})

$btnReport.Add_Click({
    if(-not $Global:DiagData){return}
    Ui-SetStatus "Generating report..."
    $rpDir = if($Global:ReportPath -and (Test-Path $Global:ReportPath)){$Global:ReportPath}else{$Global:ScriptDir}
    $rp = "$rpDir\WinDiag-AI_Report_$(Get-Date -F 'yyyyMMdd_HHmmss').html"
    try {
        Add-Type -AssemblyName System.Web
        $D = $Global:DiagData
        $Global:E = $null
        if($Global:ExtJson){
            try{ $Global:E = $Global:ExtJson | ConvertFrom-Json -EA Stop }catch{ Ui-Log "ExtJSON parse error: $($_.Exception.Message)" "ERROR" }
        }
        $aiT = if($Global:AiAnalysis){$Global:AiAnalysis}else{"AI analysis not performed."}
        $E = $Global:E
        # Convert AI text to HTML
        $aiLines = ($aiT -split '\r?\n')
        $ahParts = [System.Collections.Generic.List[string]]::new()
        foreach($aiLine in $aiLines){
            $hl = [System.Web.HttpUtility]::HtmlEncode($aiLine)
            if($hl -match '^#{1,3}\s+(.+)$'){
                $ahParts.Add("<h3 style='color:#60a5fa;margin:15px 0 8px 0;'>$($matches[1])</h3>")
                continue
            }
            $hl = $hl -replace '\*\*([^*]+)\*\*','<strong>$1</strong>'
            $hl = $hl -replace '`([^`]+)`','<code>$1</code>'
            $ahParts.Add("$hl<br>")
        }
        $ah = $ahParts -join "`n"
        $cc=@($D.Events|Where-Object{$_.Level-eq"Critical"}).Count;$ec=@($D.Events|Where-Object{$_.Level-eq"Error"}).Count
        $wc=@($D.Events|Where-Object{$_.Level-eq"Warning"}).Count
        $sc2=@($D.Services).Count; $pu=$D.Updates.PendingCount

        function KvRows($h){
            $o=""
            if($null -eq $h){return $o}
            if($h -is [hashtable]){
                foreach($kv in $h.GetEnumerator()){$o+="<tr><td>$([System.Web.HttpUtility]::HtmlEncode($kv.Key))</td><td>$([System.Web.HttpUtility]::HtmlEncode($kv.Value))</td></tr>"}
            } elseif($h -is [PSCustomObject]){
                foreach($p in $h.PSObject.Properties){$o+="<tr><td>$([System.Web.HttpUtility]::HtmlEncode($p.Name))</td><td>$([System.Web.HttpUtility]::HtmlEncode($p.Value))</td></tr>"}
            }
            $o
        }
        function AsArray($v){ if($null -eq $v){return @()} if($v -is [array]){return $v} return @($v) }

        $sysInfoHtml  = KvRows $D.SystemInfo
        $hardwareHtml = KvRows $D.Hardware
        $batteryHtml  = KvRows $D.Battery
        $securityHtml = KvRows $D.Security
        $integrityHtml= KvRows $D.Integrity
        $searchHtml   = KvRows $D.SearchIndex

        $disksArr     = AsArray $D.Disks
        $eventsArr    = AsArray $D.Events
        $servicesArr  = AsArray $D.Services
        $adaptersArr  = AsArray $D.Network.Adapters
        $dnsArr       = AsArray $D.Network.DNSTests
        $topRamArr    = AsArray $D.Processes.TopRAM
        $topCpuArr    = AsArray $D.Processes.TopCPU
        $startupArr   = AsArray $D.Startup
        $tasksArr     = AsArray $D.ScheduledTasks
        $autorunsArr  = AsArray $D.Autoruns
        $softwareArr  = AsArray $D.Software
        $updatesArr   = AsArray $D.Updates.Pending

        $diskRows=""
        foreach($d in $disksArr){
            $cls=if($d.Status-eq"CRITICAL"){"sc"}elseif($d.Status-eq"WARNING"){"sw"}else{"sok"}
            $diskRows+="<tr><td>$($d.Drive)</td><td>$($d.Label)</td><td>$($d.TotalGB)GB</td><td>$($d.FreeGB)GB</td><td>$($d.FreePercent)</td><td class='$cls'>$($d.Status)</td></tr>"
        }
        $eventRows=""
        foreach($e in ($eventsArr|Select-Object -First 30)){
            $lc=switch($e.Level){"Critical"{"sc"};"Error"{"sw"};default{""}}
            $emsg=[System.Web.HttpUtility]::HtmlEncode($e.Message)
            if($emsg.Length-gt 150){$emsg=$emsg.Substring(0,150)+"..."}
            $eventRows+="<tr><td>$($e.Time)</td><td>$($e.Log)</td><td class='$lc'>$($e.Level)</td><td>$([System.Web.HttpUtility]::HtmlEncode($e.Source))</td><td>$($e.EventID)</td><td style='font-size:12px'>$emsg</td></tr>"
        }
        $svcRows=""
        foreach($s in $servicesArr){$svcRows+="<tr><td>$([System.Web.HttpUtility]::HtmlEncode($s.DisplayName))</td><td>$($s.Name)</td><td class='sw'>$($s.State)</td><td>$($s.RunAs)</td></tr>"}
        $netRows=""
        foreach($a in $adaptersArr){$netRows+="<tr><td>$($a.Adapter)</td><td>$($a.IP)</td><td>$($a.Gateway)</td><td>$($a.DNS)</td><td>$($a.Speed)</td><td>$($a.MAC)</td></tr>"}
        $dnsText=($dnsArr-join' | ')
        $ramRows=""
        foreach($p in $topRamArr){$ramRows+="<tr><td>$([System.Web.HttpUtility]::HtmlEncode($p.Name))</td><td>$($p.PID)</td><td>$($p.RAM)</td><td>$($p.CPU)</td></tr>"}
        $cpuRows=""
        foreach($p in $topCpuArr){$cpuRows+="<tr><td>$([System.Web.HttpUtility]::HtmlEncode($p.Name))</td><td>$($p.PID)</td><td>$($p.CPU)</td><td>$($p.RAM)</td></tr>"}
        $startupRows=""
        foreach($s in $startupArr){$startupRows+="<tr><td>$([System.Web.HttpUtility]::HtmlEncode($s.Name))</td><td style='font-size:11px'>$([System.Web.HttpUtility]::HtmlEncode($s.Command))</td><td style='font-size:11px'>$($s.Location)</td></tr>"}
        $taskRows=""
        foreach($t in $tasksArr){$taskRows+="<tr><td>$([System.Web.HttpUtility]::HtmlEncode($t.Name))</td><td style='font-size:11px'>$($t.Path)</td><td>$($t.State)</td><td>$($t.LastRun)</td><td style='font-size:11px'>$([System.Web.HttpUtility]::HtmlEncode($t.Author))</td></tr>"}
        $autorunRows=""
        foreach($a in $autorunsArr){$autorunRows+="<tr><td>$($a.Type)</td><td>$([System.Web.HttpUtility]::HtmlEncode($a.Name))</td><td style='font-size:11px'>$([System.Web.HttpUtility]::HtmlEncode($a.Location))</td></tr>"}
        $swRows=""
        foreach($s in $softwareArr){$swRows+="<tr><td>$([System.Web.HttpUtility]::HtmlEncode($s.Name))</td><td>$($s.Version)</td><td>$([System.Web.HttpUtility]::HtmlEncode($s.Publisher))</td></tr>"}
        $updRows=""
        foreach($u in $updatesArr){$updRows+="<tr><td>$([System.Web.HttpUtility]::HtmlEncode($u.Title))</td><td>$($u.KB)</td><td>$($u.Severity)</td></tr>"}

        $battSec=if($D.Battery.Status -and $D.Battery.Status -ne "N/A"){"<details class='sec' open><summary><h2>&#x1F50B; Battery</h2></summary><table class='kv'>$batteryHtml</table></details>"}else{""}
        $hwSec=if($D.Hardware.RAM -or $D.Hardware.GPU){"<details class='sec' open><summary><h2>&#x1F4BB; Hardware</h2></summary><table class='kv'>$hardwareHtml</table></details>"}else{""}
        $taskSec=if($tasksArr.Count-gt 0){"<details class='sec'><summary><h2>&#x23F0; Scheduled Tasks ($($tasksArr.Count))</h2></summary><table><tr><th>Name</th><th>Path</th><th>State</th><th>Last Run</th><th>Author</th></tr>$taskRows</table></details>"}else{""}
        $autoSec=if($autorunsArr.Count-gt 0){"<details class='sec'><summary><h2>&#x1F527; Autoruns ($($autorunsArr.Count))</h2></summary><table><tr><th>Type</th><th>Name</th><th>Location</th></tr>$autorunRows</table></details>"}else{""}
        $swSec=if($softwareArr.Count-gt 0){"<details class='sec'><summary><h2>&#x1F4E6; Installed Software ($($softwareArr.Count))</h2></summary><table><tr><th>Name</th><th>Version</th><th>Publisher</th></tr>$swRows</table></details>"}else{""}
        $updSec=""
        if($pu -gt 0){$updSec="<details class='sec' open><summary><h2>&#x1F4E5; Pending Updates ($pu)</h2></summary><table><tr><th>Title</th><th>KB</th><th>Severity</th></tr>$updRows</table></details>"}
        elseif($pu -eq 0){$updSec="<details class='sec'><summary><h2>&#x1F4E5; Windows Updates</h2></summary><p class='sok'>System is up to date</p></details>"}
        else{$updSec="<details class='sec'><summary><h2>&#x1F4E5; Windows Updates</h2></summary><p class='sw'>Could not check updates</p></details>"}

        # SMART details section
        $smartSec = ""
        $smartArr2 = AsArray $Global:E.SmartDetails
        if($smartArr2.Count -gt 0){
            $smartRows=""
            foreach($sd in $smartArr2){
                $hCls=if($sd.Health -ne "Healthy"){"sw"}else{"sok"}
                $smartRows+="<tr><td>$([System.Web.HttpUtility]::HtmlEncode($sd.Model))</td><td>$($sd.MediaType)</td><td>$($sd.BusType)</td><td>$($sd.Size)</td><td class='$hCls'>$($sd.Health)</td><td>$($sd.Temperature)</td><td>$($sd.PowerOnHours)</td><td>$($sd.ReadErrors)</td><td>$($sd.WriteErrors)</td><td>$($sd.Wear)</td></tr>"
            }
            $smartSec="<details class='sec' open><summary><h2>&#x1F4BE; SMART Details ($($smartArr2.Count))</h2></summary><table><tr><th>Model</th><th>Type</th><th>Bus</th><th>Size</th><th>Health</th><th>Temp</th><th>Power-On Hours</th><th>Read Errors</th><th>Write Errors</th><th>Wear</th></tr>$smartRows</table></details>"
        }

        # Drivers section
        $driverSec = ""
        $drvArr2 = AsArray $Global:E.Drivers
        if($drvArr2.Count -gt 0){
            $probDrv2 = @($drvArr2 | Where-Object { $_.Problem -and $_.Problem -ne "None" -and $_.Problem -ne "0" })
            if($probDrv2.Count -gt 0){
                $drvRows=""
                foreach($drv in $probDrv2){
                    $drvRows+="<tr><td>$([System.Web.HttpUtility]::HtmlEncode($drv.Name))</td><td>$($drv.Manufacturer)</td><td>$($drv.DriverVersion)</td><td>$($drv.DriverDate)</td><td class='sw'>$($drv.Problem)</td><td>$($drv.IsSigned)</td></tr>"
                }
                $driverSec="<details class='sec' open><summary><h2>&#x1F698; Problem Drivers ($($probDrv2.Count))</h2></summary><table><tr><th>Device</th><th>Manufacturer</th><th>Version</th><th>Date</th><th>Problem</th><th>Signed</th></tr>$drvRows</table></details>"
            }
            # All drivers (collapsed)
            $allDrvRows=""
            foreach($drv in ($drvArr2 | Select-Object -First 100)){
                $pCls=if($drv.Problem -and $drv.Problem -ne "None" -and $drv.Problem -ne "0"){"sw"}else{""}
                $allDrvRows+="<tr><td>$([System.Web.HttpUtility]::HtmlEncode($drv.Name))</td><td>$($drv.Manufacturer)</td><td>$($drv.DriverVersion)</td><td>$($drv.DriverDate)</td><td class='$pCls'>$(if($drv.Problem){$drv.Problem}else{'OK'})</td><td>$($drv.IsSigned)</td><td>$($drv.DeviceClass)</td></tr>"
            }
            $driverSec+="<details class='sec'><summary><h2>&#x1F698; All Drivers ($($drvArr2.Count))</h2></summary><table><tr><th>Device</th><th>Manufacturer</th><th>Version</th><th>Date</th><th>Status</th><th>Signed</th><th>Class</th></tr>$allDrvRows</table></details>"
        }

        # Custom checks section
        $customSec = ""
        $ccArr = AsArray $Global:E.CustomChecks
        if($ccArr.Count -gt 0){
            $ccContent=""
            foreach($cc2 in $ccArr){
                $stCls=if($cc2.Status -eq "ERROR"){"sw"}else{"sok"}
                $ccContent+="<div style='margin:12px 0;padding:12px;background:#0f172a;border-radius:6px'><strong class='$stCls'>$([System.Web.HttpUtility]::HtmlEncode($cc2.Script)) [$($cc2.Status)]</strong><pre style='margin-top:8px;white-space:pre-wrap;font-size:11px'>$([System.Web.HttpUtility]::HtmlEncode($cc2.Output))</pre></div>"
            }
            $customSec="<details class='sec'><summary><h2>&#x1F4DC; Custom Checks ($($ccArr.Count))</h2></summary>$ccContent</details>"
        }

        # BSOD section
        $bsodSec = ""
        $bsodArr2 = AsArray $Global:E.BSOD
        if($bsodArr2.Count -gt 0){
            $bsodRows=""
            foreach($b in $bsodArr2){
                $bCls = if($b.Type -eq "BSOD"){"sc"}elseif($b.Type -match "Dump|Crash"){"sw"}else{""}
                $bInfo = [System.Web.HttpUtility]::HtmlEncode($b.Info)
                if($bInfo.Length -gt 200){ $bInfo = $bInfo.Substring(0,200) + "..." }
                $bsodRows += "<tr><td class='$bCls'>$($b.Type)</td><td>$($b.Time)</td><td style='font-size:12px'>$bInfo</td></tr>"
            }
            $bsodReal = @($bsodArr2 | Where-Object { $_.Type -eq "BSOD" -or $_.Type -eq "MiniDump" -or $_.Type -eq "FullDump" })
            $bsodHeader = if($bsodReal.Count -gt 0){"&#x26A0; BSOD / Crash Logs ($($bsodArr2.Count) entries)"}else{"BSOD / Crash Logs (No crashes)"}
            $bsodOpenAttr = if($bsodReal.Count -gt 0){" open"}else{""}
            $bsodSec = "<details class='sec'$bsodOpenAttr><summary><h2>$bsodHeader</h2></summary><table><tr><th>Type</th><th>Time</th><th>Details</th></tr>$bsodRows</table></details>"
        } else {
            $bsodSec = "<details class='sec'><summary><h2>&#x1F4A4; BSOD / Crash Logs</h2></summary><p class='sok'>No crash dumps or BSOD events detected</p></details>"
        }

        # RAM Test section
        $ramTestSec = ""
        $rtData = $Global:E.RAMTest
        if($rtData){
            $rtStatus = if($rtData.Status){$rtData.Status}else{"Not run"}
            $rtHasResults = ($rtData.HasResults -eq $true -or $rtData.HasResults -eq "True")
            $rtWHEA = if($rtData.WHEAErrors -and $rtData.WHEAErrors -gt 0){$rtData.WHEAErrors}else{0}
            $rtContent = "<table class='kv'><tr><td>Status</td><td>$rtStatus</td></tr>"
            if($rtWHEA -gt 0){ $rtContent += "<tr><td>WHEA Memory Errors</td><td class='sc'>$rtWHEA</td></tr>" }
            if($rtData.LastTestDate){ $rtContent += "<tr><td>Last Test</td><td>$($rtData.LastTestDate)</td></tr>" }
            if($rtData.TestResult){ $rtContent += "<tr><td>Result</td><td>$([System.Web.HttpUtility]::HtmlEncode($rtData.TestResult))</td></tr>" }
            $rtContent += "</table>"
            if($rtHasResults -and $rtData.Results){
                $rtRows=""
                foreach($r in @($rtData.Results)){
                    $rCls=if($r.Level -eq "Error"){"sc"}elseif($r.Level -eq "Warning"){"sw"}else{""}
                    $rtRows += "<tr><td>$($r.Time)</td><td class='$rCls'>$($r.Level)</td><td>$([System.Web.HttpUtility]::HtmlEncode($r.Message))</td></tr>"
                }
                $rtContent += "<table style='margin-top:12px'><tr><th>Time</th><th>Level</th><th>Message</th></tr>$rtRows</table>"
            }
            $rtOpenAttr = if($rtWHEA -gt 0){" open"}else{""}
            $ramTestSec = "<details class='sec'$rtOpenAttr><summary><h2>&#x1F9E0; RAM Test Logs (mdsched)</h2></summary>$rtContent</details>"
        }

        # Perfmon section
        $perfmonSec = ""
        $pmData = $Global:E.Perfmon
        if($pmData -and $pmData.Available -eq $true){
            $pmRows=""
            foreach($prop in @("CPUAvg","CPUPeak","RAMUsedGB","RAMFreeGB","DiskReadMBs","DiskWriteMBs","NetworkSentMBs","NetworkRecvMBs","PageFaultsPerSec","ContextSwitchesPerSec")){
                $val = $pmData.$prop
                if($null -ne $val){ $pmRows += "<tr><td>$prop</td><td>$val</td></tr>" }
            }
            if($pmData.TopCPUProcesses){ $pmRows += "<tr><td>Top CPU Processes</td><td style='font-size:12px'>$([System.Web.HttpUtility]::HtmlEncode($pmData.TopCPUProcesses))</td></tr>" }
            if($pmData.TopRAMProcesses){ $pmRows += "<tr><td>Top RAM Processes</td><td style='font-size:12px'>$([System.Web.HttpUtility]::HtmlEncode($pmData.TopRAMProcesses))</td></tr>" }
            $pmWarnings=""
            if($pmData.Warnings){ foreach($w in @($pmData.Warnings)){ $pmWarnings += "<div class='sw' style='margin-top:8px'>&#x26A0; $([System.Web.HttpUtility]::HtmlEncode($w))</div>" } }
            $pmOpenAttr = if($pmData.Warnings -and @($pmData.Warnings).Count -gt 0){" open"}else{""}
            $perfmonSec = "<details class='sec'$pmOpenAttr><summary><h2>&#x1F4C8; Performance Snapshot</h2></summary><table class='kv'>$pmRows</table>$pmWarnings</details>"
        } elseif($pmData -and $pmData.Status){
            $perfmonSec = "<details class='sec'><summary><h2>&#x1F4C8; Performance Report</h2></summary><p class='sw'>$([System.Web.HttpUtility]::HtmlEncode($pmData.Status))</p></details>"
        }

        $html = @"
<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>WinDiag-AI Report - $env:COMPUTERNAME</title>
<style>
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:'Segoe UI',sans-serif;background:#0f172a;color:#e2e8f0;line-height:1.6;padding:20px}
.ct{max-width:1200px;margin:0 auto}
.hd{background:linear-gradient(135deg,#1e293b,#0f172a);border:1px solid #334155;border-radius:12px;padding:30px;margin-bottom:24px;text-align:center}
.hd h1{color:#60a5fa;font-size:28px;margin-bottom:8px}.hd .sub{color:#94a3b8;font-size:14px}
.cds{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:16px;margin-bottom:24px}
.cd{background:#1e293b;border:1px solid #334155;border-radius:10px;padding:20px;text-align:center}
.cd .v{font-size:32px;font-weight:700}.cd .l{color:#94a3b8;font-size:13px;margin-top:4px}
.crit .v{color:#f87171}.warn .v{color:#fbbf24}.ok .v{color:#34d399}.info .v{color:#60a5fa}
details.sec{background:#1e293b;border:1px solid #334155;border-radius:10px;padding:24px;margin-bottom:20px}
details.sec > summary{cursor:pointer;list-style:none}
details.sec > summary::-webkit-details-marker{display:none}
details.sec > summary h2{color:#60a5fa;font-size:18px;padding-bottom:8px;border-bottom:1px solid #334155;position:relative;padding-right:30px}
details.sec > summary h2::after{content:'\25B6';position:absolute;right:0;font-size:12px;top:4px;transition:transform .2s}
details.sec[open] > summary h2::after{transform:rotate(90deg)}
table{width:100%;border-collapse:collapse;font-size:13px;margin-top:12px}
th{background:#334155;color:#e2e8f0;padding:10px 12px;text-align:left;font-weight:600}
td{padding:8px 12px;border-bottom:1px solid #1e293b}
tr:nth-child(even){background:rgba(51,65,85,.3)}
.kv td:first-child{color:#94a3b8;width:200px;font-weight:500}
.ai{background:linear-gradient(135deg,#1e1b4b,#1e293b);border:1px solid #4338ca;border-radius:10px;padding:24px;margin-bottom:20px}
.ai h2{color:#a78bfa;border-bottom-color:#4338ca}
.aic{font-size:14px;line-height:1.8}
.aic code{background:#334155;padding:2px 6px;border-radius:4px;font-family:Consolas,monospace;font-size:13px}
.ft{text-align:center;color:#64748b;font-size:12px;padding:20px 0}
.ft a{color:#60a5fa;text-decoration:none}.ft a:hover{text-decoration:underline}
.sok{color:#34d399}.sw{color:#fbbf24}.sc{color:#f87171;font-weight:700}
th.sortable{cursor:pointer;user-select:none;position:relative;padding-right:20px}
th.sortable::after{content:'\2195';position:absolute;right:4px;opacity:0.4;font-size:10px}
th.sortable.asc::after{content:'\25B2';opacity:0.8}
th.sortable.desc::after{content:'\25BC';opacity:0.8}
</style>
<script>
document.addEventListener('DOMContentLoaded',function(){
  document.querySelectorAll('table:not(.kv) th').forEach(function(th){th.classList.add('sortable')});
});
document.addEventListener('click',function(e){
  var th=e.target;if(th.tagName!=='TH'||!th.classList.contains('sortable'))return;
  var table=th.closest('table'),idx=Array.from(th.parentNode.children).indexOf(th);
  var body=table.querySelector('tbody')||table;
  var rows=Array.from(body.querySelectorAll('tr')).filter(function(r){return !r.querySelector('th')});
  var asc=!th.classList.contains('asc');
  th.parentNode.querySelectorAll('th').forEach(function(h){h.classList.remove('asc','desc')});
  th.classList.add(asc?'asc':'desc');
  rows.sort(function(a,b){
    var at=(a.children[idx]||{}).textContent||'',bt=(b.children[idx]||{}).textContent||'';
    var an=parseFloat(at.replace(/[,%]/g,'')),bn=parseFloat(bt.replace(/[,%]/g,''));
    if(!isNaN(an)&&!isNaN(bn))return asc?an-bn:bn-an;
    return asc?at.localeCompare(bt):bt.localeCompare(at);
  });
  rows.forEach(function(r){body.appendChild(r)});
});
</script>
</head>
<body><div class="ct">
<div class="hd"><h1>&#x1F5A5; WinDiag-AI Report</h1><div class="sub">$env:COMPUTERNAME &mdash; $(Get-Date -F 'dddd, dd MMMM yyyy - HH:mm:ss')</div></div>
<div class="cds">
<div class="cd $(if($cc-gt 0){'crit'}elseif($ec-gt 0){'warn'}else{'ok'})"><div class="v">$cc/$ec</div><div class="l">Critical/Errors</div></div>
<div class="cd $(if($wc-gt 10){'warn'}else{'info'})"><div class="v">$wc</div><div class="l">Warnings</div></div>
<div class="cd $(if($sc2-gt 0){'warn'}else{'ok'})"><div class="v">$sc2</div><div class="l">Service Issues</div></div>
<div class="cd info"><div class="v">$pu</div><div class="l">Updates</div></div>
</div>
<details class="sec" open><summary><h2>&#x1F310; Network Connectivity Test</h2></summary>
<div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:10px;margin-top:12px" id="netTestCards">
<div style="background:#0f172a;border:1px solid #334155;border-radius:8px;padding:14px;text-align:center"><div style="font-size:11px;color:#94a3b8;text-transform:uppercase;letter-spacing:0.1em">google.com</div><div id="nt-google" style="font-size:20px;font-weight:700;color:#fbbf24;margin-top:6px">Checking...</div></div>
<div style="background:#0f172a;border:1px solid #334155;border-radius:8px;padding:14px;text-align:center"><div style="font-size:11px;color:#94a3b8;text-transform:uppercase;letter-spacing:0.1em">debian.org</div><div id="nt-debian" style="font-size:20px;font-weight:700;color:#fbbf24;margin-top:6px">Checking...</div></div>
<div style="background:#0f172a;border:1px solid #334155;border-radius:8px;padding:14px;text-align:center"><div style="font-size:11px;color:#94a3b8;text-transform:uppercase;letter-spacing:0.1em">karanik.gr</div><div id="nt-karanik" style="font-size:20px;font-weight:700;color:#fbbf24;margin-top:6px">Checking...</div></div>
</div>
<div style="text-align:center;margin-top:12px"><button onclick="runNetTest()" style="background:#1e293b;border:1px solid #334155;color:#60a5fa;padding:6px 16px;border-radius:6px;cursor:pointer;font-size:12px">&#x25B6; Run Test</button></div>
<script>
async function testHost(url,id){var el=document.getElementById(id);el.textContent='Checking...';el.style.color='#fbbf24';var t=Date.now();try{await fetch(url,{method:'HEAD',mode:'no-cors',cache:'no-store',signal:AbortSignal.timeout(6000)});var ms=Date.now()-t;el.textContent=ms+'ms';el.style.color='#34d399'}catch(e){el.textContent=e.name==='TimeoutError'?'Timeout':'Failed';el.style.color='#f87171'}}
function runNetTest(){testHost('https://www.google.com','nt-google');testHost('https://www.debian.org','nt-debian');testHost('https://karanik.gr','nt-karanik')}
runNetTest();
</script>
</details>
<details class="sec" open><summary><h2>&#x2139; System Information</h2></summary><table class="kv">$sysInfoHtml</table></details>
$hwSec
$battSec
<details class="sec" open><summary><h2>&#x1F4BE; Disks</h2></summary><table><tr><th>Drive</th><th>Label</th><th>Total</th><th>Free</th><th>Free%</th><th>Status</th></tr>$diskRows</table></details>
$smartSec
<details class="sec" open><summary><h2>&#x1F4CB; Events ($($Global:EventHours)h)</h2></summary><table><tr><th>Time</th><th>Log</th><th>Level</th><th>Source</th><th>ID</th><th>Message</th></tr>$eventRows</table></details>
<details class="sec" open><summary><h2>&#x2699; Services (Stopped Auto-Start: $($servicesArr.Count))</h2></summary>
$(if($servicesArr.Count -gt 0){"<table><tr><th>Service</th><th>Name</th><th>State</th><th>RunAs</th></tr>$svcRows</table>"}else{"<p class='sok'>All auto-start services running</p>"})
$(
    $allSvc = AsArray $Global:E.AllServices
    if($allSvc.Count -gt 0){
        $allSvcRows=""
        foreach($s in $allSvc){
            $stCls=if($s.State -eq "Running"){"sok"}elseif($s.State -eq "Stopped"){"sw"}else{""}
            $allSvcRows+="<tr><td>$([System.Web.HttpUtility]::HtmlEncode($s.DisplayName))</td><td>$($s.Name)</td><td class='$stCls'>$($s.State)</td><td>$($s.StartMode)</td><td>$($s.RunAs)</td></tr>"
        }
        "<details style='margin-top:16px'><summary style='color:#94a3b8;font-size:13px;cursor:pointer'>All Services ($($allSvc.Count))</summary><table><tr><th>Service</th><th>Name</th><th>State</th><th>StartMode</th><th>RunAs</th></tr>$allSvcRows</table></details>"
    }
)
</details>
$driverSec
<details class="sec" open><summary><h2>&#x1F527; System Integrity</h2></summary><table class="kv">$integrityHtml</table><p style="color:#94a3b8;font-size:12px;margin-top:8px"><code>DISM /Online /Cleanup-Image /RestoreHealth</code> then <code>sfc /scannow</code></p></details>
<details class="sec" open><summary><h2>&#x1F50D; Search/Indexing</h2></summary><table class="kv">$searchHtml</table></details>
$(
    $ua = AsArray $Global:E.UserInfo.LocalAccounts
    if($ua.Count -gt 0){
        $ur=""; foreach($u in $ua){$ec2=if($u.Enabled -eq $true -or $u.Enabled -eq "True"){"sok"}else{"sw"};$ur+="<tr><td>$([System.Web.HttpUtility]::HtmlEncode($u.Name))</td><td class='$ec2'>$($u.Enabled)</td><td>$($u.PasswordRequired)</td><td>$($u.LastLogon)</td></tr>"}
        $uf=AsArray $Global:E.UserInfo.UserFolders; $ufp=if($uf.Count -gt 0){"<p style='color:#94a3b8;font-size:12px;margin-top:10px'>User folders: $($uf -join ', ')</p>"}else{""}
        $as=AsArray $Global:E.UserInfo.ActiveSessions; $asp=if($as.Count -gt 0){"<details style='margin-top:12px'><summary style='color:#94a3b8;font-size:12px;cursor:pointer'>Active Sessions</summary><pre style='background:#0f172a;padding:12px;border-radius:6px;font-size:11px;margin-top:8px;white-space:pre-wrap'>$($as -join "`n")</pre></details>"}else{""}
        "<details class='sec' open><summary><h2>&#x1F464; User Accounts ($($ua.Count))</h2></summary><table><tr><th>Name</th><th>Enabled</th><th>Password Required</th><th>Last Logon</th></tr>$ur</table>$ufp$asp</details>"
    }
)
<details class="sec" open><summary><h2>&#x1F310; Network</h2></summary><table><tr><th>Adapter</th><th>IP</th><th>Gateway</th><th>DNS</th><th>Speed</th><th>MAC</th></tr>$netRows</table><p style="color:#94a3b8;font-size:12px;margin-top:10px">DNS: $dnsText</p></details>
$(if($Global:E.ExternalIP.IP){"<details class='sec' open><summary><h2>&#x1F30D; External IP</h2></summary><table class='kv'><tr><td>External IP</td><td>$([System.Web.HttpUtility]::HtmlEncode($Global:E.ExternalIP.IP))</td></tr></table></details>"})
$(
    $nc = AsArray $Global:E.NetConnections
    if($nc.Count -gt 0){
        $ncr=""; foreach($c in $nc){$ncr+="<tr><td>$($c.LocalAddr)</td><td>$($c.RemoteAddr)</td><td>$($c.Process)</td><td>$($c.PID)</td></tr>"}
        "<details class='sec'><summary><h2>&#x1F4E1; Established Connections ($($nc.Count))</h2></summary><table><tr><th>Local</th><th>Remote</th><th>Process</th><th>PID</th></tr>$ncr</table></details>"
    }
)
$(
    $nr = AsArray $Global:E.NetRoutes
    if($nr.Count -gt 0){
        $nrr=""; foreach($r in $nr){$nrr+="<tr><td>$($r.Destination)</td><td>$($r.NextHop)</td><td>$($r.Metric)</td><td>$($r.Interface)</td></tr>"}
        "<details class='sec'><summary><h2>&#x1F5FA; Routing Table ($($nr.Count))</h2></summary><table><tr><th>Destination</th><th>Next Hop</th><th>Metric</th><th>Interface</th></tr>$nrr</table></details>"
    }
)
$(
    $at = AsArray $Global:E.ArpTable
    if($at.Count -gt 0){
        $atr=""; foreach($a in $at){$atr+="<tr><td>$($a.IP)</td><td>$($a.MAC)</td><td>$($a.State)</td><td>$($a.Interface)</td></tr>"}
        "<details class='sec'><summary><h2>&#x1F4CB; ARP Table ($($at.Count))</h2></summary><table><tr><th>IP</th><th>MAC</th><th>State</th><th>Interface</th></tr>$atr</table></details>"
    }
)
$(
    $md = AsArray $Global:E.MappedDrives
    if($md.Count -gt 0){
        $mdr=""; foreach($m in $md){$mdr+="<tr><td>$($m.Drive)</td><td>$($m.Path)</td><td>$($m.Free)</td><td>$($m.Used)</td></tr>"}
        "<details class='sec'><summary><h2>&#x1F4C1; Mapped Drives ($($md.Count))</h2></summary><table><tr><th>Drive</th><th>Path</th><th>Free</th><th>Used</th></tr>$mdr</table></details>"
    }
)
<details class="sec" open><summary><h2>&#x1F6E1; Security</h2></summary><table class="kv">$securityHtml</table></details>
$updSec
$(
    $hfa = AsArray $Global:E.Hotfixes
    if($hfa.Count -gt 0){
        $hfr=""; foreach($h in $hfa){$hfr+="<tr><td>$($h.HotFixID)</td><td>$($h.InstalledOn)</td><td>$($h.Description)</td></tr>"}
        "<details class='sec' open><summary><h2>&#x1F4CB; Installed Hotfixes ($($hfa.Count))</h2></summary><table><tr><th>KB</th><th>Installed</th><th>Description</th></tr>$hfr</table></details>"
    }
)
<details class="sec"><summary><h2>&#x1F4CA; Top Processes by RAM</h2></summary><table><tr><th>Name</th><th>PID</th><th>RAM (MB)</th><th>CPU (s)</th></tr>$ramRows</table></details>
<details class="sec"><summary><h2>&#x1F4CA; Top Processes by CPU</h2></summary><table><tr><th>Name</th><th>PID</th><th>CPU (s)</th><th>RAM (MB)</th></tr>$cpuRows</table></details>
<details class="sec"><summary><h2>&#x1F680; Startup Programs ($($startupArr.Count))</h2></summary><table><tr><th>Name</th><th>Command</th><th>Location</th></tr>$startupRows</table></details>
$taskSec
$autoSec
$swSec
$(
    $ha = AsArray $Global:E.Hosts.Entries; $htc = if($Global:E.Hosts.TotalCount){$Global:E.Hosts.TotalCount}else{$ha.Count}
    if($htc -gt 0){
        $hr=""; foreach($h in $ha){$hr+="<tr><td>$([System.Web.HttpUtility]::HtmlEncode($h.IP))</td><td>$([System.Web.HttpUtility]::HtmlEncode($h.Hostname))</td></tr>"}
        $tn=if($Global:E.Hosts.Truncated){"<p style='color:#fbbf24;font-size:12px;margin-top:8px'>Showing first $($ha.Count) of $htc entries</p>"}else{""}
        "<details class='sec'><summary><h2>&#x1F4C3; Hosts File ($htc entries)</h2></summary><table><tr><th>IP</th><th>Hostname</th></tr>$hr</table>$tn</details>"
    }
)
$(
    $cla = AsArray $Global:E.ChkdskLogs
    if($cla.Count -gt 0){
        $clr=""; foreach($c in $cla){$clr+="<tr><td>$($c.Time)</td><td style='font-size:12px'>$([System.Web.HttpUtility]::HtmlEncode($c.Message))</td></tr>"}
        "<details class='sec'><summary><h2>&#x1F4BE; Chkdsk Logs ($($cla.Count))</h2></summary><table><tr><th>Time</th><th>Message</th></tr>$clr</table></details>"
    }
)
$(
    $rtk = KvRows $Global:E.RemoteTools
    if($rtk){"<details class='sec'><summary><h2>&#x1F4E1; Remote Tools</h2></summary><table class='kv'>$rtk</table></details>"}
)
$(
    if($Global:E.BatteryReport.Available -eq $true -or $Global:E.BatteryReport.Available -eq "True"){
        $brk = KvRows $Global:E.BatteryReport
        "<details class='sec'><summary><h2>&#x1F50B; Battery Report (powercfg)</h2></summary><table class='kv'>$brk</table></details>"
    }
)
$bsodSec
$ramTestSec
$perfmonSec
$customSec
<div class="ai"><h2>&#x1F916; AI Analysis</h2><div class="aic">$ah</div></div>
<div class="ft">WinDiag-AI &bull; $(Get-Date -F 'yyyy-MM-dd HH:mm:ss')<br><a href="https://karanik.gr" target="_blank" style="color:#60a5fa;text-decoration:none">karanik.gr</a></div>
</div></body></html>
"@
        [IO.File]::WriteAllText($rp,$html,[Text.Encoding]::UTF8)
        Ui-Log "Report: $rp" "OK"; Ui-SetStatus "Report saved."; Start-Process $rp
    } catch { Ui-Log "Report error: $($_.Exception.Message)" "ERROR" }
})

# ── Download button (toolbar) ──
$btnPull.Add_Click({
    # If download is in progress, cancel it
    if($btnPull.Tag -eq "downloading"){
        $Global:CancelDownload.TryAdd("cancel", $true) | Out-Null
        $Global:CancelDownload["cancel"] = $true
        Ui-Log "Cancelling download..." "WARN"
        Ui-SetStatus "Cancelling..."
        $btnPull.Content = [char]0x2B07 + " Download"
        $btnPull.Tag = $null
        return
    }

    $m=Get-SelModel
    if($m-eq"Custom"-or($cmbModel.SelectedItem-and $cmbModel.SelectedItem-like"*Custom*")){
        try{Add-Type -AssemblyName Microsoft.VisualBasic -EA SilentlyContinue}catch{}
        try{$m=[Microsoft.VisualBasic.Interaction]::InputBox("Model name (e.g. llama3.2, mistral, phi3:mini):","Download Model","")}catch{$m=""}
    }
    if(-not $m){return}; if(-not(Test-OllamaRunning)){[Windows.MessageBox]::Show("Ollama not running.","Error","OK","Warning");return}

    # Reset cancel flag and start download
    $Global:CancelDownload.TryAdd("cancel", $false) | Out-Null
    $Global:CancelDownload["cancel"] = $false

    Set-UiBusy $true
    # Re-enable pull button as Cancel
    $btnPull.IsEnabled = $true
    $btnPull.Content = [char]0x274C + " Cancel"
    $btnPull.Tag = "downloading"

    Ui-Log "Downloading '$m' from $($Global:OllamaUrl)/api/pull" "INFO";Ui-SetStatus "Downloading $m..."
    Start-BackgroundJob -Job $Global:StreamingPullScript -Params @{ModelName=$m}
})

$btnCleanup.Add_Click({
    $ms=Get-OllamaModels;$su=Get-OllamaStorageSize
    if($ms.Count-eq 0){return}
    $list=($ms|ForEach-Object{"  - $_"})-join"`n"
    $a=[Windows.MessageBox]::Show("Models ($su GB):`n$list`n`n[Yes] Delete ALL`n[No] Delete '$($Script:SelectedModel)'`n[Cancel] Keep","Cleanup","YesNoCancel","Question")
    if($a-eq"Yes"){foreach($m in $ms){Start-Process "ollama" -ArgumentList "rm","$($m-replace':latest$','')" -NoNewWindow -Wait -EA SilentlyContinue};Ui-Log "All removed" "OK"}
    elseif($a-eq"No"-and $Script:SelectedModel){Start-Process "ollama" -ArgumentList "rm","$($Script:SelectedModel)" -NoNewWindow -Wait -EA SilentlyContinue;Ui-Log "Removed" "OK"}
    Refresh-Ollama
})

$btnClearLog.Add_Click({$txtLog.Document.Blocks.Clear();$txtAI.Document.Blocks.Clear()})

$mnuExit.Add_Click({ $W.Close() })

$mnuAbout.Add_Click({
    $aw = [Windows.Window]::new()
    $aw.Title = "About $AppName"; $aw.Width = 340; $aw.Height = 220
    $aw.WindowStartupLocation = "CenterOwner"; $aw.Owner = $W
    $aw.ResizeMode = "NoResize"; $aw.Background = [Windows.Media.Brushes]::White
    $sp = [Windows.Controls.StackPanel]::new(); $sp.Margin = [Windows.Thickness]::new(20)
    $t1 = [Windows.Controls.TextBlock]::new(); $t1.Text = "$AppName v$AppVer"; $t1.FontSize = 16; $t1.FontWeight = "SemiBold"; $t1.Margin = [Windows.Thickness]::new(0,0,0,6); $sp.Children.Add($t1)
    $t2 = [Windows.Controls.TextBlock]::new(); $t2.Text = "Windows Diagnostics with AI Analysis"; $t2.Foreground = [Windows.Media.Brushes]::Gray; $t2.Margin = [Windows.Thickness]::new(0,0,0,12); $sp.Children.Add($t2)
    $t3 = [Windows.Controls.TextBlock]::new(); $t3.Text = "by Nikolaos Karanikolas"; $t3.Margin = [Windows.Thickness]::new(0,0,0,4); $sp.Children.Add($t3)
    $hl = [Windows.Documents.Hyperlink]::new([Windows.Documents.Run]::new("karanik.gr"))
    $hl.NavigateUri = [Uri]::new("https://karanik.gr")
    $hl.Add_RequestNavigate({ param($s,$e) Start-Process $e.Uri.AbsoluteUri; $e.Handled = $true })
    $t4 = [Windows.Controls.TextBlock]::new(); $t4.Inlines.Add($hl); $sp.Children.Add($t4)
    $t5 = [Windows.Controls.TextBlock]::new(); $t5.Text = "`nScript folder: $($Global:ScriptDir)"; $t5.FontSize = 10; $t5.Foreground = [Windows.Media.Brushes]::Gray; $t5.TextWrapping = "Wrap"; $sp.Children.Add($t5)
    $aw.Content = $sp; Apply-DialogTheme $aw; $aw.ShowDialog() | Out-Null
})

# ── Export to PDF ──
$mnuExportPdf.Add_Click({
    if(-not $Global:DiagData){return}
    # Generate HTML report first, then open print dialog
    Ui-Log "Exporting PDF: Generate HTML and use browser Print > Save as PDF" "INFO"
    $btnReport.RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Primitives.ButtonBase]::ClickEvent))
    [Windows.MessageBox]::Show("To save as PDF:`n`n1. The HTML report opened in your browser`n2. Press Ctrl+P (Print)`n3. Select 'Save as PDF' as printer`n4. Click Save`n`nThe HTML report includes all diagnostic data.","Export to PDF","OK","Information")
})


# ── Custom Checks ──
$mnuCustomChecks.Add_Click({
    if(-not (Test-Path $Global:CustomChecksDir)){
        $a = [Windows.MessageBox]::Show("Custom checks folder not found:`n$($Global:CustomChecksDir)`n`nCreate it now? You can place .ps1 scripts there and they will run during scan.","Custom Checks","YesNo","Question")
        if($a -eq "Yes"){
            New-Item -ItemType Directory -Path $Global:CustomChecksDir -Force | Out-Null
            # Create example script
            $example = @'
# Example custom check - WinDiag-AI
# Place .ps1 files in this folder - they run automatically during scan
# Output is captured and included in reports

Write-Output "Custom check ran successfully at $(Get-Date)"
Write-Output "Hostname: $env:COMPUTERNAME"
'@
            [IO.File]::WriteAllText((Join-Path $Global:CustomChecksDir "_example.ps1"), $example, [Text.Encoding]::UTF8)
            Ui-Log "Created custom checks folder: $($Global:CustomChecksDir)" "OK"
        }
    }
    if(Test-Path $Global:CustomChecksDir){
        Start-Process "explorer.exe" -ArgumentList $Global:CustomChecksDir
    }
})

# ── Settings ──
$mnuSettings.Add_Click({
    $sw = [Windows.Window]::new()
    $sw.Title = "Settings"; $sw.Width = 800; $sw.Height = 650
    $sw.WindowStartupLocation = "CenterOwner"; $sw.Owner = $W
    $sw.ResizeMode = "CanResize"; $sw.Background = [Windows.Media.Brushes]::White

    $tabs = [Windows.Controls.TabControl]::new(); $tabs.Margin = [Windows.Thickness]::new(10)

    # ── Tab 1: Models (with info column) ──
    $tabModels = [Windows.Controls.TabItem]::new(); $tabModels.Header = "AI Models"
    $msp = [Windows.Controls.StackPanel]::new(); $msp.Margin = [Windows.Thickness]::new(10)
    $mlbl = [Windows.Controls.TextBlock]::new(); $mlbl.Text = "Installed & Available Models"; $mlbl.FontWeight = "SemiBold"; $mlbl.FontSize = 14; $mlbl.Margin = [Windows.Thickness]::new(0,0,0,8); $msp.Children.Add($mlbl)

    $mList = [Windows.Controls.ListView]::new(); $mList.Height = 320
    $gv = [Windows.Controls.GridView]::new()
    foreach($col in @(@{H="Model";W=120},@{H="Size";W=70},@{H="Quality";W=90},@{H="Status";W=80},@{H="MinRAM";W=55},@{H="Speed";W=120},@{H="URL";W=160})){
        $c = [Windows.Controls.GridViewColumn]::new(); $c.Header = $col.H; $c.Width = $col.W; $c.DisplayMemberBinding = [Windows.Data.Binding]::new($col.H); $gv.Columns.Add($c)
    }
    $mList.View = $gv
    $installed = Get-OllamaModels
    foreach($rm in $RecommendedModels){
        $isInst = $installed | Where-Object { $_ -like "$($rm.Name)*" }
        $status = if($isInst){"Installed"}else{"Not installed"}
        # Parse speed and URL from Info
        $speed = ""; $url = ""
        if($rm.Info -match 'Speed:\s*(.+)'){$speed = $matches[1].Trim()}
        if($rm.Info -match 'URL:\s*(.+)'){$url = $matches[1].Trim()}
        $mList.Items.Add([PSCustomObject]@{Model=$rm.Name;Size=$rm.Size;Quality=$rm.Quality;Status=$status;MinRAM="$($rm.MinRAM) GB";Speed=$speed;URL=$url}) | Out-Null
    }
    $msp.Children.Add($mList)

    # Download status label
    $lblDlStatus = [Windows.Controls.TextBlock]::new(); $lblDlStatus.Text = ""; $lblDlStatus.Foreground = [Windows.Media.Brushes]::Gray; $lblDlStatus.Margin = [Windows.Thickness]::new(0,4,0,4); $lblDlStatus.FontSize = 11
    $msp.Children.Add($lblDlStatus)

    $btnPanel = [Windows.Controls.StackPanel]::new(); $btnPanel.Orientation = "Horizontal"; $btnPanel.Margin = [Windows.Thickness]::new(0,4,0,0)
    $btnDL = [Windows.Controls.Button]::new(); $btnDL.Content = "Download Selected"; $btnDL.Padding = [Windows.Thickness]::new(12,4,12,4); $btnDL.Margin = [Windows.Thickness]::new(0,0,8,0)
    $btnDel = [Windows.Controls.Button]::new(); $btnDel.Content = "Delete Selected"; $btnDel.Padding = [Windows.Thickness]::new(12,4,12,4); $btnDel.Margin = [Windows.Thickness]::new(0,0,8,0)
    $btnOpenF = [Windows.Controls.Button]::new(); $btnOpenF.Content = "Open Models Folder"; $btnOpenF.Padding = [Windows.Thickness]::new(12,4,12,4)

    $btnDL.Add_Click({
        $sel = $mList.SelectedItem; if(-not $sel){[Windows.MessageBox]::Show("Select a model first.","Info","OK","Information");return}
        $mn = $sel.Model; if(-not(Test-OllamaRunning)){[Windows.MessageBox]::Show("Ollama not running.","Error","OK","Warning");return}
        $sw.Close()
        # Reset cancel flag
        $Global:CancelDownload.TryAdd("cancel", $false) | Out-Null
        $Global:CancelDownload["cancel"] = $false
        Set-UiBusy $true
        # Show cancel button
        $btnPull.IsEnabled = $true
        $btnPull.Content = [char]0x274C + " Cancel"
        $btnPull.Tag = "downloading"
        Ui-Log "Downloading '$mn' from $($Global:OllamaUrl)/api/pull" "INFO"; Ui-SetStatus "Downloading $mn..."
        Start-BackgroundJob -Job $Global:StreamingPullScript -Params @{ModelName=$mn}
    })
    $btnDel.Add_Click({
        $sel = $mList.SelectedItem; if(-not $sel){return}
        $mn = $sel.Model; if(-not(Test-OllamaRunning)){return}
        $a=[Windows.MessageBox]::Show("Delete model '$mn'?","Confirm","YesNo","Question")
        if($a -eq "Yes"){
            try{ Invoke-RestMethod -Uri "$($Global:OllamaUrl)/api/delete" -Method Delete -Body (@{name=$mn}|ConvertTo-Json) -ContentType "application/json" -EA Stop; Ui-Log "'$mn' deleted" "OK" }catch{ Ui-Log "Delete failed: $($_.Exception.Message)" "ERROR" }
            $sw.Close(); Refresh-Ollama
        }
    })
    $btnOpenF.Add_Click({ $p = Get-OllamaStoragePath; if(Test-Path $p){Start-Process "explorer.exe" -ArgumentList $p} })
    $btnPanel.Children.Add($btnDL); $btnPanel.Children.Add($btnDel); $btnPanel.Children.Add($btnOpenF)
    $msp.Children.Add($btnPanel)
    $tabModels.Content = $msp; $tabs.Items.Add($tabModels)

    # ── Tab 2: General Settings (with tooltips) ──
    $tabGeneral = [Windows.Controls.TabItem]::new(); $tabGeneral.Header = "General"
    $gScroll = [Windows.Controls.ScrollViewer]::new(); $gScroll.VerticalScrollBarVisibility = "Auto"
    $gsp = [Windows.Controls.StackPanel]::new(); $gsp.Margin = [Windows.Thickness]::new(10)

    # Inline helper: create a label with tooltip and add field to panel
    # We do this inline to avoid PS 5.1 issues with nested functions in event handlers

    # -- Ollama URL --
    $w1 = [Windows.Controls.StackPanel]::new(); $w1.Margin = [Windows.Thickness]::new(0,0,0,12)
    $l1 = [Windows.Controls.TextBlock]::new(); $l1.Text = "Ollama URL (?)"; $l1.FontWeight = "SemiBold"; $l1.Margin = [Windows.Thickness]::new(0,0,0,4); $l1.Cursor = [Windows.Input.Cursors]::Help
    $l1.ToolTip = "The Ollama server address. Default: http://localhost:11434`nFor remote: http://192.168.1.100:11434"
    $txtUrl = [Windows.Controls.TextBox]::new(); $txtUrl.Text = $Global:OllamaUrl
    $w1.Children.Add($l1); $w1.Children.Add($txtUrl); $gsp.Children.Add($w1)

    # -- Report Output Folder --
    $w2 = [Windows.Controls.StackPanel]::new(); $w2.Margin = [Windows.Thickness]::new(0,0,0,12)
    $l2 = [Windows.Controls.TextBlock]::new(); $l2.Text = "Report Output Folder (?)"; $l2.FontWeight = "SemiBold"; $l2.Margin = [Windows.Thickness]::new(0,0,0,4); $l2.Cursor = [Windows.Input.Cursors]::Help
    $l2.ToolTip = "Where HTML reports are saved.`nDefault: script folder.`nChange to any valid path."
    $txtRptPath = [Windows.Controls.TextBox]::new(); $txtRptPath.Text = if($Global:ReportPath){$Global:ReportPath}else{$Global:ScriptDir}
    $w2.Children.Add($l2); $w2.Children.Add($txtRptPath); $gsp.Children.Add($w2)

    # -- Log File Path --
    $w3 = [Windows.Controls.StackPanel]::new(); $w3.Margin = [Windows.Thickness]::new(0,0,0,12)
    $l3 = [Windows.Controls.TextBlock]::new(); $l3.Text = "Log File Path (?)"; $l3.FontWeight = "SemiBold"; $l3.Margin = [Windows.Thickness]::new(0,0,0,4); $l3.Cursor = [Windows.Input.Cursors]::Help
    $l3.ToolTip = "Log file for all actions. Default: WinDiag-AI.log in script folder.`nLeave empty to disable file logging."
    $txtLogPath = [Windows.Controls.TextBox]::new(); $txtLogPath.Text = if($Global:LogFilePath){$Global:LogFilePath}else{Join-Path $Global:ScriptDir "WinDiag-AI.log"}
    $w3.Children.Add($l3); $w3.Children.Add($txtLogPath); $gsp.Children.Add($w3)

    # -- AI Temperature --
    $w4 = [Windows.Controls.StackPanel]::new(); $w4.Margin = [Windows.Thickness]::new(0,0,0,12)
    $l4 = [Windows.Controls.TextBlock]::new(); $l4.Text = "AI Temperature (?)"; $l4.FontWeight = "SemiBold"; $l4.Margin = [Windows.Thickness]::new(0,0,0,4); $l4.Cursor = [Windows.Input.Cursors]::Help
    $l4.ToolTip = "How creative the AI is.`n0.1 = Precise, technical`n0.3 = Recommended for diagnostics`n0.5 = Balanced`n0.9 = More creative, free"
    $sldTemp = [Windows.Controls.Slider]::new(); $sldTemp.Minimum = 0.1; $sldTemp.Maximum = 0.9; $sldTemp.Value = if($Global:AiTemp){$Global:AiTemp}else{0.3}; $sldTemp.TickFrequency = 0.1; $sldTemp.IsSnapToTickEnabled = $true
    $lblTemp = [Windows.Controls.TextBlock]::new(); $lblTemp.Text = "Current: $($sldTemp.Value)"; $lblTemp.Foreground = [Windows.Media.Brushes]::Gray; $lblTemp.FontSize = 11
    $sldTemp.Add_ValueChanged({ $lblTemp.Text = "Current: $([math]::Round($sldTemp.Value,1))" })
    $w4.Children.Add($l4); $w4.Children.Add($sldTemp); $w4.Children.Add($lblTemp); $gsp.Children.Add($w4)

    # -- AI Max Tokens --
    $w5 = [Windows.Controls.StackPanel]::new(); $w5.Margin = [Windows.Thickness]::new(0,0,0,12)
    $l5 = [Windows.Controls.TextBlock]::new(); $l5.Text = "AI Max Tokens (?)"; $l5.FontWeight = "SemiBold"; $l5.Margin = [Windows.Thickness]::new(0,0,0,4); $l5.Cursor = [Windows.Input.Cursors]::Help
    $l5.ToolTip = "Max length of AI response in tokens.`n1024 = Short (~500 words)`n4096 = Recommended (~2000 words)`n8192 = Very detailed (~4000 words)`nHigher = slower response."
    $sldTok = [Windows.Controls.Slider]::new(); $sldTok.Minimum = 1024; $sldTok.Maximum = 8192; $sldTok.Value = if($Global:AiMaxTokens){$Global:AiMaxTokens}else{4096}; $sldTok.TickFrequency = 512; $sldTok.IsSnapToTickEnabled = $true
    $lblTok = [Windows.Controls.TextBlock]::new(); $lblTok.Text = "Current: $([int]$sldTok.Value)"; $lblTok.Foreground = [Windows.Media.Brushes]::Gray; $lblTok.FontSize = 11
    $sldTok.Add_ValueChanged({ $lblTok.Text = "Current: $([int]$sldTok.Value)" })
    $w5.Children.Add($l5); $w5.Children.Add($sldTok); $w5.Children.Add($lblTok); $gsp.Children.Add($w5)

    # -- Events Timeframe --
    $w6 = [Windows.Controls.StackPanel]::new(); $w6.Margin = [Windows.Thickness]::new(0,0,0,12)
    $l6 = [Windows.Controls.TextBlock]::new(); $l6.Text = "Events Timeframe (?)"; $l6.FontWeight = "SemiBold"; $l6.Margin = [Windows.Thickness]::new(0,0,0,4); $l6.Cursor = [Windows.Input.Cursors]::Help
    $l6.ToolTip = "How far back to look in Event Logs.`n12h = Recent only`n24h = Recommended`n48h = Last 2 days`n7d = Full week (slower scan)"
    $cmbEvt = [Windows.Controls.ComboBox]::new()
    foreach($opt in @("12 hours","24 hours","48 hours","7 days")){$cmbEvt.Items.Add($opt)|Out-Null}
    $cmbEvt.SelectedIndex = if($Global:EventHours -eq 12){0}elseif($Global:EventHours -eq 48){2}elseif($Global:EventHours -eq 168){3}else{1}
    $w6.Children.Add($l6); $w6.Children.Add($cmbEvt); $gsp.Children.Add($w6)

    # -- Auto-Save Report --
    $w7 = [Windows.Controls.StackPanel]::new(); $w7.Margin = [Windows.Thickness]::new(0,0,0,12)
    $l7 = [Windows.Controls.TextBlock]::new(); $l7.Text = "Auto-Save Report (?)"; $l7.FontWeight = "SemiBold"; $l7.Margin = [Windows.Thickness]::new(0,0,0,4); $l7.Cursor = [Windows.Input.Cursors]::Help
    $l7.ToolTip = "If enabled, automatically saves HTML report after each scan.`nSaved to Report Output Folder."
    $chkAutoSave = [Windows.Controls.CheckBox]::new(); $chkAutoSave.Content = "Auto-save report after scan"; $chkAutoSave.IsChecked = $Global:AutoSaveReport
    $w7.Children.Add($l7); $w7.Children.Add($chkAutoSave); $gsp.Children.Add($w7)

    # -- Custom Checks Folder --
    $w8 = [Windows.Controls.StackPanel]::new(); $w8.Margin = [Windows.Thickness]::new(0,0,0,12)
    $l8 = [Windows.Controls.TextBlock]::new(); $l8.Text = "Custom Checks Folder (?)"; $l8.FontWeight = "SemiBold"; $l8.Margin = [Windows.Thickness]::new(0,0,0,4); $l8.Cursor = [Windows.Input.Cursors]::Help
    $l8.ToolTip = "Folder with .ps1 scripts that run during scan.`nPlace your own diagnostic scripts here.`nOutput is captured and included in reports."
    $txtCustomDir = [Windows.Controls.TextBox]::new(); $txtCustomDir.Text = $Global:CustomChecksDir
    $w8.Children.Add($l8); $w8.Children.Add($txtCustomDir); $gsp.Children.Add($w8)


    $btnSave = [Windows.Controls.Button]::new(); $btnSave.Content = "Save Settings"; $btnSave.FontWeight = "SemiBold"; $btnSave.Padding = [Windows.Thickness]::new(20,6,20,6); $btnSave.HorizontalAlignment = "Left"; $btnSave.Margin = [Windows.Thickness]::new(0,8,0,0)
    $btnSave.Add_Click({
        $Global:OllamaUrl = $txtUrl.Text.TrimEnd('/')
        $Global:ReportPath = $txtRptPath.Text
        $Global:LogFilePath = $txtLogPath.Text
        $Global:AiTemp = [math]::Round($sldTemp.Value,1)
        $Global:AiMaxTokens = [int]$sldTok.Value
        $evtMap = @{0=12;1=24;2=48;3=168}; $Global:EventHours = $evtMap[$cmbEvt.SelectedIndex]
        $Global:AutoSaveReport = [bool]$chkAutoSave.IsChecked
        $Global:CustomChecksDir = $txtCustomDir.Text
        Ui-Log "Settings saved: URL=$($Global:OllamaUrl) Temp=$($Global:AiTemp) MaxTok=$($Global:AiMaxTokens) Events=$($Global:EventHours)h AutoSave=$($Global:AutoSaveReport) Report=$($Global:ReportPath)" "OK"
        $sw.Close(); Refresh-Ollama
    })
    $gsp.Children.Add($btnSave)
    $gScroll.Content = $gsp
    $tabGeneral.Content = $gScroll; $tabs.Items.Add($tabGeneral)

    $sw.Content = $tabs; Apply-DialogTheme $sw; $sw.ShowDialog() | Out-Null
})

$btnOllama.Add_Click({
    $inst=Test-OllamaInstalled;$run=Test-OllamaRunning
    if(-not $inst){
        $a=[Windows.MessageBox]::Show("Ollama NOT installed.`n`nFree, local AI. Models at:`n$env:USERPROFILE\.ollama\models`n`n[Yes] winget install`n[No] Open browser","Setup","YesNoCancel","Information")
        if($a-eq"Yes"){Ui-Log "Installing..." "INFO";Ui-SetStatus "Installing..."
            try{$p=Start-Process "winget" -ArgumentList "install","Ollama.Ollama","--accept-package-agreements","--accept-source-agreements" -NoNewWindow -Wait -PassThru -EA Stop
                if($p.ExitCode-eq 0){$env:Path=[Environment]::GetEnvironmentVariable("Path","Machine")+";"+[Environment]::GetEnvironmentVariable("Path","User");Ui-Log "Installed!" "OK";Start-Process "ollama" -ArgumentList "serve" -WindowStyle Hidden -EA SilentlyContinue;Start-Sleep 5}
            }catch{Ui-Log "Error: $($_.Exception.Message)" "ERROR"}}
        elseif($a-eq"No"){Start-Process "https://ollama.com/download/windows"}
    }elseif(-not $run){
        $a=[Windows.MessageBox]::Show("Start Ollama?","Start?","YesNo","Question")
        if($a-eq"Yes"){Start-Process "ollama" -ArgumentList "serve" -WindowStyle Hidden -EA SilentlyContinue;Start-Sleep 5;if(Test-OllamaRunning){Ui-Log "Started!" "OK"}else{Ui-Log "Not responding" "WARN"}}
    }else{[Windows.MessageBox]::Show("Ollama running!`nURL: $($Global:OllamaUrl)`nStorage: $(Get-OllamaStoragePath)`nUsed: $(Get-OllamaStorageSize) GB","Status","OK","Information")}
    Refresh-Ollama
})
#endregion

#region Startup
$Host.UI.RawUI.WindowTitle = "$AppName - Console Output"
Write-Host ""
Write-Host "  =============================================" -ForegroundColor Cyan
Write-Host "    $AppName - Console Log" -ForegroundColor White
Write-Host "    by Nikolaos Karanikolas | karanik.gr" -ForegroundColor Gray
Write-Host "  =============================================" -ForegroundColor Cyan
Write-Host ""
Ui-Log "Started $AppName" "INFO"
Ui-Log "Computer: $env:COMPUTERNAME | User: $env:USERNAME" "INFO"
Ui-Log "Script folder: $($Global:ScriptDir)" "INFO"
Ui-Log "Report path: $($Global:ReportPath)" "INFO"
Ui-Log "Log file: $($Global:LogFilePath)" "INFO"
if(-not([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)){Ui-Log "Not Administrator - some checks limited" "WARN"}
# Create custom checks folder info
if(-not (Test-Path $Global:CustomChecksDir)){ Ui-Log "Custom checks: folder not found (File > Custom Checks to create)" "INFO" }
else{ $ccCount = @(Get-ChildItem $Global:CustomChecksDir -Filter "*.ps1" -EA SilentlyContinue).Count; Ui-Log "Custom checks: $ccCount scripts in $($Global:CustomChecksDir)" "INFO" }
Refresh-Ollama
Ui-Log "Ready. Click 'Scan System' to begin." "INFO"
Ui-SetStatus "Ready."
$W.ShowDialog()|Out-Null
#endregion
