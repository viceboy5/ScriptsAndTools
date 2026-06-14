# ════════════════════════════════════════════════════════════════════════════════
# DEBUG LAUNCHER for CardQueueEditorWPF.ps1
# Run this instead of the real editor — it catches runtime errors and shows
# a scrollable WPF error dialog so you can read the full stack trace.
# ════════════════════════════════════════════════════════════════════════════════

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

function Show-ErrorDialog([string]$title, [string]$message) {
    $xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="$title" Width="820" Height="540"
        WindowStartupLocation="CenterScreen" Background="#1A1B22">
  <Grid Margin="12">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>
    <TextBlock Grid.Row="0" Text="$title" FontSize="15" FontWeight="Bold"
               Foreground="#E05555" Margin="0,0,0,8"/>
    <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto"
                  Background="#0D0E12" BorderBrush="#333" BorderThickness="1">
      <TextBox Name="TbMsg" IsReadOnly="True" TextWrapping="Wrap"
               Background="#0D0E12" Foreground="#E8E0C8" FontFamily="Consolas"
               FontSize="12" BorderThickness="0" Padding="8"
               VerticalScrollBarVisibility="Disabled"/>
    </ScrollViewer>
    <Button Grid.Row="2" Content="Copy to Clipboard" HorizontalAlignment="Left"
            Margin="0,10,0,0" Padding="14,6" Background="#3A5080" Foreground="White"
            BorderThickness="0" Name="BtnCopy"/>
  </Grid>
</Window>
"@
    $reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($xaml))
    $win    = [System.Windows.Markup.XamlReader]::Load($reader)
    $win.FindName("TbMsg").Text = $message
    $win.FindName("BtnCopy").Add_Click({
        [System.Windows.Clipboard]::SetText($win.FindName("TbMsg").Text)
    })
    $win.ShowDialog() | Out-Null
}

# ── Redirect all output to a log so nothing is swallowed ─────────────────────
$logFile = Join-Path $PSScriptRoot "debug_editor.log"
"" | Set-Content $logFile   # clear previous run

$editorScript = Join-Path $PSScriptRoot "CardQueueEditorWPF.ps1"

# Run in same process so WPF works, but wrap with a detailed trap
try {
    # Temporarily redirect warnings/verbose to log
    $global:DebugLog = $logFile

    # Source the script — errors propagate here
    . $editorScript

} catch {
    $ex   = $_
    $full = @"
MESSAGE:
  $($ex.Exception.Message)

SCRIPT POSITION:
  $($ex.InvocationInfo.ScriptName) : line $($ex.InvocationInfo.ScriptLineNumber) char $($ex.InvocationInfo.OffsetInLine)

LINE TEXT:
  $($ex.InvocationInfo.Line.Trim())

STACK TRACE:
$($ex.ScriptStackTrace)

.NET INNER EXCEPTION:
$($ex.Exception.InnerException)
"@

    # Also dump to log file for easy copy
    $full | Set-Content $logFile
    Add-Content $logFile "`n`n--- Full Error Record ---`n$ex"

    Write-Host "`n===== EDITOR LAUNCH ERROR =====`n$full" -ForegroundColor Red
    Show-ErrorDialog "CardQueueEditor Launch Error" $full
}
