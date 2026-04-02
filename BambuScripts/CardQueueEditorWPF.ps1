# ════════════════════════════════════════════════════════════════════════════════
# WPF BATCH PRE-FLIGHT EDITOR SKELETON
# ════════════════════════════════════════════════════════════════════════════════

# Load WPF & WinForms Assemblies
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.IO.Compression.FileSystem

$ErrorActionPreference = 'Stop'

# --- THE XAML LAYOUT (The Blueprint for the Main Window) ---
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Batch Pre-Flight Editor - WPF Engine"
        Width="1550" Height="850" MinWidth="1200" MinHeight="600"
        Background="#16171B" WindowStartupLocation="CenterScreen">
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="60"/>
            <RowDefinition Height="*"/>
        </Grid.RowDefinitions>

        <Border Background="#1C1D23" Grid.Row="0">
            <Grid>
                <TextBlock Name="LblGlobalTitle" Text="Loading files into queue..." Foreground="White" FontSize="18" FontWeight="Bold" VerticalAlignment="Center" Margin="15,0,0,0"/>
                <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" VerticalAlignment="Center" Margin="0,0,15,0">
                    <Button Name="BtnCombineData" Content="Combine TSV Data" Background="#E8A135" Foreground="White" FontWeight="Bold" Width="180" Height="40" Margin="0,0,15,0" BorderThickness="0" Cursor="Hand"/>
                    <Button Name="BtnProcessAll" Content="Add All To Queue" Background="#4CAF72" Foreground="White" FontWeight="Bold" Width="220" Height="40" IsEnabled="False" BorderThickness="0" Cursor="Hand"/>
                </StackPanel>
            </Grid>
        </Border>

        <ScrollViewer Grid.Row="1" Background="#0D0E10" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled">
            <StackPanel Name="MainStack" Orientation="Vertical" Margin="15"/>
        </ScrollViewer>
    </Grid>
</Window>
"@

# Read the XAML into a live Window object
$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [System.Windows.Markup.XamlReader]::Load($reader)

# Map our named UI elements to PowerShell variables
$lblGlobalTitle = $window.FindName("LblGlobalTitle")
$btnProcessAll  = $window.FindName("BtnProcessAll")
$btnCombineData = $window.FindName("BtnCombineData")
$mainStack      = $window.FindName("MainStack")

# --- HELPER FUNCTIONS ---
function Get-WpfColor([string]$hex) {
    if ([string]::IsNullOrWhiteSpace($hex)) { $hex = "#808080" }
    if ($hex.Length -eq 9) { $hex = "#" + $hex.Substring(1,6) } # Strip alpha if RGBA
    return [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString($hex))
}

function Create-TextBlock([string]$text, [string]$hexColor, [int]$fontSize, [string]$weight) {
    $tb = New-Object System.Windows.Controls.TextBlock
    $tb.Text = $text
    $tb.Foreground = Get-WpfColor $hexColor
    $tb.FontSize = $fontSize
    if ($weight -eq "Bold") { $tb.FontWeight = [System.Windows.FontWeights]::Bold }
    $tb.VerticalAlignment = "Center"
    return $tb
}

# --- 1. MOCK QUEUE SETUP ---
$gpQueue = [ordered]@{
    "ROOT_FantasyTest" = [ordered]@{
        "C:\MockPath\Cyclops" = "Cyclops_Full.3mf"
        "C:\MockPath\Dragon" = "Dragon_Full.3mf"
    }
}
$script:jobs = New-Object System.Collections.ArrayList

# --- 2. DYNAMIC UI GENERATION (WPF Style) ---

function Build-PJob($parentPath, $anchorFile, $gpJob) {
    $pJob = @{
        FolderPath = $parentPath
        IsDone = $false
        IsQueued = $false
    }

    $pBorder = New-Object System.Windows.Controls.Border
    $pBorder.Background = Get-WpfColor "#1A1C22"
    $pBorder.BorderBrush = Get-WpfColor "#2A2C35"
    $pBorder.BorderThickness = New-Object System.Windows.Thickness(1)
    $pBorder.Margin = New-Object System.Windows.Thickness(0,0,0,10)
    $pBorder.CornerRadius = New-Object System.Windows.CornerRadius(4)
    $pBorder.Padding = New-Object System.Windows.Thickness(10)

    $pGrid = New-Object System.Windows.Controls.Grid
    $pGrid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{Width=[System.Windows.GridLength]::Auto}))
    $pGrid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{Width=(New-Object System.Windows.GridLength(1, [System.Windows.GridUnitType]::Star))}))
    $pBorder.Child = $pGrid

    # ─── LEFT COLUMN: The "Card" Area ───
    $leftStack = New-Object System.Windows.Controls.StackPanel
    $leftStack.Orientation = "Horizontal"
    [System.Windows.Controls.Grid]::SetColumn($leftStack, 0)
    $pGrid.Children.Add($leftStack) | Out-Null

    # Main Image Card
    $cardBorder = New-Object System.Windows.Controls.Border
    $cardBorder.Background = Get-WpfColor "#000000"
    $cardBorder.BorderBrush = Get-WpfColor "#2A2C35"
    $cardBorder.BorderThickness = New-Object System.Windows.Thickness(1)
    $cardBorder.Width = 350
    $cardBorder.Height = 350
    $cardBorder.Margin = New-Object System.Windows.Thickness(0,0,15,0)

    $cardContent = New-Object System.Windows.Controls.Grid
    $tbCard = Create-TextBlock "[MOCK IMAGE PREVIEW]" "#A0A0A0" 14 "Bold"
    $tbCard.HorizontalAlignment = "Center"
    $cardContent.Children.Add($tbCard) | Out-Null
    $cardBorder.Child = $cardContent
    $leftStack.Children.Add($cardBorder) | Out-Null

    # Pick Image Card
    $pickBorder = New-Object System.Windows.Controls.Border
    $pickBorder.Background = Get-WpfColor "#000000"
    $pickBorder.BorderBrush = Get-WpfColor "#2A2C35"
    $pickBorder.BorderThickness = New-Object System.Windows.Thickness(1)
    $pickBorder.Width = 350
    $pickBorder.Height = 350

    $pickContent = New-Object System.Windows.Controls.Grid
    $tbPick = Create-TextBlock "[MOCK GCODE PREVIEW]" "#D95F5F" 14 "Bold"
    $tbPick.HorizontalAlignment = "Center"
    $pickContent.Children.Add($tbPick) | Out-Null
    $pickBorder.Child = $pickContent
    $leftStack.Children.Add($pickBorder) | Out-Null

    # ─── RIGHT COLUMN: The UI Controls ───
    $rightStack = New-Object System.Windows.Controls.StackPanel
    $rightStack.Orientation = "Vertical"
    $rightStack.Margin = New-Object System.Windows.Thickness(15,0,0,0)
    [System.Windows.Controls.Grid]::SetColumn($rightStack, 1)
    $pGrid.Children.Add($rightStack) | Out-Null

    # Header Row (Folder Name + Buttons)
    $headerGrid = New-Object System.Windows.Controls.Grid
    $headerGrid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{Width=(New-Object System.Windows.GridLength(1, [System.Windows.GridUnitType]::Star))}))
    $headerGrid.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{Width=[System.Windows.GridLength]::Auto}))

    $lblFolder = Create-TextBlock "Folder: $(Split-Path $parentPath -Leaf)" "#FFFFFF" 14 "Bold"
    [System.Windows.Controls.Grid]::SetColumn($lblFolder, 0)
    $headerGrid.Children.Add($lblFolder) | Out-Null

    $btnStack = New-Object System.Windows.Controls.StackPanel
    $btnStack.Orientation = "Horizontal"
    [System.Windows.Controls.Grid]::SetColumn($btnStack, 1)

    $btnRefresh = New-Object System.Windows.Controls.Button
    $btnRefresh.Content = "Refresh"
    $btnRefresh.Background = Get-WpfColor "#2A2C35"
    $btnRefresh.Foreground = Get-WpfColor "#FFFFFF"
    $btnRefresh.Width = 100; $btnRefresh.Height = 25; $btnRefresh.BorderThickness = 0
    $btnRefresh.Margin = New-Object System.Windows.Thickness(0,0,10,0); $btnRefresh.Cursor = [System.Windows.Input.Cursors]::Hand
    $btnStack.Children.Add($btnRefresh) | Out-Null

    $btnRemove = New-Object System.Windows.Controls.Button
    $btnRemove.Content = "Remove Folder"
    $btnRemove.Background = Get-WpfColor "#D95F5F"
    $btnRemove.Foreground = Get-WpfColor "#FFFFFF"
    $btnRemove.Width = 120; $btnRemove.Height = 25; $btnRemove.BorderThickness = 0
    $btnRemove.Cursor = [System.Windows.Input.Cursors]::Hand

    $btnRemove.Add_Click({
        $gpJob.ParentListStack.Children.Remove($pBorder)
        $gpJob.Parents.Remove($pJob) | Out-Null
        if ($gpJob.Parents.Count -eq 0) {
            $mainStack.Children.Remove($gpJob.Container)
            $script:jobs.Remove($gpJob) | Out-Null
        }
    })
    $btnStack.Children.Add($btnRemove) | Out-Null

    $headerGrid.Children.Add($btnStack) | Out-Null
    $rightStack.Children.Add($headerGrid) | Out-Null

    # Edit Boxes Box
    $editBox = New-Object System.Windows.Controls.Border
    $editBox.Background = Get-WpfColor "#1C1D23"
    $editBox.BorderBrush = Get-WpfColor "#2A2C35"
    $editBox.BorderThickness = New-Object System.Windows.Thickness(1)
    $editBox.Margin = New-Object System.Windows.Thickness(0,15,0,0)
    $editBox.Padding = New-Object System.Windows.Thickness(10)

    $editStack = New-Object System.Windows.Controls.StackPanel
    $editStack.Orientation = "Horizontal"

    $charStack = New-Object System.Windows.Controls.StackPanel
    $charStack.Margin = New-Object System.Windows.Thickness(0,0,20,0)
    $charStack.Children.Add((Create-TextBlock "Character *" "#A0A0A0" 12 "Normal")) | Out-Null
    $tbChar = New-Object System.Windows.Controls.TextBox
    $tbChar.Text = "MockCharacter"
    $tbChar.Width = 200; $tbChar.Background = Get-WpfColor "#1E2028"; $tbChar.Foreground = Get-WpfColor "#FFFFFF"
    $charStack.Children.Add($tbChar) | Out-Null
    $editStack.Children.Add($charStack) | Out-Null

    $adjStack = New-Object System.Windows.Controls.StackPanel
    $adjStack.Children.Add((Create-TextBlock "Adjective (Optional)" "#A0A0A0" 12 "Normal")) | Out-Null
    $tbAdj = New-Object System.Windows.Controls.TextBox
    $tbAdj.Width = 200; $tbAdj.Background = Get-WpfColor "#1E2028"; $tbAdj.Foreground = Get-WpfColor "#FFFFFF"
    $adjStack.Children.Add($tbAdj) | Out-Null
    $editStack.Children.Add($adjStack) | Out-Null

    $editBox.Child = $editStack
    $rightStack.Children.Add($editBox) | Out-Null

    # The Add to Queue Button
    $btnApply = New-Object System.Windows.Controls.Button
    $btnApply.Content = "Add to Queue"
    $btnApply.Background = Get-WpfColor "#4CAF72"
    $btnApply.Foreground = Get-WpfColor "#FFFFFF"
    $btnApply.FontWeight = [System.Windows.FontWeights]::Bold
    $btnApply.Width = 150; $btnApply.Height = 35; $btnApply.BorderThickness = 0
    $btnApply.HorizontalAlignment = "Right"
    $btnApply.Margin = New-Object System.Windows.Thickness(0,15,0,0); $btnApply.Cursor = [System.Windows.Input.Cursors]::Hand
    $rightStack.Children.Add($btnApply) | Out-Null

    $gpJob.ParentListStack.Children.Add($pBorder) | Out-Null
    return $pJob
}

function Build-GpJob($gpPath, $parentDict) {
    $gpName = if ($gpPath -notlike "ROOT_*") { Split-Path $gpPath -Leaf } else { "(No Parent Folder)" }

    $gpJob = @{
        GpPath = $gpPath
        Parents = New-Object System.Collections.ArrayList
    }
    $script:jobs.Add($gpJob) | Out-Null

    $container = New-Object System.Windows.Controls.Border
    $container.Background = Get-WpfColor "#1C1D23"
    $container.BorderBrush = Get-WpfColor "#2A2C35"
    $container.BorderThickness = New-Object System.Windows.Thickness(1)
    $container.Margin = New-Object System.Windows.Thickness(0,0,0,20)
    $container.CornerRadius = New-Object System.Windows.CornerRadius(6)

    $gpStack = New-Object System.Windows.Controls.StackPanel
    $container.Child = $gpStack
    $gpJob.Container = $container

    $headerGrid = New-Object System.Windows.Controls.Grid
    $headerGrid.Background = Get-WpfColor "#2A2C35"
    $headerGrid.Height = 60

    $lblGP = Create-TextBlock "Grandparent Theme:  $gpName" "#E8A135" 14 "Bold"
    $lblGP.Margin = New-Object System.Windows.Thickness(15,0,0,0)
    $lblGP.HorizontalAlignment = "Left"
    $headerGrid.Children.Add($lblGP) | Out-Null

    $btnRemoveGp = New-Object System.Windows.Controls.Button
    $btnRemoveGp.Content = "Remove Group"
    $btnRemoveGp.Background = Get-WpfColor "#D95F5F"
    $btnRemoveGp.Foreground = Get-WpfColor "#FFFFFF"
    $btnRemoveGp.Width = 140; $btnRemoveGp.Height = 30; $btnRemoveGp.BorderThickness = 0
    $btnRemoveGp.HorizontalAlignment = "Right"; $btnRemoveGp.Margin = New-Object System.Windows.Thickness(0,0,15,0)
    $btnRemoveGp.Cursor = [System.Windows.Input.Cursors]::Hand
    $btnRemoveGp.Add_Click({
        $mainStack.Children.Remove($container)
        $script:jobs.Remove($gpJob) | Out-Null
    })
    $headerGrid.Children.Add($btnRemoveGp) | Out-Null

    $gpStack.Children.Add($headerGrid) | Out-Null

    $parentListStack = New-Object System.Windows.Controls.StackPanel
    $parentListStack.Margin = New-Object System.Windows.Thickness(15)
    $gpStack.Children.Add($parentListStack) | Out-Null
    $gpJob.ParentListStack = $parentListStack

    foreach ($pKey in $parentDict.Keys) {
        $pJob = Build-PJob $pKey $parentDict[$pKey] $gpJob
        $gpJob.Parents.Add($pJob) | Out-Null
    }

    $mainStack.Children.Add($container) | Out-Null
}

# --- 3. STARTUP LOGIC ---

$window.Add_Loaded({
    $idx = 1
    foreach ($gpPath in $gpQueue.Keys) {
        $lblGlobalTitle.Text = "Extracting & Analyzing Group $idx of $($gpQueue.Count)..."
        # Native WPF way to force the UI to refresh its text immediately
        [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke([System.Action]{}, [System.Windows.Threading.DispatcherPriority]::Background)

        Build-GpJob $gpPath $gpQueue[$gpPath]
        $idx++
    }

    $lblGlobalTitle.Text = "Queue Dashboard ($($gpQueue.Count) Theme(s) found)"
    $btnProcessAll.IsEnabled = $true
})

# Launch the WPF Window!
$window.ShowDialog() | Out-Null