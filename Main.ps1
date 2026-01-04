#requires -Version 5.1
Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase

# -----------------------------
# Paths
# -----------------------------
$root = Split-Path -Parent $MyInvocation.MyCommand.Path

$mainWindowXamlFilePath = Join-Path $root "MainWindow.xaml"
$toolsConfigurationFilePath = Join-Path $root "tools.json"
$skillsConfigurationFilePath = Join-Path $root "skills.json"

$imgDirectoryPath = Join-Path $root "img"
$defaultIconRelativePath = "img\apps.png"
$defaultIconFilePath = Join-Path $root $defaultIconRelativePath

$openInNewIconFilePath = Join-Path $root "img\open_in_new.png"
$searchIconFilePath = Join-Path $root "img\search.png"
$closeIconFilePath = Join-Path $root "img\close.png"

foreach ($p in @($mainWindowXamlFilePath, $toolsConfigurationFilePath, $skillsConfigurationFilePath)) {
    if (!(Test-Path -LiteralPath $p)) { throw "必須ファイルが見つかりません: $p" }
}
if (!(Test-Path -LiteralPath $imgDirectoryPath)) { throw "img フォルダが見つかりません: $imgDirectoryPath" }

# -----------------------------
# Utils (minimum)
# -----------------------------
function Show-Message {
    param([Parameter(Mandatory = $true)][string]$Message, [string]$Title = "Launcher")
    [System.Windows.MessageBox]::Show($Message, $Title, 'OK', 'Information') | Out-Null
}

function Read-JsonFile {
    param([Parameter(Mandatory = $true)][string]$FilePath)
    $json = Get-Content -LiteralPath $FilePath -Raw -Encoding UTF8
    if ([string]::IsNullOrWhiteSpace($json)) { return $null }
    $json | ConvertFrom-Json
}

function Get-PropString {
    param($Obj, [Parameter(Mandatory = $true)][string]$Name, [string]$Default = "")
    if ($null -eq $Obj) { return $Default }
    $p = $Obj.PSObject.Properties[$Name]
    if ($null -eq $p -or $null -eq $p.Value) { return $Default }
    [string]$p.Value
}

function Get-PropBool {
    param($Obj, [Parameter(Mandatory = $true)][string]$Name, [bool]$Default = $false)
    if ($null -eq $Obj) { return $Default }
    $p = $Obj.PSObject.Properties[$Name]
    if ($null -eq $p -or $null -eq $p.Value) { return $Default }
    [bool]$p.Value
}

function To-StringArray {
    param($Value)
    if ($null -eq $Value) { return @() }
    if ($Value -is [string]) {
        if ([string]::IsNullOrWhiteSpace($Value)) { return @() }
        return @([string]$Value)
    }
    if ($Value -is [System.Collections.IEnumerable]) {
        $list = @()
        foreach ($v in $Value) {
            $s = [string]$v
            if (![string]::IsNullOrWhiteSpace($s)) { $list += $s }
        }
        return @($list)
    }
    $s = [string]$Value
    if ([string]::IsNullOrWhiteSpace($s)) { return @() }
    @($s)
}

function Resolve-Env {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return $Value }
    [Environment]::ExpandEnvironmentVariables($Value)
}

function FileUriOrEmpty {
    param([Parameter(Mandatory = $true)][string]$Path)
    if (!(Test-Path -LiteralPath $Path)) { return "" }
    ([System.Uri]::new((Resolve-Path -LiteralPath $Path).Path)).AbsoluteUri
}

function Find-ParentButton {
    param([Parameter(Mandatory = $true)]$SourceObject)

    $current = $SourceObject
    while ($null -ne $current) {
        if ($current -is [System.Windows.Controls.Button]) { return $current }
        if (!($current -is [System.Windows.DependencyObject])) { return $null }
        $current = [System.Windows.Media.VisualTreeHelper]::GetParent($current)
    }
    $null
}

# -----------------------------
# Icons
# -----------------------------
$openInNewIconUri = FileUriOrEmpty $openInNewIconFilePath
$searchIconUri = FileUriOrEmpty $searchIconFilePath
$closeIconUri = FileUriOrEmpty $closeIconFilePath

function Resolve-IconUriForTool {
    param([Parameter(Mandatory = $true)]$ToolDef)

    $iconPathValue = Get-PropString $ToolDef "IconPath" ""
    if (![string]::IsNullOrWhiteSpace($iconPathValue)) {
        $resolved = $iconPathValue
        if (![System.IO.Path]::IsPathRooted($resolved)) { $resolved = Join-Path $root $resolved }
        $u = FileUriOrEmpty $resolved
        if (![string]::IsNullOrWhiteSpace($u)) { return $u }
    }

    FileUriOrEmpty $defaultIconFilePath
}

# -----------------------------
# Load config
# -----------------------------
$toolsConfiguration = Read-JsonFile -FilePath $toolsConfigurationFilePath
$skillsConfiguration = Read-JsonFile -FilePath $skillsConfigurationFilePath
if ($null -eq $toolsConfiguration -or $null -eq $skillsConfiguration) {
    throw "tools.json / skills.json の読み込みに失敗しました。"
}

$toolDefById = @{}
foreach ($t in @($toolsConfiguration.Tools)) {
    $id = Get-PropString $t "Id" ""
    if (![string]::IsNullOrWhiteSpace($id)) { $toolDefById[$id] = $t }
}

# skills.json の ToolIds を「初期状態」として保持（不変）
$baseSkillToolIdsById = @{}
foreach ($s in @($skillsConfiguration.Skills)) {
    $sid = Get-PropString $s "Id" ""
    if ([string]::IsNullOrWhiteSpace($sid)) { continue }
    $toolIds = @()
    $p = $s.PSObject.Properties["ToolIds"]
    if ($null -ne $p) { $toolIds = To-StringArray $p.Value }
    $baseSkillToolIdsById[$sid] = @($toolIds)
}

# セッション中の変更（=追加/解除）を保持するが、スキル切替で毎回リセットする
$sessionSkillToolIdsById = @{}
function Reset-SessionSkillToolIds {
    $script:sessionSkillToolIdsById = @{}
    foreach ($k in $baseSkillToolIdsById.Keys) {
        # 配列をコピー（参照共有しない）
        $script:sessionSkillToolIdsById[$k] = @($baseSkillToolIdsById[$k])
    }
}
Reset-SessionSkillToolIds

function Get-EffectiveToolIdsForSkill {
    param([Parameter(Mandatory = $true)][string]$SkillId)
    if ($sessionSkillToolIdsById.ContainsKey($SkillId)) { return @($sessionSkillToolIdsById[$SkillId]) }
    @()
}

function Set-ToolIdsForSkill {
    param([Parameter(Mandatory = $true)][string]$SkillId, [Parameter(Mandatory = $true)][string[]]$ToolIds)

    $set = @{}
    $list = New-Object System.Collections.Generic.List[string]
    foreach ($id in $ToolIds) {
        $s = [string]$id
        if ([string]::IsNullOrWhiteSpace($s)) { continue }
        if ($set.ContainsKey($s)) { continue }
        $set[$s] = $true
        $null = $list.Add($s)
    }
    $sessionSkillToolIdsById[$SkillId] = @($list.ToArray())
}

# -----------------------------
# Launch
# -----------------------------
function Start-Tool {
    param([Parameter(Mandatory = $true)]$ToolVm)

    $path = Resolve-Env (Get-PropString $ToolVm "Path" "")
    $args = Get-PropString $ToolVm "Args" ""
    $wd = Resolve-Env (Get-PropString $ToolVm "WorkingDirectory" "")
    $admin = Get-PropBool $ToolVm "RunAsAdmin" $false

    try {
        if ($path -match '^(https?://)') { Start-Process $path; return }

        $p = @{ FilePath = $path }
        if (![string]::IsNullOrWhiteSpace($args)) { $p.ArgumentList = $args }
        if (![string]::IsNullOrWhiteSpace($wd)) { $p.WorkingDirectory = $wd }

        if ($admin) { Start-Process @p -Verb RunAs }
        else { Start-Process @p }
    }
    catch {
        Show-Message -Message ("起動に失敗しました:`n{0}" -f $_.Exception.Message)
    }
}

function New-ToolViewModel {
    param([Parameter(Mandatory = $true)]$ToolDef, [Parameter(Mandatory = $true)][bool]$CanAddToSelectedSkill)

    [pscustomobject]@{
        Id                    = Get-PropString $ToolDef "Id" ""
        Name                  = Get-PropString $ToolDef "Name" ""
        Path                  = Get-PropString $ToolDef "Path" ""
        Args                  = Get-PropString $ToolDef "Args" ""
        WorkingDirectory      = Get-PropString $ToolDef "WorkingDirectory" ""
        RunAsAdmin            = Get-PropBool   $ToolDef "RunAsAdmin" $false
        IconImageSource       = (Resolve-IconUriForTool -ToolDef $ToolDef)
        OpenInNewIconSource   = $openInNewIconUri
        CanAddToSelectedSkill = $CanAddToSelectedSkill
    }
}

# -----------------------------
# Load UI
# -----------------------------
$xaml = Get-Content -LiteralPath $mainWindowXamlFilePath -Raw -Encoding UTF8
$sr = New-Object System.IO.StringReader($xaml)
$xr = [System.Xml.XmlReader]::Create($sr)
$mainWindow = [Windows.Markup.XamlReader]::Load($xr)

$mainWindow.DataContext = [pscustomobject]@{
    SearchIconSource = $searchIconUri
    CloseIconSource  = $closeIconUri
}

$skillComboBox = $mainWindow.FindName("skillComboBox")
$applicationSearchTextBox = $mainWindow.FindName("applicationSearchTextBox")
$clearSearchButton = $mainWindow.FindName("clearSearchButton")

$applicationListBox = $mainWindow.FindName("applicationListBox")
$skillApplicationListBox = $mainWindow.FindName("skillApplicationListBox")

$applicationCountTextBlock = $mainWindow.FindName("applicationCountTextBlock")
$skillApplicationCountTextBlock = $mainWindow.FindName("skillApplicationCountTextBlock")
$bulkLaunchButton = $mainWindow.FindName("bulkLaunchButton")

$skillViewModels = New-Object 'System.Collections.ObjectModel.ObservableCollection[object]'
$applicationViewModels = New-Object 'System.Collections.ObjectModel.ObservableCollection[object]'
$skillApplicationViewModels = New-Object 'System.Collections.ObjectModel.ObservableCollection[object]'

$skillComboBox.ItemsSource = $skillViewModels
$applicationListBox.ItemsSource = $applicationViewModels
$skillApplicationListBox.ItemsSource = $skillApplicationViewModels

function Get-SelectedSkillId {
    if ($null -eq $skillComboBox.SelectedValue) { return "" }
    [string]$skillComboBox.SelectedValue
}

function Refresh-SkillViewModels {
    $skillViewModels.Clear()
    foreach ($s in @($skillsConfiguration.Skills)) {
        $vm = [pscustomobject]@{
            Id   = Get-PropString $s "Id" ""
            Name = Get-PropString $s "Name" ""
        }
        $null = $vm | Add-Member -MemberType ScriptMethod -Name ToString -Value { $this.Name } -Force
        $skillViewModels.Add($vm) | Out-Null
    }
}

function Get-ToolIdSetForSelectedSkill {
    $sid = Get-SelectedSkillId
    if ([string]::IsNullOrWhiteSpace($sid)) { return @{} }

    $set = @{}
    foreach ($id in @(Get-EffectiveToolIdsForSkill -SkillId $sid)) {
        if (![string]::IsNullOrWhiteSpace($id)) { $set[$id] = $true }
    }
    $set
}

function Get-FilteredToolDefinitions {
    $q = $applicationSearchTextBox.Text
    if ($null -eq $q) { $q = "" }
    $q = $q.Trim().ToLowerInvariant()

    $result = @()
    foreach ($t in @($toolsConfiguration.Tools)) {
        $name = (Get-PropString $t "Name" "").ToLowerInvariant()
        $path = (Get-PropString $t "Path" "").ToLowerInvariant()
        if ($q.Length -gt 0 -and -not ($name.Contains($q) -or $path.Contains($q))) { continue }
        $result += $t
    }
    $result
}

function Refresh-ApplicationViewModels {
    $applicationViewModels.Clear()

    $sid = Get-SelectedSkillId
    $set = Get-ToolIdSetForSelectedSkill

    foreach ($t in @(Get-FilteredToolDefinitions)) {
        $id = Get-PropString $t "Id" ""
        $canAdd = $true
        if ([string]::IsNullOrWhiteSpace($sid)) { $canAdd = $false }
        if (![string]::IsNullOrWhiteSpace($id) -and $set.ContainsKey($id)) { $canAdd = $false }

        $applicationViewModels.Add((New-ToolViewModel -ToolDef $t -CanAddToSelectedSkill $canAdd)) | Out-Null
    }

    $applicationCountTextBlock.Text = ("表示: {0} 件" -f $applicationViewModels.Count)
}

function Refresh-SkillApplicationViewModels {
    $skillApplicationViewModels.Clear()

    $sid = Get-SelectedSkillId
    if ([string]::IsNullOrWhiteSpace($sid)) {
        $skillApplicationCountTextBlock.Text = "表示: 0 件"
        $bulkLaunchButton.IsEnabled = $false
        return
    }

    foreach ($id in @(Get-EffectiveToolIdsForSkill -SkillId $sid)) {
        if ([string]::IsNullOrWhiteSpace($id)) { continue }
        if (!$toolDefById.ContainsKey($id)) { continue }
        $skillApplicationViewModels.Add((New-ToolViewModel -ToolDef $toolDefById[$id] -CanAddToSelectedSkill $false)) | Out-Null
    }

    $skillApplicationCountTextBlock.Text = ("表示: {0} 件" -f $skillApplicationViewModels.Count)
    $bulkLaunchButton.IsEnabled = ($skillApplicationViewModels.Count -gt 0)
}

function Select-InitialSkill {
    if ($skillViewModels.Count -gt 0) { $skillComboBox.SelectedIndex = 0 }
}

function Add-ToolIdToSkill {
    param([Parameter(Mandatory = $true)][string]$ToolId, [Parameter(Mandatory = $true)][string]$SkillId)
    $cur = @(Get-EffectiveToolIdsForSkill -SkillId $SkillId)
    if ($cur -contains $ToolId) { return }
    Set-ToolIdsForSkill -SkillId $SkillId -ToolIds @($cur + @($ToolId))
}

function Remove-ToolIdFromSkill {
    param([Parameter(Mandatory = $true)][string]$ToolId, [Parameter(Mandatory = $true)][string]$SkillId)
    $cur = @(Get-EffectiveToolIdsForSkill -SkillId $SkillId)
    $next = @()
    foreach ($x in $cur) { if ($x -ne $ToolId) { $next += $x } }
    Set-ToolIdsForSkill -SkillId $SkillId -ToolIds $next
}

# 初期描画
Refresh-SkillViewModels
Select-InitialSkill
Refresh-SkillApplicationViewModels
Refresh-ApplicationViewModels

# -----------------------------
# Events
# -----------------------------
$skillComboBox.Add_SelectionChanged({
        # ★ここが要点：スキルを切り替えたら、前回の追加/解除を破棄して初期状態に戻す
        Reset-SessionSkillToolIds

        Refresh-SkillApplicationViewModels
        Refresh-ApplicationViewModels
    })

$applicationSearchTextBox.Add_TextChanged({
        Refresh-ApplicationViewModels
    })

$clearSearchButton.Add_Click({
        $applicationSearchTextBox.Text = ""
        Refresh-ApplicationViewModels
    })

$applicationListBox.AddHandler(
    [System.Windows.Controls.Primitives.ButtonBase]::ClickEvent,
    [System.Windows.RoutedEventHandler] {
        param($sender, $eventArgs)

        $btn = Find-ParentButton -SourceObject $eventArgs.OriginalSource
        if ($null -eq $btn) { return }

        if ($btn.Name -eq "openFromAllApplicationsNameButton") {
            if ($null -ne $btn.Tag) { Start-Tool -ToolVm $btn.Tag }
            return
        }

        if ($btn.Name -eq "addToSkillButton") {
            $vm = $btn.Tag
            if ($null -eq $vm) { return }

            $sid = Get-SelectedSkillId
            if ([string]::IsNullOrWhiteSpace($sid)) { return }

            Add-ToolIdToSkill -ToolId ([string]$vm.Id) -SkillId $sid
            Refresh-SkillApplicationViewModels
            Refresh-ApplicationViewModels
            return
        }
    }
)

$skillApplicationListBox.AddHandler(
    [System.Windows.Controls.Primitives.ButtonBase]::ClickEvent,
    [System.Windows.RoutedEventHandler] {
        param($sender, $eventArgs)

        $btn = Find-ParentButton -SourceObject $eventArgs.OriginalSource
        if ($null -eq $btn) { return }

        if ($btn.Name -eq "openFromSkillApplicationsNameButton") {
            if ($null -ne $btn.Tag) { Start-Tool -ToolVm $btn.Tag }
            return
        }

        if ($btn.Name -eq "unlinkFromSkillButton") {
            $vm = $btn.Tag
            if ($null -eq $vm) { return }

            $sid = Get-SelectedSkillId
            if ([string]::IsNullOrWhiteSpace($sid)) { return }

            Remove-ToolIdFromSkill -ToolId ([string]$vm.Id) -SkillId $sid
            Refresh-SkillApplicationViewModels
            Refresh-ApplicationViewModels
            return
        }
    }
)

$bulkLaunchButton.Add_Click({
        if ($skillApplicationViewModels.Count -le 0) { return }
        foreach ($vm in @($skillApplicationViewModels)) {
            Start-Tool -ToolVm $vm
            Start-Sleep -Milliseconds 150
        }
    })

$null = $mainWindow.ShowDialog()