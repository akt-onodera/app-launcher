#requires -Version 5.1
Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase

$applicationRootDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path

$mainWindowXamlFilePath = Join-Path $applicationRootDirectory "MainWindow.xaml"
$toolsConfigurationFilePath = Join-Path $applicationRootDirectory "tools.json"
$skillsConfigurationFilePath = Join-Path $applicationRootDirectory "skills.json"

$imgDirectoryPath = Join-Path $applicationRootDirectory "img"
$openInNewIconFilePath = Join-Path $imgDirectoryPath "open_in_new.png"
$searchIconFilePath = Join-Path $imgDirectoryPath "search.png"
$closeIconFilePath = Join-Path $imgDirectoryPath "close.png"

function Read-JsonFile {
    param([Parameter(Mandatory = $true)][string]$FilePath)

    $jsonText = Get-Content -LiteralPath $FilePath -Raw -Encoding UTF8
    if ([string]::IsNullOrWhiteSpace($jsonText)) { return $null }
    return ($jsonText | ConvertFrom-Json)
}

function Get-PropertyString {
    param(
        $Object,
        [Parameter(Mandatory = $true)][string]$PropertyName,
        [string]$DefaultValue = ""
    )

    if ($null -eq $Object) { return $DefaultValue }
    $property = $Object.PSObject.Properties[$PropertyName]
    if ($null -eq $property -or $null -eq $property.Value) { return $DefaultValue }
    return ([string]$property.Value)
}

function Get-PropertyBoolean {
    param(
        $Object,
        [Parameter(Mandatory = $true)][string]$PropertyName,
        [bool]$DefaultValue = $false
    )

    if ($null -eq $Object) { return $DefaultValue }
    $property = $Object.PSObject.Properties[$PropertyName]
    if ($null -eq $property -or $null -eq $property.Value) { return $DefaultValue }
    return ([bool]$property.Value)
}

function Get-PropertyStringArray {
    param(
        $Object,
        [Parameter(Mandatory = $true)][string]$PropertyName
    )

    if ($null -eq $Object) { return @() }
    $property = $Object.PSObject.Properties[$PropertyName]
    if ($null -eq $property -or $null -eq $property.Value) { return @() }

    $value = $property.Value
    if ($value -is [string]) {
        if ([string]::IsNullOrWhiteSpace($value)) { return @() }
        return @([string]$value)
    }

    $resultList = @()
    foreach ($item in @($value)) {
        $text = [string]$item
        if (![string]::IsNullOrWhiteSpace($text)) { $resultList += $text }
    }
    return @($resultList)
}

function Find-ParentButton {
    param([Parameter(Mandatory = $true)]$SourceObject)

    $currentObject = $SourceObject
    while ($null -ne $currentObject) {
        if ($currentObject -is [System.Windows.Controls.Button]) { return $currentObject }
        if (!($currentObject -is [System.Windows.DependencyObject])) { return $null }
        $currentObject = [System.Windows.Media.VisualTreeHelper]::GetParent($currentObject)
    }
    return $null
}

function Resolve-ImagePathForTool {
    param([Parameter(Mandatory = $true)]$ToolDefinition)

    $imagePathValue = Get-PropertyString -Object $ToolDefinition -PropertyName "Image" -DefaultValue ""
    if ([string]::IsNullOrWhiteSpace($imagePathValue)) { return "" }

    if ([System.IO.Path]::IsPathRooted($imagePathValue)) { return $imagePathValue }
    return (Join-Path $applicationRootDirectory $imagePathValue)
}

function Resolve-ToolPath {
    param([Parameter(Mandatory = $true)][string]$PathValue)

    if ([string]::IsNullOrWhiteSpace($PathValue)) { return $PathValue }
    if ($PathValue -match '^(https?://)') { return $PathValue }
    if ([System.IO.Path]::IsPathRooted($PathValue)) { return $PathValue }
    return (Join-Path $applicationRootDirectory $PathValue)
}

$toolsConfiguration = Read-JsonFile -FilePath $toolsConfigurationFilePath
$skillsConfiguration = Read-JsonFile -FilePath $skillsConfigurationFilePath
if ($null -eq $toolsConfiguration -or $null -eq $skillsConfiguration) {
    throw "tools.json / skills.json の読み込みに失敗しました。"
}

$toolDefinitionById = @{}
foreach ($toolDefinition in @($toolsConfiguration.Tools)) {
    $toolId = Get-PropertyString -Object $toolDefinition -PropertyName "Id" -DefaultValue ""
    if (![string]::IsNullOrWhiteSpace($toolId)) { $toolDefinitionById[$toolId] = $toolDefinition }
}

function Start-ExcelWorkbook {
    param(
        [Parameter(Mandatory = $true)][string]$FilePath,
        [Parameter(Mandatory = $true)][bool]$ReadOnly
    )

    $excelApplication = $null
    $workbook = $null

    try {
        $excelApplication = New-Object -ComObject Excel.Application
        $excelApplication.DisplayAlerts = $false
        $excelApplication.Visible = $true

        # UpdateLinks = 0, ReadOnly = $ReadOnly
        $workbook = $excelApplication.Workbooks.Open($FilePath, 0, $ReadOnly)

        return
    }
    finally {
        if ($null -ne $workbook) {
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)
        }
        if ($null -ne $excelApplication) {
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excelApplication)
        }
        return
    }
}

function Start-Tool {
    param([Parameter(Mandatory = $true)]$ToolViewModel)

    $rawPathValue = Get-PropertyString -Object $ToolViewModel -PropertyName "Path" -DefaultValue ""
    $pathValue = Resolve-ToolPath -PathValue ([Environment]::ExpandEnvironmentVariables($rawPathValue))

    $defaultApp = (Get-PropertyString -Object $ToolViewModel -PropertyName "DefaultApp" -DefaultValue "").ToLowerInvariant()
    $readOnly = (Get-PropertyBoolean -Object $ToolViewModel -PropertyName "ReadOnly" -DefaultValue $false)

    if ($pathValue -match '^(https?://)') {
        if ($defaultApp -eq "chrome") {
            Start-Process -FilePath "chrome.exe" -ArgumentList @("--new-window", $pathValue)
            return
        }
        if ($defaultApp -eq "edge") {
            Start-Process -FilePath "msedge.exe" -ArgumentList @("--new-window", $pathValue)
            return
        }

        Start-Process $pathValue
        return
    }
    
    if ($defaultApp -eq "excel") {
        Start-ExcelWorkbook -FilePath $pathValue -ReadOnly $readOnly
        return
    }

    if ($defaultApp -eq "word" -or $defaultApp -eq "powerpoint") {
        $applicationExecutablePath = ""
        if ($defaultApp -eq "word") { $applicationExecutablePath = "winword.exe" }
        if ($defaultApp -eq "powerpoint") { $applicationExecutablePath = "powerpnt.exe" }

        $argumentList = @()
        if ($readOnly) { $argumentList += "/r" }
        $argumentList += $pathValue

        Start-Process -FilePath $applicationExecutablePath -ArgumentList $argumentList
        return
    }
}

function New-ToolViewModel {
    param(
        [Parameter(Mandatory = $true)]$ToolDefinition,
        [Parameter(Mandatory = $true)][bool]$CanAddToSelectedSkill
    )

    return [pscustomobject]@{
        Id                    = Get-PropertyString -Object $ToolDefinition -PropertyName "Id" -DefaultValue ""
        Name                  = Get-PropertyString -Object $ToolDefinition -PropertyName "Name" -DefaultValue ""
        Path                  = Get-PropertyString -Object $ToolDefinition -PropertyName "Path" -DefaultValue ""
        DefaultApp            = Get-PropertyString -Object $ToolDefinition -PropertyName "DefaultApp" -DefaultValue ""
        ReadOnly              = Get-PropertyBoolean -Object $ToolDefinition -PropertyName "ReadOnly" -DefaultValue $false
        IconImageSource       = (Resolve-ImagePathForTool -ToolDefinition $ToolDefinition)
        OpenInNewIconSource   = $openInNewIconFilePath
        CanAddToSelectedSkill = $CanAddToSelectedSkill
    }
}

$xamlContent = Get-Content -LiteralPath $mainWindowXamlFilePath -Raw -Encoding UTF8
$stringReader = New-Object System.IO.StringReader($xamlContent)
$xmlReader = [System.Xml.XmlReader]::Create($stringReader)
$mainWindow = [Windows.Markup.XamlReader]::Load($xmlReader)

$mainWindow.DataContext = [pscustomobject]@{
    SearchIconSource = $searchIconFilePath
    CloseIconSource  = $closeIconFilePath
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

$currentSkillToolIds = @()

function Get-SelectedSkillId {
    if ($null -eq $skillComboBox.SelectedValue) { return "" }
    return ([string]$skillComboBox.SelectedValue)
}

function Refresh-SkillViewModels {
    $skillViewModels.Clear()
    foreach ($skillDefinition in @($skillsConfiguration.Skills)) {
        $skillViewModel = [pscustomobject]@{
            Id   = Get-PropertyString -Object $skillDefinition -PropertyName "Id" -DefaultValue ""
            Name = Get-PropertyString -Object $skillDefinition -PropertyName "Name" -DefaultValue ""
        }
        $null = $skillViewModel | Add-Member -MemberType ScriptMethod -Name ToString -Value { return $this.Name } -Force
        $skillViewModels.Add($skillViewModel) | Out-Null
    }
    return
}

function Reset-CurrentSkillToolIds {
    $selectedSkillId = Get-SelectedSkillId
    if ([string]::IsNullOrWhiteSpace($selectedSkillId)) {
        $script:currentSkillToolIds = @()
        return
    }

    foreach ($skillDefinition in @($skillsConfiguration.Skills)) {
        $skillId = Get-PropertyString -Object $skillDefinition -PropertyName "Id" -DefaultValue ""
        if ($skillId -ne $selectedSkillId) { continue }
        $script:currentSkillToolIds = @(Get-PropertyStringArray -Object $skillDefinition -PropertyName "ToolIds")
        return
    }

    $script:currentSkillToolIds = @()
    return
}

function Refresh-ApplicationViewModels {
    $applicationViewModels.Clear()

    $selectedSkillId = Get-SelectedSkillId

    $toolIdSet = @{}
    foreach ($toolId in @($currentSkillToolIds)) {
        if (![string]::IsNullOrWhiteSpace($toolId)) { $toolIdSet[$toolId] = $true }
    }

    $searchText = $applicationSearchTextBox.Text
    if ($null -eq $searchText) { $searchText = "" }
    $searchText = $searchText.Trim().ToLowerInvariant()

    foreach ($toolDefinition in @($toolsConfiguration.Tools)) {
        $toolNameLower = (Get-PropertyString -Object $toolDefinition -PropertyName "Name" -DefaultValue "").ToLowerInvariant()
        $toolPathLower = (Get-PropertyString -Object $toolDefinition -PropertyName "Path" -DefaultValue "").ToLowerInvariant()
        if ($searchText.Length -gt 0 -and -not ($toolNameLower.Contains($searchText) -or $toolPathLower.Contains($searchText))) { continue }

        $toolId = Get-PropertyString -Object $toolDefinition -PropertyName "Id" -DefaultValue ""

        $canAddToSelectedSkill = $true
        if ([string]::IsNullOrWhiteSpace($selectedSkillId)) { $canAddToSelectedSkill = $false }
        if (![string]::IsNullOrWhiteSpace($toolId) -and $toolIdSet.ContainsKey($toolId)) { $canAddToSelectedSkill = $false }

        $applicationViewModels.Add((New-ToolViewModel -ToolDefinition $toolDefinition -CanAddToSelectedSkill $canAddToSelectedSkill)) | Out-Null
    }

    $applicationCountTextBlock.Text = ("表示: {0} 件" -f $applicationViewModels.Count)
    return
}

function Refresh-SkillApplicationViewModels {
    $skillApplicationViewModels.Clear()

    $selectedSkillId = Get-SelectedSkillId
    if ([string]::IsNullOrWhiteSpace($selectedSkillId)) {
        $skillApplicationCountTextBlock.Text = "表示: 0 件"
        $bulkLaunchButton.IsEnabled = $false
        return
    }

    foreach ($toolId in @($currentSkillToolIds)) {
        if ([string]::IsNullOrWhiteSpace($toolId)) { continue }
        if (!$toolDefinitionById.ContainsKey($toolId)) { continue }

        $toolDefinition = $toolDefinitionById[$toolId]
        $skillApplicationViewModels.Add((New-ToolViewModel -ToolDefinition $toolDefinition -CanAddToSelectedSkill $false)) | Out-Null
    }

    $skillApplicationCountTextBlock.Text = ("表示: {0} 件" -f $skillApplicationViewModels.Count)
    $bulkLaunchButton.IsEnabled = ($skillApplicationViewModels.Count -gt 0)
    return
}

function Add-ToolIdToCurrentSkill {
    param([Parameter(Mandatory = $true)][string]$ToolId)

    if ([string]::IsNullOrWhiteSpace($ToolId)) { return }
    if ($currentSkillToolIds -contains $ToolId) { return }
    $script:currentSkillToolIds = @($currentSkillToolIds + @($ToolId))
    return
}

function Remove-ToolIdFromCurrentSkill {
    param([Parameter(Mandatory = $true)][string]$ToolId)

    $nextToolIds = @()
    foreach ($existingToolId in @($currentSkillToolIds)) {
        if ($existingToolId -ne $ToolId) { $nextToolIds += $existingToolId }
    }
    $script:currentSkillToolIds = @($nextToolIds)
    return
}

Refresh-SkillViewModels
if ($skillViewModels.Count -gt 0) { $skillComboBox.SelectedIndex = 0 }
Reset-CurrentSkillToolIds
Refresh-SkillApplicationViewModels
Refresh-ApplicationViewModels

$skillComboBox.Add_SelectionChanged({
        Reset-CurrentSkillToolIds
        Refresh-SkillApplicationViewModels
        Refresh-ApplicationViewModels
        return
    })

$applicationSearchTextBox.Add_TextChanged({
        Refresh-ApplicationViewModels
        return
    })

$clearSearchButton.Add_Click({
        $applicationSearchTextBox.Text = ""
        Refresh-ApplicationViewModels
        return
    })

$applicationListBox.AddHandler(
    [System.Windows.Controls.Primitives.ButtonBase]::ClickEvent,
    [System.Windows.RoutedEventHandler] {
        param($sender, $eventArgs)

        $clickedButton = Find-ParentButton -SourceObject $eventArgs.OriginalSource
        if ($null -eq $clickedButton) { return }

        if ($clickedButton.Name -eq "openFromAllApplicationsNameButton") {
            if ($null -ne $clickedButton.Tag) { Start-Tool -ToolViewModel $clickedButton.Tag }
            return
        }

        if ($clickedButton.Name -eq "addToSkillButton") {
            $toolViewModel = $clickedButton.Tag
            if ($null -eq $toolViewModel) { return }

            Add-ToolIdToCurrentSkill -ToolId ([string]$toolViewModel.Id)
            Refresh-SkillApplicationViewModels
            Refresh-ApplicationViewModels
            return
        }

        return
    }
)

$skillApplicationListBox.AddHandler(
    [System.Windows.Controls.Primitives.ButtonBase]::ClickEvent,
    [System.Windows.RoutedEventHandler] {
        param($sender, $eventArgs)

        $clickedButton = Find-ParentButton -SourceObject $eventArgs.OriginalSource
        if ($null -eq $clickedButton) { return }

        if ($clickedButton.Name -eq "openFromSkillApplicationsNameButton") {
            if ($null -ne $clickedButton.Tag) { Start-Tool -ToolViewModel $clickedButton.Tag }
            return
        }

        if ($clickedButton.Name -eq "unlinkFromSkillButton") {
            $toolViewModel = $clickedButton.Tag
            if ($null -eq $toolViewModel) { return }

            Remove-ToolIdFromCurrentSkill -ToolId ([string]$toolViewModel.Id)
            Refresh-SkillApplicationViewModels
            Refresh-ApplicationViewModels
            return
        }

        return
    }
)

$bulkLaunchButton.Add_Click({
        if ($skillApplicationViewModels.Count -le 0) { return }
        foreach ($toolViewModel in @($skillApplicationViewModels)) {
            Start-Tool -ToolViewModel $toolViewModel
            Start-Sleep -Milliseconds 150
        }
        return
    })

$null = $mainWindow.ShowDialog()