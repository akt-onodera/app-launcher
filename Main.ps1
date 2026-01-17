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

foreach ($requiredFilePath in @(
        $mainWindowXamlFilePath,
        $toolsConfigurationFilePath,
        $skillsConfigurationFilePath,
        $openInNewIconFilePath,
        $searchIconFilePath,
        $closeIconFilePath
    )) {
    if (!(Test-Path -LiteralPath $requiredFilePath)) {
        throw "必須ファイルが見つかりません: $requiredFilePath"
    }
}

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

function Convert-ToAbsoluteFileUri {
    param([Parameter(Mandatory = $true)][string]$FilePath)

    $absolutePath = (Resolve-Path -LiteralPath $FilePath).Path
    return ([System.Uri]::new($absolutePath)).AbsoluteUri
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

function Resolve-IconUriForTool {
    param([Parameter(Mandatory = $true)]$ToolDefinition)

    $iconPathValue = Get-PropertyString -Object $ToolDefinition -PropertyName "IconPath" -DefaultValue ""
    if ([string]::IsNullOrWhiteSpace($iconPathValue)) { return "" }

    $resolvedIconFilePath = $iconPathValue
    if (![System.IO.Path]::IsPathRooted($resolvedIconFilePath)) {
        $resolvedIconFilePath = Join-Path $applicationRootDirectory $resolvedIconFilePath
    }

    if (!(Test-Path -LiteralPath $resolvedIconFilePath)) { return "" }
    return (Convert-ToAbsoluteFileUri -FilePath $resolvedIconFilePath)
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

$openInNewIconUri = Convert-ToAbsoluteFileUri -FilePath $openInNewIconFilePath
$searchIconUri = Convert-ToAbsoluteFileUri -FilePath $searchIconFilePath
$closeIconUri = Convert-ToAbsoluteFileUri -FilePath $closeIconFilePath

function Start-Tool {
    param([Parameter(Mandatory = $true)]$ToolViewModel)

    $pathValue = [Environment]::ExpandEnvironmentVariables(
        (Get-PropertyString -Object $ToolViewModel -PropertyName "Path" -DefaultValue "")
    )
    $argumentsValue = Get-PropertyString -Object $ToolViewModel -PropertyName "Args" -DefaultValue ""
    $workingDirectoryValue = [Environment]::ExpandEnvironmentVariables(
        (Get-PropertyString -Object $ToolViewModel -PropertyName "WorkingDirectory" -DefaultValue "")
    )
    $runAsAdministrator = Get-PropertyBoolean -Object $ToolViewModel -PropertyName "RunAsAdmin" -DefaultValue $false

    if ($pathValue -match '^(https?://)') {
        Start-Process $pathValue
        return
    }

    $startProcessParameters = @{ FilePath = $pathValue }
    if (![string]::IsNullOrWhiteSpace($argumentsValue)) { $startProcessParameters.ArgumentList = $argumentsValue }
    if (![string]::IsNullOrWhiteSpace($workingDirectoryValue)) { $startProcessParameters.WorkingDirectory = $workingDirectoryValue }

    if ($runAsAdministrator) { Start-Process @startProcessParameters -Verb RunAs }
    else { Start-Process @startProcessParameters }

    return
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
        Args                  = Get-PropertyString -Object $ToolDefinition -PropertyName "Args" -DefaultValue ""
        WorkingDirectory      = Get-PropertyString -Object $ToolDefinition -PropertyName "WorkingDirectory" -DefaultValue ""
        RunAsAdmin            = Get-PropertyBoolean -Object $ToolDefinition -PropertyName "RunAsAdmin" -DefaultValue $false
        IconImageSource       = (Resolve-IconUriForTool -ToolDefinition $ToolDefinition)
        OpenInNewIconSource   = $openInNewIconUri
        CanAddToSelectedSkill = $CanAddToSelectedSkill
    }
}

$xamlContent = Get-Content -LiteralPath $mainWindowXamlFilePath -Raw -Encoding UTF8
$stringReader = New-Object System.IO.StringReader($xamlContent)
$xmlReader = [System.Xml.XmlReader]::Create($stringReader)
$mainWindow = [Windows.Markup.XamlReader]::Load($xmlReader)

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