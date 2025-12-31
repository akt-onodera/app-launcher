#requires -Version 5.1
Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase

$applicationRootDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path
$mainWindowXamlFilePath = Join-Path $applicationRootDirectory "MainWindow.xaml"
$toolsConfigurationFilePath = Join-Path $applicationRootDirectory "tools.json"
$skillsConfigurationFilePath = Join-Path $applicationRootDirectory "skills.json"

$iconsDirectoryPath = Join-Path $applicationRootDirectory "icons"
$defaultIconRelativePath = "icons\apps.png"
$defaultIconFilePath = Join-Path $applicationRootDirectory $defaultIconRelativePath

$openInNewIconFilePath = Join-Path $applicationRootDirectory "icons\open_in_new.png"
$searchIconFilePath = Join-Path $applicationRootDirectory "icons\search.png"
$closeIconFilePath = Join-Path $applicationRootDirectory "icons\close.png"

$userDataDirectoryPath = Join-Path $env:APPDATA "ESS-ToolLauncher"
$userSettingsFilePath = Join-Path $userDataDirectoryPath "settings.json"
$userSkillToolIdsFilePath = Join-Path $userDataDirectoryPath "skillToolIds.json"

if (!(Test-Path -LiteralPath $userDataDirectoryPath)) {
    New-Item -ItemType Directory -Path $userDataDirectoryPath | Out-Null
}

function Show-Message {
    param(
        [Parameter(Mandatory = $true)][string]$message,
        [string]$title = "Launcher"
    )
    [System.Windows.MessageBox]::Show($message, $title, 'OK', 'Information') | Out-Null
}

function Read-JsonFile {
    param([Parameter(Mandatory = $true)][string]$filePath)

    if (!(Test-Path -LiteralPath $filePath)) { return $null }
    $jsonText = Get-Content -LiteralPath $filePath -Raw -Encoding UTF8
    if ([string]::IsNullOrWhiteSpace($jsonText)) { return $null }
    $jsonText | ConvertFrom-Json
}

function Write-JsonFile {
    param(
        [Parameter(Mandatory = $true)][string]$filePath,
        [Parameter(Mandatory = $true)]$dataObject
    )
    ($dataObject | ConvertTo-Json -Depth 10) | Set-Content -LiteralPath $filePath -Encoding UTF8
}

function Resolve-EnvironmentVariables {
    param([string]$value)
    if ([string]::IsNullOrWhiteSpace($value)) { return $value }
    [Environment]::ExpandEnvironmentVariables($value)
}

function Get-PropertyString {
    param($object, [string]$propertyName, [string]$defaultValue = "")
    if ($null -eq $object) { return $defaultValue }
    $property = $object.PSObject.Properties[$propertyName]
    if ($null -eq $property -or $null -eq $property.Value) { return $defaultValue }
    [string]$property.Value
}

function Get-PropertyBoolean {
    param($object, [string]$propertyName, [bool]$defaultValue = $false)
    if ($null -eq $object) { return $defaultValue }
    $property = $object.PSObject.Properties[$propertyName]
    if ($null -eq $property -or $null -eq $property.Value) { return $defaultValue }
    [bool]$property.Value
}

function Find-ParentButton {
    param([Parameter(Mandatory = $true)]$sourceObject)

    $currentObject = $sourceObject
    while ($null -ne $currentObject) {
        if ($currentObject -is [System.Windows.Controls.Button]) { return $currentObject }
        if (!($currentObject -is [System.Windows.DependencyObject])) { return $null }
        $currentObject = [System.Windows.Media.VisualTreeHelper]::GetParent($currentObject)
    }
    $null
}

function Convert-FilePathToFileUri {
    param([Parameter(Mandatory = $true)][string]$filePath)

    try { return ([System.Uri]::new($filePath)).AbsoluteUri }
    catch {
        $normalizedPath = $filePath.Replace("\", "/")
        if ($normalizedPath -notmatch '^[a-zA-Z]:/') { return "" }
        "file:///$normalizedPath"
    }
}

function Get-StaticIconUri {
    param([Parameter(Mandatory = $true)][string]$filePath)
    if (Test-Path -LiteralPath $filePath) { return (Convert-FilePathToFileUri -filePath $filePath) }
    ""
}

$openInNewIconUri = Get-StaticIconUri -filePath $openInNewIconFilePath
$searchIconUri = Get-StaticIconUri -filePath $searchIconFilePath
$closeIconUri = Get-StaticIconUri -filePath $closeIconFilePath

function Ensure-DefaultConfigurationFiles {
    param(
        [Parameter(Mandatory = $true)][string]$toolsConfigurationFilePath,
        [Parameter(Mandatory = $true)][string]$skillsConfigurationFilePath,
        [Parameter(Mandatory = $true)][string]$iconsDirectoryPath,
        [Parameter(Mandatory = $true)][string]$defaultIconRelativePath
    )

    if (!(Test-Path -LiteralPath $iconsDirectoryPath)) {
        New-Item -ItemType Directory -Path $iconsDirectoryPath | Out-Null
    }

    if (!(Test-Path -LiteralPath $toolsConfigurationFilePath)) {
        $defaultTools = @{
            Tools = @(
                @{ Id = "gmail"; Name = "Gmail"; Path = "https://mail.google.com/"; Args = ""; WorkingDirectory = ""; RunAsAdmin = $false; IconPath = $defaultIconRelativePath },
                @{ Id = "notepad"; Name = "メモ帳"; Path = "notepad.exe"; Args = ""; WorkingDirectory = ""; RunAsAdmin = $false; IconPath = $defaultIconRelativePath },
                @{ Id = "calc"; Name = "電卓"; Path = "calc.exe"; Args = ""; WorkingDirectory = ""; RunAsAdmin = $false; IconPath = $defaultIconRelativePath }
            )
        }
        Write-JsonFile -filePath $toolsConfigurationFilePath -dataObject $defaultTools
    }

    if (!(Test-Path -LiteralPath $skillsConfigurationFilePath)) {
        $defaultSkills = @{
            Skills = @(
                @{ Id = "general"; Name = "General"; ToolIds = @("gmail") },
                @{ Id = "dev"; Name = "Dev"; ToolIds = @() },
                @{ Id = "support"; Name = "Support"; ToolIds = @() }
            )
        }
        Write-JsonFile -filePath $skillsConfigurationFilePath -dataObject $defaultSkills
    }
}

function Normalize-UserSettings {
    param($userSettingsObject)

    if ($null -eq $userSettingsObject) {
        return [pscustomobject]@{ SelectedSkillId = "" }
    }

    if ($userSettingsObject -is [System.Collections.IDictionary]) {
        $userSettingsObject = [pscustomobject]$userSettingsObject
    }

    if ($null -eq $userSettingsObject.PSObject.Properties["SelectedSkillId"]) {
        Add-Member -InputObject $userSettingsObject -MemberType NoteProperty -Name "SelectedSkillId" -Value "" -Force
    }

    $userSettingsObject
}

function Resolve-IconFilePath {
    param(
        [Parameter(Mandatory = $true)][string]$iconPathValue,
        [Parameter(Mandatory = $true)][string]$applicationRootDirectory
    )

    $resolvedIconFilePath = $iconPathValue
    if (![System.IO.Path]::IsPathRooted($resolvedIconFilePath)) {
        $resolvedIconFilePath = Join-Path $applicationRootDirectory $resolvedIconFilePath
    }
    $resolvedIconFilePath
}

function Get-IconImageSourceForTool {
    param(
        [Parameter(Mandatory = $true)]$toolDefinition,
        [Parameter(Mandatory = $true)][string]$applicationRootDirectory,
        [Parameter(Mandatory = $true)][string]$defaultIconFilePath
    )

    $iconPathValue = Get-PropertyString $toolDefinition "IconPath" ""
    if (![string]::IsNullOrWhiteSpace($iconPathValue)) {
        $resolvedIconFilePath = Resolve-IconFilePath -iconPathValue $iconPathValue -applicationRootDirectory $applicationRootDirectory
        if (Test-Path -LiteralPath $resolvedIconFilePath) {
            return (Convert-FilePathToFileUri -filePath $resolvedIconFilePath)
        }
    }

    if (Test-Path -LiteralPath $defaultIconFilePath) {
        return (Convert-FilePathToFileUri -filePath $defaultIconFilePath)
    }

    ""
}

function New-ToolViewModel {
    param(
        [Parameter(Mandatory = $true)]$toolDefinition,
        [Parameter(Mandatory = $true)][bool]$canAddToSelectedSkill
    )

    $iconImageSource = Get-IconImageSourceForTool `
        -toolDefinition $toolDefinition `
        -applicationRootDirectory $applicationRootDirectory `
        -defaultIconFilePath $defaultIconFilePath

    [pscustomobject]@{
        Id                    = Get-PropertyString $toolDefinition "Id" ""
        Name                  = Get-PropertyString $toolDefinition "Name" ""
        Path                  = Get-PropertyString $toolDefinition "Path" ""
        Args                  = Get-PropertyString $toolDefinition "Args" ""
        WorkingDirectory      = Get-PropertyString $toolDefinition "WorkingDirectory" ""
        RunAsAdmin            = Get-PropertyBoolean $toolDefinition "RunAsAdmin" $false
        IconImageSource       = $iconImageSource
        OpenInNewIconSource   = $openInNewIconUri
        CanAddToSelectedSkill = $canAddToSelectedSkill
    }
}

function Start-Tool {
    param([Parameter(Mandatory = $true)]$toolViewModel)

    $pathValue = Resolve-EnvironmentVariables (Get-PropertyString $toolViewModel "Path" "")
    $argumentsValue = Get-PropertyString $toolViewModel "Args" ""
    $workingDirectoryValue = Resolve-EnvironmentVariables (Get-PropertyString $toolViewModel "WorkingDirectory" "")
    $runAsAdministrator = Get-PropertyBoolean $toolViewModel "RunAsAdmin" $false

    try {
        if ($pathValue -match '^(https?://)') {
            Start-Process $pathValue
            return
        }

        $startProcessParameters = @{ FilePath = $pathValue }

        if (![string]::IsNullOrWhiteSpace($argumentsValue)) {
            $startProcessParameters.ArgumentList = $argumentsValue
        }
        if (![string]::IsNullOrWhiteSpace($workingDirectoryValue)) {
            $startProcessParameters.WorkingDirectory = $workingDirectoryValue
        }

        if ($runAsAdministrator) {
            Start-Process @startProcessParameters -Verb RunAs
        }
        else {
            Start-Process @startProcessParameters
        }
    }
    catch {
        Show-Message -message ("起動に失敗しました:`n{0}" -f $_.Exception.Message)
    }
}

Ensure-DefaultConfigurationFiles `
    -toolsConfigurationFilePath $toolsConfigurationFilePath `
    -skillsConfigurationFilePath $skillsConfigurationFilePath `
    -iconsDirectoryPath $iconsDirectoryPath `
    -defaultIconRelativePath $defaultIconRelativePath

$toolsConfiguration = Read-JsonFile -filePath $toolsConfigurationFilePath
$skillsConfiguration = Read-JsonFile -filePath $skillsConfigurationFilePath
if ($null -eq $toolsConfiguration -or $null -eq $skillsConfiguration) {
    throw "tools.json / skills.json を読み込めません。"
}

$userSettings = Normalize-UserSettings (Read-JsonFile -filePath $userSettingsFilePath)

$toolDefinitionById = @{}
foreach ($toolDefinition in @($toolsConfiguration.Tools)) {
    $toolIdentifier = Get-PropertyString $toolDefinition "Id" ""
    if (![string]::IsNullOrWhiteSpace($toolIdentifier)) {
        $toolDefinitionById[$toolIdentifier] = $toolDefinition
    }
}

$knownToolIds = @($toolDefinitionById.Keys) | Sort-Object Length -Descending

function Find-SkillDefinitionById {
    param([Parameter(Mandatory = $true)][string]$skillId)

    foreach ($skillDefinition in @($skillsConfiguration.Skills)) {
        if ((Get-PropertyString $skillDefinition "Id" "") -eq $skillId) {
            return $skillDefinition
        }
    }
    $null
}

function Normalize-ToolIdList {
    param(
        $toolIdsValue,
        [string[]]$knownToolIds
    )

    if ($null -eq $toolIdsValue) { return @() }

    if ($toolIdsValue -is [System.Collections.IEnumerable] -and -not ($toolIdsValue -is [string])) {
        $normalizedList = @()
        foreach ($item in $toolIdsValue) {
            $text = [string]$item
            if (![string]::IsNullOrWhiteSpace($text)) { $normalizedList += $text }
        }
        return @($normalizedList)
    }

    $toolIdsText = [string]$toolIdsValue
    if ([string]::IsNullOrWhiteSpace($toolIdsText)) { return @() }

    if ($toolIdsText -match '[,;\s]') {
        @($toolIdsText -split '[,;\s]+' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    }
    else {
        $recoveredList = @()
        $remainingText = $toolIdsText

        while ($remainingText.Length -gt 0) {
            $matchedToolId = $null
            foreach ($candidate in $knownToolIds) {
                if ($remainingText.StartsWith($candidate)) { $matchedToolId = $candidate; break }
            }

            if ($null -eq $matchedToolId) { $recoveredList += $remainingText; break }
            $recoveredList += $matchedToolId
            $remainingText = $remainingText.Substring($matchedToolId.Length)
        }

        @($recoveredList)
    }
}

function Load-UserSkillToolIdsMap {
    param(
        [Parameter(Mandatory = $true)][string]$filePath,
        [Parameter(Mandatory = $true)][string[]]$knownToolIds
    )

    $mapObject = Read-JsonFile -filePath $filePath
    if ($null -eq $mapObject) { return @{} }

    $property = $mapObject.PSObject.Properties["SkillToolIdsById"]
    if ($null -eq $property -or $null -eq $property.Value) { return @{} }

    $resultMap = @{}
    foreach ($entryProperty in $property.Value.PSObject.Properties) {
        $skillId = [string]$entryProperty.Name
        $resultMap[$skillId] = (Normalize-ToolIdList -toolIdsValue $entryProperty.Value -knownToolIds $knownToolIds)
    }
    $resultMap
}

function Save-UserSkillToolIdsMap {
    param(
        [Parameter(Mandatory = $true)][string]$filePath,
        [Parameter(Mandatory = $true)][hashtable]$skillToolIdsById
    )

    Write-JsonFile -filePath $filePath -dataObject ([pscustomobject]@{ SkillToolIdsById = $skillToolIdsById })
}

$userSkillToolIdsById = Load-UserSkillToolIdsMap -filePath $userSkillToolIdsFilePath -knownToolIds $knownToolIds

function Get-EffectiveToolIdsForSkill {
    param([Parameter(Mandatory = $true)][string]$skillId)

    if ($userSkillToolIdsById.ContainsKey($skillId)) {
        return @($userSkillToolIdsById[$skillId])
    }

    $skillDefinition = Find-SkillDefinitionById -skillId $skillId
    if ($null -eq $skillDefinition) { return @() }

    $toolIdsValue = $skillDefinition.PSObject.Properties["ToolIds"].Value
    @(Normalize-ToolIdList -toolIdsValue $toolIdsValue -knownToolIds $knownToolIds)
}

function Set-UserToolIdsForSkill {
    param(
        [Parameter(Mandatory = $true)][string]$skillId,
        [Parameter(Mandatory = $true)][string[]]$toolIds
    )

    $uniqueMap = @{}
    $normalizedList = New-Object System.Collections.Generic.List[string]
    foreach ($toolId in $toolIds) {
        $text = [string]$toolId
        if ([string]::IsNullOrWhiteSpace($text)) { continue }
        if ($uniqueMap.ContainsKey($text)) { continue }
        $uniqueMap[$text] = $true
        $null = $normalizedList.Add($text)
    }

    $userSkillToolIdsById[$skillId] = @($normalizedList.ToArray())
    Save-UserSkillToolIdsMap -filePath $userSkillToolIdsFilePath -skillToolIdsById $userSkillToolIdsById
}

if (!(Test-Path -LiteralPath $mainWindowXamlFilePath)) {
    throw "MainWindow.xaml が見つかりません: $mainWindowXamlFilePath"
}

$xamlContent = Get-Content -LiteralPath $mainWindowXamlFilePath -Raw -Encoding UTF8
$stringReader = New-Object System.IO.StringReader($xamlContent)
$xmlTextReader = [System.Xml.XmlReader]::Create($stringReader)
$mainWindow = [Windows.Markup.XamlReader]::Load($xmlTextReader)

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

function Get-ToolIdSetForSelectedSkill {
    $selectedSkillId = Get-SelectedSkillId
    if ([string]::IsNullOrWhiteSpace($selectedSkillId)) { return @{} }

    $toolIdSet = @{}
    foreach ($toolIdentifier in @(Get-EffectiveToolIdsForSkill -skillId $selectedSkillId)) {
        $text = [string]$toolIdentifier
        if (![string]::IsNullOrWhiteSpace($text)) { $toolIdSet[$text] = $true }
    }
    $toolIdSet
}

function Refresh-SkillViewModels {
    $skillViewModels.Clear()

    foreach ($skillDefinition in @($skillsConfiguration.Skills)) {
        $skillViewModel = [pscustomobject]@{
            Id   = Get-PropertyString $skillDefinition "Id" ""
            Name = Get-PropertyString $skillDefinition "Name" ""
        }

        $null = $skillViewModel | Add-Member -MemberType ScriptMethod -Name ToString -Value { $this.Name } -Force
        $skillViewModels.Add($skillViewModel) | Out-Null
    }
}

function Get-FilteredToolDefinitions {
    $searchText = $applicationSearchTextBox.Text
    if ($null -eq $searchText) { $searchText = "" }
    $searchTextLower = $searchText.Trim().ToLowerInvariant()

    $result = @()
    foreach ($toolDefinition in @($toolsConfiguration.Tools)) {
        $toolName = Get-PropertyString $toolDefinition "Name" ""
        $toolPath = Get-PropertyString $toolDefinition "Path" ""

        if ($searchTextLower.Length -gt 0) {
            $matchName = $toolName.ToLowerInvariant().Contains($searchTextLower)
            $matchPath = $toolPath.ToLowerInvariant().Contains($searchTextLower)
            if (!($matchName -or $matchPath)) { continue }
        }

        $result += $toolDefinition
    }
    $result
}

function Refresh-ApplicationViewModels {
    $applicationViewModels.Clear()

    $toolIdSet = Get-ToolIdSetForSelectedSkill
    $selectedSkillId = Get-SelectedSkillId

    foreach ($toolDefinition in @(Get-FilteredToolDefinitions)) {
        $toolIdentifier = Get-PropertyString $toolDefinition "Id" ""

        $isAlreadyInSkill = $false
        if (![string]::IsNullOrWhiteSpace($toolIdentifier) -and $toolIdSet.ContainsKey($toolIdentifier)) {
            $isAlreadyInSkill = $true
        }

        $canAddToSelectedSkill = $true
        if ([string]::IsNullOrWhiteSpace($selectedSkillId)) { $canAddToSelectedSkill = $false }
        if ($isAlreadyInSkill) { $canAddToSelectedSkill = $false }

        $applicationViewModels.Add((New-ToolViewModel -toolDefinition $toolDefinition -canAddToSelectedSkill $canAddToSelectedSkill)) | Out-Null
    }

    $applicationCountTextBlock.Text = ("表示: {0} 件" -f $applicationViewModels.Count)
}

function Refresh-SkillApplicationViewModels {
    $skillApplicationViewModels.Clear()

    $selectedSkillId = Get-SelectedSkillId
    if ([string]::IsNullOrWhiteSpace($selectedSkillId)) {
        $skillApplicationCountTextBlock.Text = "表示: 0 件"
        $bulkLaunchButton.IsEnabled = $false
        return
    }

    foreach ($toolIdentifierValue in @(Get-EffectiveToolIdsForSkill -skillId $selectedSkillId)) {
        $toolIdentifier = [string]$toolIdentifierValue
        if ([string]::IsNullOrWhiteSpace($toolIdentifier)) { continue }
        if (!$toolDefinitionById.ContainsKey($toolIdentifier)) { continue }

        $toolDefinition = $toolDefinitionById[$toolIdentifier]
        $skillApplicationViewModels.Add((New-ToolViewModel -toolDefinition $toolDefinition -canAddToSelectedSkill $false)) | Out-Null
    }

    $skillApplicationCountTextBlock.Text = ("表示: {0} 件" -f $skillApplicationViewModels.Count)
    $bulkLaunchButton.IsEnabled = ($skillApplicationViewModels.Count -gt 0)
}

function Select-InitialSkill {
    $initialSkillId = Get-PropertyString $userSettings "SelectedSkillId" ""
    if ([string]::IsNullOrWhiteSpace($initialSkillId)) {
        if ($skillViewModels.Count -gt 0) { $initialSkillId = [string]$skillViewModels[0].Id }
    }

    $skillComboBox.SelectedValue = $initialSkillId
    if ($null -eq $skillComboBox.SelectedValue -and $skillViewModels.Count -gt 0) {
        $skillComboBox.SelectedIndex = 0
    }
}

function Add-ToolIdToSkill {
    param(
        [Parameter(Mandatory = $true)][string]$toolId,
        [Parameter(Mandatory = $true)][string]$skillId
    )

    $currentToolIds = @(Get-EffectiveToolIdsForSkill -skillId $skillId)
    if ($currentToolIds -contains $toolId) { return }

    Set-UserToolIdsForSkill -skillId $skillId -toolIds @($currentToolIds + @($toolId))
}

function Remove-ToolIdFromSkill {
    param(
        [Parameter(Mandatory = $true)][string]$toolId,
        [Parameter(Mandatory = $true)][string]$skillId
    )

    $currentToolIds = @(Get-EffectiveToolIdsForSkill -skillId $skillId)
    $updatedToolIds = @()
    foreach ($existingToolId in $currentToolIds) {
        if ([string]$existingToolId -ne $toolId) { $updatedToolIds += @([string]$existingToolId) }
    }
    Set-UserToolIdsForSkill -skillId $skillId -toolIds $updatedToolIds
}

Refresh-SkillViewModels
Select-InitialSkill
Refresh-SkillApplicationViewModels
Refresh-ApplicationViewModels

$skillComboBox.Add_SelectionChanged({
        $userSettings.SelectedSkillId = Get-SelectedSkillId
        Write-JsonFile -filePath $userSettingsFilePath -dataObject $userSettings
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

        try {
            $clickedButton = Find-ParentButton -sourceObject $eventArgs.OriginalSource
            if ($null -eq $clickedButton) { return }

            if ($clickedButton.Name -eq "openFromAllApplicationsNameButton") {
                $toolViewModel = $clickedButton.Tag
                if ($null -ne $toolViewModel) { Start-Tool -toolViewModel $toolViewModel }
            }
            elseif ($clickedButton.Name -eq "addToSkillButton") {
                $toolViewModel = $clickedButton.Tag
                if ($null -eq $toolViewModel) { return }

                $selectedSkillId = Get-SelectedSkillId
                if ([string]::IsNullOrWhiteSpace($selectedSkillId)) { return }

                Add-ToolIdToSkill -toolId ([string]$toolViewModel.Id) -skillId $selectedSkillId
                Refresh-SkillApplicationViewModels
                Refresh-ApplicationViewModels
            }
        }
        catch { }
    }
)

$skillApplicationListBox.AddHandler(
    [System.Windows.Controls.Primitives.ButtonBase]::ClickEvent,
    [System.Windows.RoutedEventHandler] {
        param($sender, $eventArgs)

        try {
            $clickedButton = Find-ParentButton -sourceObject $eventArgs.OriginalSource
            if ($null -eq $clickedButton) { return }

            if ($clickedButton.Name -eq "openFromSkillApplicationsNameButton") {
                $toolViewModel = $clickedButton.Tag
                if ($null -ne $toolViewModel) { Start-Tool -toolViewModel $toolViewModel }
            }
            elseif ($clickedButton.Name -eq "unlinkFromSkillButton") {
                $toolViewModel = $clickedButton.Tag
                if ($null -eq $toolViewModel) { return }

                $selectedSkillId = Get-SelectedSkillId
                if ([string]::IsNullOrWhiteSpace($selectedSkillId)) { return }

                Remove-ToolIdFromSkill -toolId ([string]$toolViewModel.Id) -skillId $selectedSkillId
                Refresh-SkillApplicationViewModels
                Refresh-ApplicationViewModels
            }
        }
        catch { }
    }
)

$bulkLaunchButton.Add_Click({
        try {
            if ($skillApplicationViewModels.Count -le 0) { return }

            foreach ($toolViewModel in @($skillApplicationViewModels)) {
                Start-Tool -toolViewModel $toolViewModel
                Start-Sleep -Milliseconds 150
            }
        }
        catch {
            Show-Message -message ("起動に失敗しました:`n{0}" -f $_.Exception.Message)
        }
    })

$null = $mainWindow.ShowDialog()