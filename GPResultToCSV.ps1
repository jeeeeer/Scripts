<#
Author: Bradley Ventura 2022

Supported .html inputs:
    - gpresult
    - RSOP
    - GPO Report export
#>

$ImportPath = "C:\Import\import.html"
$ExportPath = "C:\Export\export.csv"

$SettingNameRegEx = 'gpmc_settingName=".*" gpmc_settingPath'
$SettingPathRegEx = 'gpmc_settingPath=".*" gpmc_settingDescription'
$ValueRegEx = '</span></td><td>.*</td><td>'
$GPNameRegEx = '</td><td>.*</td></tr>'
$SubSettingsRegex = 'class="subtable_frame"'
$SubSubSettingsRegex = 'class="subtable"'

$FinalList = @()

$Content = Get-Content -Path $ImportPath

Function Strip-characters {
    [CmdletBinding()]
    Param([string]$String)

    $WorkingString = $String

    $Hash = [ordered]@{
        "&quot;" = '"'
        "&amp"   = '&'
        "&nbsp;" = ' '
        "&lt"    = '<'
        "&gt;"   = '>'
        "&#39;"  = "'"
    }        

    Foreach($Character in $Hash.GetEnumerator()){
        if($string -like $Character.Name) { $WorkingString = $WorkingString.Replace($Character.Name,$Character.Value) }
    }

    return $WorkingString
}

For($i=0;$i -lt ($Content.Count); $i++){
    $GPSettingName,$GPValue,$GPSettingPath,$GPName,$SubSettingsName,$SubSettingsValue,$SubSettingsList = $null

    if($Content[$i] -match $SettingNameRegex){
        $GPSettingName = (($Matches.Values).Split('"'))[1]
        if($Content[$i] -match $SettingPathRegex) { $GPSettingPath = (($Matches.Values).Split('"'))[1] }
        if($Content[$i] -match $ValueRegEx) { $GPValue = (((($matches.values).Split('>'))[3]).Split('<'))[0] }
        if($Content[$i] -match $GPNameRegex){ $GPName = (((($Matches.Values).Split('>'))[4]).Split('<'))[0] }
        if($Content[$i+1] -match $SubSettingsRegex){
            if($Content[$i+2] -match $SubSubSettingsRegex){
                $SubSettingsName = ($Content[$i+3] -split {$_ -eq '<' -or $_ -eq '>'})[4]
                $i = $i+4
                [String[]]$SubSettingsList
                While($Content[$i] -notlike '</table>*'){
                    $SubSettingsList = $SubSettingsList + (, [String](($Content[$i] -split {$_ -eq '<' -or $_ -eq '>'})[4]))
                    $i++
                }
                $SubSettingsValue = $SubSettingsList -join "`n"
            }
            else {
                $SubSettingsName = (($Content[$i+2] -split {$_ -eq '<' -or $_ -eq '>'})[4])
                $SubSettingsValue = (($Content[$i+2] -split {$_ -eq '<' -or $_ -eq '>'})[8])
            }
        }
    $object = New-Object PSObject -Property @{
        GPSettingName   = (Strip-Characters $GPSettingName)
        GPValue         = (Strip-Characters $GPValue)
        GPPath          = (Strip-Characters $GPSettingPath)
        GPName          = (Strip-Characters $GPName)
        SubSettingName  = (Strip-Characters $SubSettingsName)
        SubSettingValue = (Strip-Characters $SubSettingsValue)
        }
 
    $FinalList+= $object
    }
}
 
$FinalList | Select GPPath,GPSettingName,GPValue,SubSettingName,SubSettingValue,GPName | Export-Csv $ExportPath -NoTypeInformation
