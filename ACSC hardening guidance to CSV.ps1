# Quick and dirty just like ur mum LOL

$URI = "https://www.cyber.gov.au/acsc/view-all-content/publications/hardening-microsoft-windows-10-version-21h1-workstations"
# "https://www.cyber.gov.au/acsc/view-all-content/publications/hardening-microsoft-365-office-2021-office-2019-and-office-2016"

$URINameSplit = ($URI -split "/")
$TaskName = $URINameSplit[$URINameSplit.Length-1]

$CategoryRegex = "<h3>.*</h3>"
$TableRegex    = '<div class="table-responsive"><table>'
$GPTable = "<p>Group Policy Setting[s]*</p>"
$RegistryTable = "<p>Registry Entr.*</p>"


$rawhtml = (Invoke-WebRequest $URI -UseBasicParsing).RawContent


$rawhtmlarray = $rawhtml.split("`n")

$ObjList = @()

For($i=0;$i -lt $rawhtmlarray.Count; $i++){
    if($rawhtmlarray[$i] -match $CategoryRegex) { $Category =  ((($Matches[0].Split('>'))[1]).Split('<')[0]).Trim() ; Continue }
    if( ($rawhtmlarray[$i] -like "<table>") -and ($rawhtmlarray[$i+4] -match $GPTable) ) {
        $TableType = "Group Policy"
        while($rawhtmlarray[$i] -notlike "*</table>*") {
            if(($rawhtmlarray[$i] -like '<td class="align-top" colspan="2">') -and ($rawhtmlarray[$i+1] -match '<p><strong>.*</strong></p>')){
                $Location = ($rawhtmlarray[$i+1].Split('>')[2]).Split('<')[0]
                $i = $i + 2; Continue
            }
            if($rawhtmlarray[$i] -like '<td class="align-top">') {
                $SubSettingsList = @()
                $SettingName = ((($rawhtmlarray[$i+1].Split('>'))[1]).Split('<'))[0]
                $Setting = $((($rawhtmlarray[$i+4].Split('>'))[1]).Split('<'))[0]
                $i = $i+5
                while($rawhtmlarray[$i] -notlike '</tr>') {
                    if($rawhtmlarray[$i] -match '<p>.*</p>') {
                        $SubSettingsList = $SubSettingsList + (, ((($rawhtmlarray[$i].Split('>'))[1]).Split('<'))[0])
                    }
                    $i++
                }
                if($SubSettingsList[0] -like "*:*") {
                    $SubSettingSplit = $SubSettingsList[0] -split "(:)"
                    $SubSettingName = -join $SubSettingSplit[0]
                    $SubSettingsList[0] = (-join $SubSettingSplit[2..$SubSettingSplit.Count]).Trim()
                }
                $SubSettings = $SubSettingsList -join "`n"
                $obj = New-Object PSObject -Property @{
                    Location = $Location
                    SettingName = $SettingName
                    Setting = $Setting
                    SubSettingName = $SubSettingName
                    SubSettings = $SubSettings
                    SettingType = $TableType
                }
                $SettingName,$Setting,$SubSettingName,$SubSettings = $null
                $ObjList += $obj
            }
            $i++
        }
    }
    
    if( ($rawhtmlarray[$i] -like "<table>") -and ($rawhtmlarray[$i+4] -match $RegistryTable) ) {
        $TableType = "Registry"
        while($rawhtmlarray[$i] -notlike "*</table>*") {
            if(($rawhtmlarray[$i] -like '<td class="align-top" colspan="2">') -and ($rawhtmlarray[$i+1] -match '<p><strong>.*</strong></p>')){
                $Location = ($rawhtmlarray[$i+1].Split('>')[2]).Split('<')[0]
                $i = $i + 2; Continue
            }
            if($rawhtmlarray[$i] -like '<td class="align-top">') {
                $SubSettingsList = @()
                $SettingName = ((($rawhtmlarray[$i+1].Split('>'))[1]).Split('<'))[0]
                $Setting = $((($rawhtmlarray[$i+4].Split('>'))[1]).Split('<'))[0]
                $i = $i+5
                
                $obj = New-Object PSObject -Property @{
                    Location = $Location
                    SettingName = $SettingName
                    Setting = $Setting
                    SubSettingName = $SubSettingName
                    SubSettings = $SubSettings
                    SettingType = $TableType
                }
                $SettingName,$Setting,$SubSettingName,$SubSettings = $null
                $ObjList += $obj
            }
            $i++
        }
    }

}

$ObjList | Select Location,SettingName,Setting,SubSettingName,SubSettings,SettingType | Export-Csv -path ("C:\Temp\" + $TaskName + ".csv") -NoTypeInformation -Encoding utf8