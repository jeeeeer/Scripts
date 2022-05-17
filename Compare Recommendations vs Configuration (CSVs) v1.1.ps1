<#
Author: Bradley Ventura April-2022
 
Configuration CSV Format
GPPath    GPSettingName    GPValue    SubSettingName    SubSettingValue    GPName

Recommendations CSV Format
GPPath    GPSettingName    GPValue    SubSetting
#>
 
# Recommendations Folder
$RecommendationsPath = "C:\Temp\CSV\Recommendations\"
# Configuration CSV
$ConfigCSVPath = "C:\Temp\CSV"
# Output Folder
$OutputFolder = "C:\Temp\Output"
 
$ConfigCSV = Import-Csv -Path $ConfigCSVPath
$RecommendationsFileNames = (Get-ChildItem *.csv).BaseName
 
Foreach($RecommendationFileName in $RecommendationsFileNames){
    $Unconfigured = @()
    $Aligned      = @()
    $Misaligned   = @()
   
    $RecommCSV = Import-Csv -Path ($RecommendationsPath + $RecommendationFileName + ".csv")
   
    Foreach($Recommendation in $RecommCSV){
        # change formatting to be consistent with GPRESULT/RSOPs
   
        #$Match = $ConfigCSV | ? { ($_.GPSettingName -like $Recommendation.GPSettingName) -and ($_.GPPath -like $Recommendation.GPPath) }
        $RecommendedSubSettingName = (($Recommendation.SubSetting) -split ": ",2)[0]
        $RecommendedSubSettingValue = (($Recommendation.SubSetting) -split ": ",2)[1]
 
        $GPPath = $Recommendation.GPPath
        $GPPath = $GPPath.Replace('\Policies\','\'); $GPPath = $GPPath.Replace('\','/'); $GPPath = $GPPath.Replace('/ ','/'); $GPPath = $GPPath.Replace(' /','/'); $GPPath = $GPPath.Trim()
        if($GPPath[($GPPath.length-1)] -eq '/') { $GPPath = $GPPath.Substring(0,$GPPath.Length-1) }
 
        $Match = $ConfigCSV | ? { (($_.GPSettingName.Trim()) -like ($Recommendation.GPSettingName.Trim())) -and (($_.GPPath.Trim()) -like $GPPath) }
        #Write-Host $GPPath
        if($Match -eq $null) {
            $Unconfigured += $Recommendation
        }
        elseif($Match.Count -gt 1) {
            Write-Host "EXCEPTION - More than one match found"
        }
        elseif($Recommendation.GPValue -like $Match.GPValue) {
            $MatchObj = New-Object PSObject -Property @{
                GPPath = $Match.GPPath
                GPSettingName = $Match.GPSettingName
                GPValue = $Match.GPValue
                SubSettingName = $Match.SubSettingName
                ConfiguredSubsetting = $Match.SubSettingValue
                RecommendedSubSetting = $RecommendedSubSettingValue
                GPOName = $Match.GPName
            }
            $Aligned += $MatchObj
        }
        else{
            $MismatchObj = New-Object PSObject -Property @{
                GPPath = $Match.GPPath
                GPSettingName = $Match.GPSettingName
                ConfiguredValue = $Match.GPValue
                RecommendedValue = $Recommendation.GPValue
                SubSettingName = $Match.SubSettingName
                ConfiguredSubsetting = $Match.SubSettingValue
                RecommendedSubSetting = $RecommendedSubSettingValue
                GPOName = $Match.GPName
            }
            $Misaligned += $MismatchObj
        }
    }
 
    $Unconfigured | Select GPPath,GPSettingName,GPValue,SubSetting | Export-Csv -Path ($OutputFolder + $RecommendationFileName + " - Unconfigured.csv") -NoTypeInformation
    $Misaligned | Select GPPath,GPSettingName,ConfiguredValue,RecommendedValue,SubSettingName,ConfiguredSubsetting,RecommendedSubSetting,GPOName | Export-Csv -Path ($OutputFolder + $RecommendationFileName + " - Misaligned.csv") -NoTypeInformation
    $Aligned | Select GPPath,GPSettingName,GPValue,SubSettingName,ConfiguredSubsetting,RecommendedSubSetting,GPOName | Export-Csv -Path ($OutputFolder + $RecommendationFileName + " - Aligned.csv") -NoTypeInformation
}