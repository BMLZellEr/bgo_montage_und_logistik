# Load Json1
$json1Path = "plz.json"
$json1 = Get-Content $json1Path | ConvertFrom-Json

# Load Json2
$json2Path = "nos_plz_to_zone.json"
$json2 = Get-Content $json2Path | ConvertFrom-Json

# Convert Json2 to a hashtable for quick lookup
$zoneMap = @{}
foreach ($key in $json2.PSObject.Properties) {
    $zoneMap[$key.Name] = $key.Value
}

# Modify Json1 by adding "zone" where PLZ matches
foreach ($item in $json1.data) {
    $plz = $item.plz
    if ($zoneMap.ContainsKey($plz)) {
        $item | Add-Member -MemberType NoteProperty -Name "zone" -Value $zoneMap[$plz] -Force
    }
}

# Save updated Json1
$json1 | ConvertTo-Json -Depth 10 | Set-Content "json1_updated.json" -Encoding UTF8

Write-Host "JSON1 updated successfully! Check json1_updated.json"
