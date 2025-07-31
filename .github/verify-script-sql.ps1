param (
    [string]$pathCheckScript
)

$nameList = @("CVGDEVDB", "CVGSITDB", "CVGUATDB")

Write-Host "🔍 Scanning SQL scripts in: $pathCheckScript"

Get-ChildItem -Path $pathCheckScript -Filter *.sql -Recurse -File | ForEach-Object {
    $filePath = $_.FullName
    Write-Host "📄 Found SQL: '$filePath'"
    $lines = Get-Content $filePath
    $errors = @()
    $hasCreateTable = $false
    $hasAdminRole = $false
    $hasQueryRole = $false

    # --- Condition 1: Found keyword (optional reporting, still printed immediately) ---
    foreach ($name in $nameList) {
        $trimmedName = $name.Trim()
        Write-Host "🔍 Checking prefix database for: '$trimmedName'"
        $escapedName = [regex]::Escape($trimmedName)

        for ($i = 0; $i -lt $lines.Length; $i++) {
            $line = $lines[$i]
            if ($line -match "(?i)$escapedName") {
                Write-Host "❌ ERROR: Found database prefix '$trimmedName' in file: $filePath (Line: $($i + 1))"
                Write-Host "    → $line"
                throw "Database prefix '$trimmedName' found in $filePath at line $($i + 1)"
            }
        }
    }

# --- Condition 2: Incorrect Schema (Perfected!) ---
for ($i = 0; $i -lt $lines.Length; $i++) {
    $line = $lines[$i]

    # Match object name: [CVG_CFG_SOMETHING] or CVG_TBL_XXX
    $pattern = "(?i)(\[\s*CVG_(CFG|TBL)[\w]*\s*\]|\bCVG_(CFG|TBL)[\w]*)"

    $matches = [regex]::Matches($line, $pattern)
    foreach ($match in $matches) {
        $matchText = $match.Value
        $matchIndex = $match.Index

        # Get text before the matched keyword
        $prefix = $line.Substring(0, $matchIndex)

        # If prefix does NOT end with something like [cvgadm]. or cvgadm.
        if ($prefix -notmatch "(?i)(\[?\s*cvgadm\s*\]?\s*\.\s*)$") {
            $errors += "[{0}] ❌ Incorrect Schema (Line {1}): {2}" -f ($errors.Count + 1), ($i + 1), $line.Trim()
        }
    }
}



    # --- Condition 3: Forgot GRANT Authorize ---
    $fileText = [System.String]::Join("`n", $lines)

    if ($fileText -match "(?i)CREATE\s+TABLE") {
        $hasCreateTable = $true
    }
    if ($fileText -match "(?i)cvg_admin_role") {
        $hasAdminRole = $true
    }
    if ($fileText -match "(?i)cvg_query_role") {
        $hasQueryRole = $true
    }

    if ($hasCreateTable -and (-not $hasAdminRole -or -not $hasQueryRole)) {
        $errors += "[{0}] ⚠️ Forgot GRANT Authorize" -f ($errors.Count + 1)
    }

    # --- Print error summary per script ---
    if ($errors.Count -gt 0) {
        Write-Host ""
        Write-Host "❗ SCRIPT ERROR: $filePath"
        $errors | ForEach-Object { Write-Host $_ }
        Write-Host "`n"
    }
}