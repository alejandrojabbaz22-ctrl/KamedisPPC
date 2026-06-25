$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$path = Join-Path (Get-Location) 'AI STUDIO REPORT.xlsx'
$wb = $excel.Workbooks.Open($path)
$ws = $wb.Sheets.Item('OVERVIEW')
try {
    for ($r = 1; $r -le 20; $r++) {
        $a = $ws.Cells.Item($r,1).Value2
        $b = $ws.Cells.Item($r,2).Value2
        $c = $ws.Cells.Item($r,3).Value2
        if ($null -ne $a -or $null -ne $b -or $null -ne $c) {
            Write-Output ('Row ' + $r + ': A=' + ($a -as [string]) + ' | B=' + ($b -as [string]) + ' | C=' + ($c -as [string]))
        }
    }
} finally {
    $wb.Close($false)
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ws) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}
