Add-Type -AssemblyName Microsoft.VisualBasic
try {
    $xl = [Microsoft.VisualBasic.Interaction]::GetObject($null, 'Excel.Application')
    Write-Output "Found Excel instance!"
    Write-Output "Workbooks count: $($xl.Workbooks.Count)"
    foreach ($wb in $xl.Workbooks) {
        Write-Output "Name: $($wb.Name) | Path: $($wb.FullName) | Saved: $($wb.Saved)"
    }
} catch {
    Write-Output "Error: $($_.Exception.Message)"
}
