# Define the path to the Excel file
$filePath = "C:\Centris\Testing\Test_umgebung\CEN_UNI_WRT_SYR_CHK_Test.xls"

# Create a new Excel application COM object
$excel = New-Object -ComObject Excel.Application

# Open the workbook
$workbook = $excel.Workbooks.Open($filePath)

# Iterate over each VBA project
foreach ($vbProject in $workbook.VBProject.VBComponents) {
    # Check if the component is a macro
    if ($vbProject.Type -eq 3) {
        Write-Output "Macro found in $($workbook.Name): $($vbProject.Name)"
    }else {
        Write-Output "Macro wasnt found"
    }
}

# Close the workbook and quit Excel
$workbook.Close()
$excel.Quit()

# Release the COM objects
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()