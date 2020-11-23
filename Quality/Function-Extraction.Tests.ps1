param(
    [parameter(Mandatory = $true)]
    [string[]]$SourcePath
)

Describe "Function Extraction Tests" -Tag 'Setup' {

    $FunctionExtractPath = Join-Path -Path $Env:TEMP -ChildPath "tmpExtraction"
    Get-ChildItem -Path $FunctionExtractPath -Recurse | Remove-Item -Force -Recurse
    Remove-Item $FunctionExtractPath -Force -ErrorAction SilentlyContinue
    New-Item -Path $FunctionExtractPath -ItemType 'Directory'

    foreach ($scriptPath in $SourcePath) {

        $moduleTestFiles = Get-ModuleFile -Path $scriptPath

        foreach ($moduleFile in $moduleTestFiles) {

            Context "Module : $moduleFile" {

                It "function extraction should complete" -TestCases @{ 'moduleFile' = $moduleFile } {

                    {

                        Export-FunctionsFromModule -Path $moduleFile

                    } | Should -Not -Throw

                }

            }

        }
    }

}
