param(
    [string]$OutFile,
    [system.Array]$lines
)
$Word = New-Object -ComObject Word.Application
$Doc = $Word.Documents.Add()
$Word.Visible = $True
$Selection = $Word.Selection
forEach ($Line in $Lines) {
    if ($Line -match "<strong>") {
        $Selection.Font.Bold = 1
        continue
    }
    if ($Line -match "</strong>") {
        $Selection.Font.Bold = 0
        continue
    }
    $Selection.TypeText($Line)
    $Selection.TypeParagraph()
}
$Doc.SaveAs($OutFile)
$Doc.Close()
$Word.Quit()
$Null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$Word)
[gc]::Collect()
[gc]::WaitForPendingFinalizers()
Remove-Variable Word 