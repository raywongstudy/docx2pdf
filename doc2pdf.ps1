# cf.: 
# - http://blog.coolorange.com/2012/04/20/export-word-to-pdf-using-powershell/
# - https://gallery.technet.microsoft.com/office/Script-to-convert-Word-f702844d
# http://blogs.technet.com/b/heyscriptingguy/archive/2013/03/24/weekend-scripter-convert-word-documents-to-pdf-files-with-powershell.aspx

param([string]$DocInput, [string]$PdfOutput = '.\output.pdf')

Add-type -AssemblyName Microsoft.Office.Interop.Word

# 
# - optimize for screen/print: wdExportOptimizeForOnScreen / wdExportOptimizeForPrint
# - content only/include markups & comments: wdExportDocumentContent / wdExportDocumentWithMarkup
# - create bookmarks: wdExportCreateHeadingBookmarks / wdExportCreateWordBookmarks

function WordToPdf([string]$wdSourceFile, [string]$wdExportFile) {
  $wdExportFormat = [Microsoft.Office.Interop.Word.WdExportFormat]::wdExportFormatPDF
  $wdOpenAfterExport = $false
  $wdExportOptimizeFor = [Microsoft.Office.Interop.Word.WdExportOptimizeFor]::wdExportOptimizeForOnScreen
  $wdExportRange = [Microsoft.Office.Interop.Word.WdExportRange]::wdExportAllDocument
  $wdStartPage = 0
  $wdEndPage = 0
  $wdExportItem = [Microsoft.Office.Interop.Word.WdExportItem]::wdExportDocumentContent
  $wdIncludeDocProps = $true
  $wdKeepIRM = $true
  $wdCreateBookmarks = [Microsoft.Office.Interop.Word.WdExportCreateBookmarks]::wdExportCreateHeadingBookmarks
  $wdDocStructureTags = $true
  $wdBitmapMissingFonts = $true
  $wdUseISO19005_1 = $false

  $wdApplication = $null;
  $wdDocument = $null;
  
  # How to: Programmatically Close Documents (without changes)
  # http://msdn.microsoft.com/en-us/library/af6z0wa2.aspx
  $doNotSaveChanges = [Microsoft.Office.Interop.Word.WdSaveOptions]::wdDoNotSaveChanges

  try
  {
         $wdApplication = New-Object -ComObject "Word.Application"
         $wdDocument = $wdApplication.Documents.Open($wdSourceFile)
         $wdDocument.ExportAsFixedFormat(
         $wdExportFile,
         $wdExportFormat,
         $wdOpenAfterExport,
         $wdExportOptimizeFor,
         $wdExportRange,
         $wdStartPage,
         $wdEndPage,
         $wdExportItem,
         $wdIncludeDocProps,
         $wdKeepIRM,
         $wdCreateBookmarks,
         $wdDocStructureTags,
         $wdBitmapMissingFonts,
         $wdUseISO19005_1
         )
  }
  catch
  {
         $wshShell = New-Object -ComObject WScript.Shell
         $wshShell.Popup($_.Exception.ToString(), 0, "Error", 0)
         $wshShell = $null
  }
  finally
  {
         if ($wdDocument)
         {
                $wdDocument.Close([ref]$doNotSaveChanges)
                $wdDocument = $null
         }
         if ($wdApplication)
         {
                $wdApplication.Quit()
                $wdApplication = $null
         }
         [GC]::Collect()
         [GC]::WaitForPendingFinalizers()
  }
}

$FullInput = (Get-Item ${DocInput}).FullName
# http://stackoverflow.com/questions/3038337/powershell-resolve-path-that-might-not-exist
$FullOutput = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath(${PdfOutput})

Write-Host "Converting ${FullInput} to ${FullOutput}..."
WordToPdf ${FullInput} ${FullOutput} 
