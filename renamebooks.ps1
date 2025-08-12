$booksDir = "c:\Users\conma\OneDrive\Downloads\Telegram Desktop"
$destDir = "c:\Users\conma\temp\"
$readMeta = "C:\Program Files\Calibre2\ebook-meta.exe"

$oldEnc = [System.Text.Encoding]::GetEncoding('utf-8')
$newEnc = [System.Text.Encoding]::GetEncoding('windows-1251')

$psi = New-Object System.Diagnostics.ProcessStartInfo
$psi.FileName = $readMeta
$psi.RedirectStandardOutput = $true
$psi.UseShellExecute = $false
$psi.StandardOutputEncoding = $oldEnc

foreach ($file in Get-ChildItem -Path $booksDir -Recurse -File | Select-Object -First 10) {
   # Start ebook-meta.exe to get file metadata
   $psi.Arguments = '"' + $file.FullName + '"'
   $proc = [System.Diagnostics.Process]::Start($psi)
   $oldEncStr = $proc.StandardOutput.ReadToEnd()
   $proc.WaitForExit()

   # Decode ebook-meta.exe output from utf-8 to windows-1251
   $oldEncBytes = $oldEnc.GetBytes($oldEncStr)
   $newEncBytes = [System.Text.Encoding]::Convert($oldEnc, $newEnc, $oldEncBytes)
   $newEncStr = $newEnc.GetString($newEncBytes)
   
   $newEncStr -split "`r`n" | ForEach-Object {
      if (-not [System.String]::IsNullOrWhiteSpace($_)) {
         $infoLine = $_ -split ":", 2
         switch ($infoLine[0].Trim()) {
            "Title" { $title = $infoLine[1].Trim() }
            "Author(s)" { $author = $infoLine[1].Trim() }
            "Published" { $year = $infoLine[1].Substring(1, 4) }
         }
      }
   }

   $string = 'copy "' + $file.FullName + '" "' + $destDir + $title + ' [' + $year + '] ' + $author + $file.Extension + '"'
   Write-Host "-Cmd-" $string
}
