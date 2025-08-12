# Source folder: where to get books
$booksDir = "c:\Users\conma\OneDrive\Downloads\Telegram Desktop"
# Destination folder: where to copy renamed books
$destDir = "d:\---temp\"
# Full path to ebook-meta.exe file from Calibre
$readMeta = "C:\Program Files\Calibre2\ebook-meta.exe"
# The name of the folder where to copy books which we can not rename
$notRenamedFolder = "not-renamed"

if (-not (Test-Path -Path $booksDir)) {
   Write-Host "Books folder is not found:" $booksDir
   exit
}

if (-not (Test-Path -Path $readMeta)) {
   Write-Host "Calibre utility is not found:" $readMeta
   exit
}

if (-not (Test-Path -Path $destDir)) {
   Write-Host "Destination folder is not found:" $destDir
   exit
}
if (-not (Test-Path -Path $($destDir + $notRenamedFolder))) {
   New-Item -ItemType Directory -Path $($destDir + $notRenamedFolder) | Out-Null
}

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
   
   # Split multiline string into separate lines
   $newEncStr -split "`r`n" | ForEach-Object {
      if (-not [System.String]::IsNullOrWhiteSpace($_)) {
         # Split the line by the colon into a title and a value.
         $infoLine = $_ -split ":", 2
         switch ($infoLine[0].Trim()) {
            "Title" { $title = $infoLine[1].Trim() }
            "Author(s)" { $author = $infoLine[1].Trim() }
            "Published" { $year = $infoLine[1].Substring(1, 4) }
         }
      }
   }

   $newFileName = ""
   # Prepare new file name. Check data a little
   if ([System.String]::IsNullOrWhiteSpace($title) -or $title -match '^\d+$') {
      $newFileName = $destDir + $notRenamedFolder + "\" + $file.BaseName
   }
   else {
      $newFileName = $title
      if (-not [System.String]::IsNullOrWhiteSpace($year)) {
         $newFileName = $newFileName + " [" + $year + "]"
      }
      if (-not [System.String]::IsNullOrWhiteSpace($author)) {
         $newFileName = $newFileName + " " + $author
      }
      $newFileName = $destDir + $newFileName
   }

   $i = 0
   $testExistance = $newFileName
   while ([System.IO.File]::Exists($testExistance + $file.Extension)) {
      $i++
      $testExistance = $newFileName + "(" + $i + ")"
   }
   $newFileName = $testExistance + $file.Extension

   Write-Host $file.FullName " -> " $newFileName

   Copy-Item $file.FullName -Destination $newFileName
}
