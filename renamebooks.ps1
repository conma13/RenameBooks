$booksDir = "c:\\Users\\conma\\OneDrive\\Downloads\\Telegram Desktop"
$calibreDir = "C:\\Program Files\\Calibre2"
$calibreCommand = "ebook-meta.exe"
$readMeta = $calibreDir + "\\" + $calibreCommand
$oldEnc = [System.Text.Encoding]::GetEncoding('cp866')
$newEnc = [System.Text.Encoding]::GetEncoding('windows-1251')
# $oldEnc = [System.Text.Encoding]::UTF8

foreach ($file in Get-ChildItem -Path $booksDir -Recurse -File | Select-Object -First 5) {
   $oldEncStr = &$readMeta $file.FullName

   Write-Host "Old Encoding" $oldEncStr

   #$bookInfoBytes = [System.Text.Encoding]::GetEncoding('cp866').GetBytes($bookInfoStr866)
   $oldEncBytes = $oldEnc.GetBytes($oldEncStr)
   $newEncBytes = [System.Text.Encoding]::Convert($oldEnc, $newEnc, $oldEncBytes)
   $newEncChars = $newEnc.GetChars($newEncBytes)
   $newEncStr = -join $newEncChars

   Write-Host "New Encoding" $newEncStr
   Write-Host "----------------------------------------------------"

   $bookInfo = &$readMeta $file.FullName | ConvertFrom-String -Delimiter ':' -PropertyNames Key, Value
   $bookInfo

   # Write-Host "----------------------------------------------------"
   # foreach ($infoLine in $bookInfo) {
   #    switch ($infoLine.Key.Trim()) {
   #       "Title" { $title = $infoLine.Value.Trim() }
   #       "Author(s)" { $author = $infoLine.Value.Trim() }
   #       "Published" { $year = $infoLine.Value.Substring(1, 4) }
   #    }
   # }
   # $string = $title + " [" + $year + "] " + $author
   # Write-Host $string
   Write-Host "-------------------endbook---------------------------------"
}
