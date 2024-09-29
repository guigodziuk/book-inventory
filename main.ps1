#Comic Book Automizer!!!!

### Functions ###

Function DCLogo {
    Write-Host "                                                  " -ForegroundColor White -BackgroundColor Blue
    Write-Host "                                                  " -ForegroundColor White -BackgroundColor Blue
    Write-Host "               ..:^!7???JJ??77!^:..               " -ForegroundColor White -BackgroundColor Blue
    Write-Host "            :^!J5GB#&&&@@@@&&&#BG5?!^.            " -ForegroundColor White -BackgroundColor Blue
    Write-Host "         .^7YB&@&@@&&&&&&&&&&&&@@@&#GY7^.         " -ForegroundColor White -BackgroundColor Blue
    Write-Host "       .~?P&5J5!B@&&&&@@@@@@&&&&&&&&@@#P?~.       " -ForegroundColor White -BackgroundColor Blue
    Write-Host "      ~7P&@B7  .Y#@@&#BGGGGB#&@@@&&&&#@@B57^      " -ForegroundColor White -BackgroundColor Blue
    Write-Host "    .!J#@&&#BP^BBG5?~^:. .^~~!?5B&@@&!7!!@#?!.    " -ForegroundColor White -BackgroundColor Blue
    Write-Host "   .!J&&&&&&&@#G7:     .!5##BG?::7G#5~  !5#&J!.   " -ForegroundColor White -BackgroundColor Blue
    Write-Host "   !?&&&&&&&@B?:      !G&@&B&@#?. :?B@5J&&&&&?!   " -ForegroundColor White -BackgroundColor Blue
    Write-Host "  ^!#@&&&&&@B!.  .^^::7#@@G7??^:~?..7#@&&&&&@B!^  " -ForegroundColor White -BackgroundColor Blue
    Write-Host "  !?@&&&&&&&?. ^?G&&#BJ7G&@#J:~P&&G!:?&&&&&&&@?!  " -ForegroundColor White -BackgroundColor Blue
    Write-Host "  75@&&&&&@B!!P#@@#B&@&P7J#@@GJG&&@5.!#&&&&&&@J!  " -ForegroundColor White -BackgroundColor Blue
    Write-Host "  75@&&&&&@B!^Y&@&B??B@@#J7P&@@@@&G! !#&&&&&&@J!  " -ForegroundColor White -BackgroundColor Blue
    Write-Host "  !?@&&&&&&&?. !P&@#Y!Y#@&G7?55PY~. :?&&&&&&&@?!  " -ForegroundColor White -BackgroundColor Blue
    Write-Host "  ^!#@&&&&&@B!  .?#@@BJP&@@Y       .7B@&&&&&@B!^  " -ForegroundColor White -BackgroundColor Blue
    Write-Host "   !?&&&&#&@&G?:  ^5&@&@@#5~      :?#@&&&&&&&?!   " -ForegroundColor White -BackgroundColor Blue
    Write-Host "   .!Y&&@5~J^G@P7:  7B#P?^      :7G&@&&&&&@&J!.   " -ForegroundColor White -BackgroundColor Blue
    Write-Host "    .!J#B?. :?G@&B57~!!:  .:^~?5BPPBYG@&&@#J!.    " -ForegroundColor White -BackgroundColor Blue
    Write-Host "      ~75##!B&&&&@@@&#BGGGGB#&@@&J   5#@&P7^      " -ForegroundColor White -BackgroundColor Blue
    Write-Host "       .~?G#@@&&&&&&&&@@@@@@&&&&GPY.YPGP?~.       " -ForegroundColor White -BackgroundColor Blue
    Write-Host "         .^7YG#&@@@&&&&&&&&&&&&@@@@PG57^.         " -ForegroundColor White -BackgroundColor Blue
    Write-Host "            :^!J5GB#&&&@@@@&&&#BG5J7^:            " -ForegroundColor White -BackgroundColor Blue
    Write-Host "               ..:~!7??JJJJ??7!^:..               " -ForegroundColor White -BackgroundColor Blue
    Write-Host "                        ..                        " -ForegroundColor White -BackgroundColor Blue
    Write-Host "                                                  " -ForegroundColor White -BackgroundColor Blue
    Write-Host "                                                  " -ForegroundColor White -BackgroundColor Blue
}    
Function Header {
    $excelFile.Columns.Item(1).ColumnWidth = 30
    $excelFile.Columns.Item(2).ColumnWidth = 55
    $excelFile.Columns.Item(3).ColumnWidth = 15
    $excelFile.Columns.Item(4).ColumnWidth = 10
    $excelFile.Columns.Item(5).ColumnWidth = 12
    $excelFile.Columns.Item(6).ColumnWidth = 8
    $excelFile.Columns.Item(7).ColumnWidth = 15
    $excelFile.Columns.Item(8).ColumnWidth = 10
    $excelFile.Columns.Item(9).ColumnWidth = 2
    $excelFile.Columns.Item(10).ColumnWidth = 8
}
Function InsertRow {
    param(
        [int]$initialRange,
        [int]$finalRange
        )
    if (!$finalRange) {
        $finalRange = $initialRange
    }
    $rangeToInsert = $excelFile.Range(("A{0}" -f $initialRange),("F{0}" -f $finalRange))
    $rangeToInsert.Insert([System.Type]::Missing)
}
Function RemoveRow {
    param(
        [int]$initialRange,
        [int]$finalRange
        )
    if (!$finalRange) {
        $finalRange = $initialRange
    }
    $range0 = $excelFile.Range(("A{0}" -f $initialRange),("F{0}" -f $finalRange))
    $range0.Select()
    $range0.Delete([System.Type]::Missing)
}
Function RegisterFolder {
    param(
        [int]$row,
        [string]$folder
        )
    $excelFile.Cells.Item($row,1) = $folder
    $excelFile.Cells.Item($row,1).Font.Bold=$true
    $excelFile.Cells.Item($row,5) = "Arc Rating:"
    $excelFile.Cells.Item($row,5).Font.Bold=$true
}
Function RegisterComic {
    param(
        [int]$row,
        [string]$subfolder
    )
    $excelFile.Cells.Item($row,2) = $subfolder
    $excelFile.Cells.Item($row,2).Font.Bold=$false
}
Function ExplodeRange {
    param(
        [int]$initialRange,
        [int]$finalRange
    )
    $range1 = $excelFile.Range(("A{0}" -f $initialRange),("A{0}" -f $finalRange))
        $range1.MergeCells = $false
    $range2 = $excelFile.Range(("E{0}" -f $initialRange),("E{0}" -f $finalRange))
        $range2.MergeCells = $false
    $range3 = $excelFile.Range(("F{0}" -f $initialRange),("F{0}" -f $finalRange))
        $range3.MergeCells = $false
}  
Function TableStyle {
    param(
        [int]$initialRange,
        [int]$finalRange
    )
    $range0 = $excelFile.Range(("A{0}" -f $initialRange),("F{0}" -f $finalRange))
        $range0.Select()
        $range0.Borders.LineStyle = 1
        $range0.Borders.Weight = 2
        $range0.HorizontalAlignment = -4108
        $range0.VerticalAlignment = -4108
        $range0.WrapText = $true
        $range0.Interior.Color = 15917529
    $range1 = $excelFile.Range(("A{0}" -f $initialRange),("A{0}" -f $finalRange))
        $range1.MergeCells = $true
        $range1.Borders.LineStyle = 1
        $range1.Borders.Weight = 3
        $range1.Interior.Color = 15917529
    $range2 = $excelFile.Range(("E{0}" -f $initialRange),("E{0}" -f $finalRange))
        $range2.MergeCells = $true
        $range2.Interior.Color = 15917529
    $range3 = $excelFile.Range(("F{0}" -f $initialRange),("F{0}" -f $finalRange))
        $range3.MergeCells = $true
        $range3.Borders.Item(2).Weight = 3
        $range3.Interior.Color = 15917529
    $range4 = $excelFile.Range(("A{0}" -f $finalRange),("F{0}" -f $finalRange))
        $range4.Borders.Item(4).Weight = 3
    $range5 = $excelFile.Range(("D{0}" -f $initialRange),("D{0}" -f $finalRange))
        $range5.Borders.Item(2).Weight = 3
}
Function RemoveFolder {               
    Write-Host "ENTRANDO NO REMOVEFOLDER" -ForegroundColor White -BackgroundColor Black
    $lineComicsRegistered = 3
    while ($null -ne $excelFile.Cells.Item($lineComicsRegistered,2).Value2) {
        $lineComicsRegistered++
    }
    while ($excelRow -lt $lineComicsRegistered) {
        Write-Host "COMICS REGISTERED $lineComicsRegistered" -ForegroundColor White -BackgroundColor Black
        Write-Host "LINE $excelRow" -ForegroundColor White -BackgroundColor Black
        Write-Host "ENTRANDO NO while do REMOVEFOLDER" -ForegroundColor White -BackgroundColor Black        
        $remNav = $excelFile.Cells.Item($excelRow,1).Value2
        $regArc = CountRegisteredArc -line $excelRow -testNav $testNav
        $countCoomicsToJump = $regArc.count
        $directoryChildPathTest = Join-Path -Path $directoryPath -ChildPath $remNav
        if (Test-Path -Path $directoryChildPathTest -PathType Container) {
            Write-Host "$remNav exists." -BackgroundColor DarkYellow -ForegroundColor White                               
            $excelRow = $excelRow + $countCoomicsToJump
        } else {                
            $remNav = $excelFile.Cells.Item($excelRow,1).Value2
            Write-Host "$remNav does not exists." -BackgroundColor Magenta
            Write-Host "$remNav doesn't exists..." -BackgroundColor DarkRed
            $y = $excelRow + $countCoomicsToJump - 1
            RemoveRow -initialRange $excelRow -finalRange $y
            $lineComicsRegistered = $lineComicsRegistered - $countCoomicsToJump
            Write-Host "$remNav deleted." -BackgroundColor DarkRed
        }
    }
    $excelRow = 3
}
Function CountRegisteredArc() {
    param(
        [int]$line,
        [string]$testNav
        )
    $registeredComics = @()
    $testNav = $excelFile.Cells.Item($line,1).Value2
    $regcom = $testNav
    $boolean = $true
    while ($boolean) {
        $regcom = $excelFile.Cells.Item($line,1).Value2
        $regbook = $excelFile.Cells.Item($line,2).Value2
        if ($regcom -ne "$testNav" -and  $null -ne $regcom) {
            $boolean = $false
            $line++
        }
        elseif ($null -eq $regcom -and $null -eq $regbook) {
            $boolean = $false
        }

        if ($boolean) {
            $registeredComics += $excelFile.Cells.Item($line,2).Value2
            $line++
        }
    }
    return $registeredComics
}

### Main ###

Clear-Host

if (!$excel) {
    $excel = New-Object -ComObject Excel.Application
}
$initialCheckLine = 3   
$directoryPath = "D:\HQ\DC Comics\001\00003 DEV"
$excel.Visible = $true                                    
$book = $excel.Workbooks.Open("C:\Users\Guilherme\OneDrive\Documents\Projetos\Pessoal\Automatizador de Quadrinhos\DC Comics.xlsx")
$excelFile = $book.Sheets(4)
$excelRow = $initialCheckLine
$FolderList=@(Get-ChildItem -Path $directoryPath -Directory | Select-Object Name)
Header
#RemoveFolder
$excelRow = $initialCheckLine
foreach($folder in $FolderList) {
    $ComicsContent=@()
    $nav=$folder.Name
    $testNav = $excelFile.Cells.Item($excelRow,1).Value2
    $directoryChildPath = Join-Path -Path $directoryPath -ChildPath $nav
    $directoryChildPathTest = Join-Path -Path $directoryPath -ChildPath $testNav
    $ComicsContent=@(Get-ChildItem -Path $directoryChildPath -File | Select-Object Name)
    Write-Host "Analysing $nav" -BackgroundColor Cyan -ForegroundColor Black
    Write-Host $excelFile.Cells.Item($excelRow,1).Value2
    Write-Host $folder.Name
    $registeredComics = CountRegisteredArc -line $excelRow -testNav $testNav
    $initialRange = $excelRow
    $regComics = $excelRow
    $trigger = 0
    $testPath = Test-Path -Path $directoryChildPathTest -PathType Container
    $countComics = $registeredComics.count
    if (!$testPath) {
    #if the registered folder doesn't exist
        while (!$testPath) {
            Write-Host "$testNav doesn't exist..." -BackgroundColor DarkRed
            $y = $excelRow + $countComics - 1
            RemoveRow -initialRange $excelRow -finalRange $y
            $lineComicsRegistered = $lineComicsRegistered - $countComics
            Write-Host "$testNav deleted." -BackgroundColor DarkRed
            $nextNav = $excelFile.Cells.Item($excelRow,1).Value2
            $directoryNextNav = Join-Path -Path $directoryPath -ChildPath $nextNav
            $testPath = Test-Path -Path $directoryNextNav -PathType Container
        }
    }
    Write-Host "$testNav exists." -BackgroundColor DarkYellow -ForegroundColor White                               
    $excelRow = $excelRow + $countCoomicsToJump
    if ($excelFile.Cells.Item($excelRow,1).Value2 -eq $folder.Name) {
    #if the folder is registered
        Write-Host "$nav already registered. Analysing Comics..." -BackgroundColor Yellow -ForegroundColor Black
        $initialRange = $excelRow
        $trigger = 0
        #verifying comics to remove
        if (($registeredComics.count -eq 1) -and ($ComicsContent.count -eq 1)) {
            $comictoremove = $excelFile.Cells.Item($regComics,2).Value2
            if ($comictoremove -notin $ComicsContent.name) {
                $excelFile.Cells.Item($regComics,2).Value2 = $ComicsContent.name
            }
        }
        elseif (($registeredComics.count -eq 1) -and ($ComicsContent.count -gt 1)) {
            $comictoremove = $excelFile.Cells.Item($regComics,2).Value2
            if ($comictoremove -notin $registeredComics.name) {
                $finalRange = $regComics + $ComicsContent.count - 1
                RemoveRow -initialRange $regComics
                InsertRow -initialRange $regComics -finalRange $finalRange
                $trigger++
                foreach($subfolder in $ComicsContent) {
                    $comic=$subfolder.Name
                    Write-Host "Registering $comic" -BackgroundColor DarkGreen
                    RegisterComic -row $regComics -subfolder $subfolder.Name
                    $regComics++
                }
            }
        }
        else {
            foreach ($book in $registeredComics) {
                $comictoremove = $excelFile.Cells.Item($regComics,2).Value2
                if ($comictoremove -notin $ComicsContent.name) {
                    if ($book -eq $comictoremove) {
                        RemoveRow -initialRange $regComics
                        $regComics--
                        $trigger++
                    }
                }
                $regComics++
            }
        }
        if ($trigger -gt 0) {
            $finalRange = $regComics-1
            RegisterFolder -row $initialRange -folder $folder.Name
            ExplodeRange -initialRange $initialRange -finalRange $finalRange
            TableStyle -initialRange $initialRange -finalRange $finalRange
        }
        #verifying comics to add
        foreach($subfolder in $ComicsContent) {
            $comic=$subfolder.Name
            $linecomictoremove = $excelRow
            $comictoremove = $excelFile.Cells.Item($linecomictoremove,2).Value2
            if ($excelFile.Cells.Item($excelRow,2).Value2 -eq $subfolder.Name) {
            #if the comic is registered
                Write-Host "$comic already registered." -BackgroundColor DarkYellow -ForegroundColor Black
                $excelRow++
            }
            else {
            #if the comic is not registered
                $trigger++
                Write-Host "Registering $comic" -BackgroundColor DarkGreen
                InsertRow -initialRange $excelRow
                RegisterComic -row $excelRow -subfolder $subfolder.Name
                $excelRow++
            }
        }
        if ($trigger -gt 0) {
            $finalRange = $excelRow-1
            ExplodeRange -initialRange $initialRange -finalRange $finalRange
            TableStyle -initialRange $initialRange -finalRange $finalRange
        }
    }
    else {
    #if the folder is not registered
        Write-Host "Registering $nav" -BackgroundColor Green -ForegroundColor Black
        $countComics = $ComicsContent.count
        $initialRange = $excelRow
        $finalRange = $excelRow + $countComics-1
        InsertRow -initialRange $initialRange -finalRange $finalRange
        RegisterFolder  -row $excelRow -folder $folder.Name     
        foreach($subfolder in $ComicsContent) {
            $comic=$subfolder.Name
            Write-Host "Registering $comic" -BackgroundColor DarkGreen
            RegisterComic -row $excelRow -subfolder $subfolder.Name
            $excelRow++
        }
        $finalRange = $excelRow - 1
        TableStyle -initialRange $initialRange -finalRange $finalRange
    }
}

Write-Host "****Comic Register Automator work is finished!****" -ForegroundColor White -BackgroundColor Blue
#DCLogo
