#Automatizador de quadrinhos

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
    $newRow = $rangeToInsert.Insert([System.Type]::Missing)
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
    $revRow = $range0.Delete([System.Type]::Missing)
}

Function RegisterFolder {
    param(
        [int]$row,
        [string]$folder
        )
    $excelFile.Cells.Item($row,1) = $folder
    $excelFile.Cells.Item($row,1).Font.Bold=$true
    $excelFile.Cells.Item($row,5) = "Média Arco:"
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
    while ($excelFile.Cells.Item($lineComicsRegistered,2).Value2 -ne $null) {
        $lineComicsRegistered++
    }
    $teste = $excelFile.Cells.Item($excelLine,1).Value2    
    while ($excelLine -lt $lineComicsRegistered) {
        Write-Host "COMICS REGISTERED $lineComicsRegistered" -ForegroundColor White -BackgroundColor Black
        Write-Host "LINE $excelLine" -ForegroundColor White -BackgroundColor Black
        Write-Host "ENTRANDO NO while do REMOVEFOLDER" -ForegroundColor White -BackgroundColor Black        
        $remNav = $excelFile.Cells.Item($excelLine,1).Value2
        $regArc = CountRegisteredArc -line $excelLine -testNav $testNav
        $countCoomicsToJump = $regArc.count
        $directoryChildPathTest = Join-Path -Path $directoryPath -ChildPath $remNav
        if (Test-Path -Path $directoryChildPathTest -PathType Container) {
            Write-Host "$remNav exists." -BackgroundColor DarkYellow -ForegroundColor White                               
            $excelLine = $excelLine + $countCoomicsToJump
        } else {                
            $remNav = $excelFile.Cells.Item($excelLine,1).Value2
            Write-Host "$remNav does not exists." -BackgroundColor Magenta
            Write-Host "$remNav doesn't exists..." -BackgroundColor DarkRed
            $y = $excelLine + $countCoomicsToJump - 1
            RemoveRow -initialRange $excelLine -finalRange $y
            $lineComicsRegistered = $lineComicsRegistered - $countCoomicsToJump
            <#while ($y -lt $countCoomicsToJump) {
                RemoveRow($excelLine)
                $y++
                $lineComicsRegistered--
            }#>
            Write-Host "$remNav deleted." -BackgroundColor DarkRed
        }
    }
    $excelLine = 3
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
$excelLine = $initialCheckLine
$folderName=@(Get-ChildItem -Path $directoryPath -Directory | Select-Object Name)
Header
RemoveFolder
$excelLine = $initialCheckLine
foreach($folder in $folderName) {
    $nav=$folder.Name
    $testNav = $excelFile.Cells.Item($excelLine,1).Value2
    $directoryChildPath = Join-Path -Path $directoryPath -ChildPath $nav
    $directoryChildPathTest = Join-Path -Path $directoryPath -ChildPath $testNav
    $subfolderName=@(Get-ChildItem -Path $directoryChildPath -File | Select-Object Name)
    Write-Host "Analysing $nav" -BackgroundColor Cyan -ForegroundColor Black
    Write-Host $excelFile.Cells.Item($excelLine,1).Value2
    Write-Host $folder.Name
    $registeredComics = CountRegisteredArc -line $excelLine -testNav $testNav
    $initialRange = $excelLine
    $regComics = $excelLine
    $trigger = 0
    if ($registeredComics) {
        foreach ($book in $registeredComics) {
            $comictoremove = $excelFile.Cells.Item($regComics,2).Value2
            if ($comictoremove -notin $subfolderName.name) {
                if ($book -eq $comictoremove) {
                    RemoveRow -row $regComics
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
    if ($excelFile.Cells.Item($excelLine,1).Value2 -eq $folder.Name) {
        #se o folder ta registrado
        Write-Host "$nav already registered. Analysing Comics..." -BackgroundColor Yellow -ForegroundColor Black
        $initialRange = $excelLine
        $countComics = $subfolderName.count
        $trigger = 0            
        foreach($subfolder in $subfolderName) {
            $comic=$subfolder.Name
            $linecomictoremove = $excelLine
            $comictoremove = $excelFile.Cells.Item($linecomictoremove,2).Value2
            if ($excelFile.Cells.Item($excelLine,2).Value2 -eq $subfolder.Name) {
                #se o comic ta registrado
                Write-Host "$comic already registered." -BackgroundColor DarkYellow -ForegroundColor Black
                $excelLine++
            }
            else {
                #se o comic não ta registrado
                $trigger++
                Write-Host "Registering $comic" -BackgroundColor DarkGreen
                InsertRow -initialRange $excelLine
                RegisterComic -row $excelLine -subfolder $subfolder.Name
                $excelLine++
                $finalRange = $excelLine-1
            }
        }
        if ($trigger -gt 0) {
            $finalRange = $excelLine-1
            ExplodeRange -initialRange $initialRange -finalRange $finalRange
            TableStyle -initialRange $initialRange -finalRange $finalRange
        }
    }
    else {
        #se o folder não ta registrado
        Write-Host "Registering $nav" -BackgroundColor Green -ForegroundColor Black
        $initialRange = $excelLine
        $finalRange = $excelLine + $countComics - 1
        InsertRow -initialRange $initialRange -finalRange $finalRange
        <#foreach($subfolder in $subfolderName) {
            InsertRow -row $excelLine
        }#>
        RegisterFolder  -row $excelLine -folder $folder.Name     
        foreach($subfolder in $subfolderName) {
            $comic=$subfolder.Name
            Write-Host "Registering $comic" -BackgroundColor DarkGreen
            RegisterComic -row $excelLine -subfolder $subfolder.Name
            $excelLine++
            $finalRange = $excelLine-1
            #TableStyle
        }
        TableStyle -initialRange $initialRange -finalRange $finalRange
    }
}
        
$excelLine = 3
$folder = $excelFile.Cells.Item($excelLine,1).Value2
$lineComicsRegistered = $initialCheckLine
Write-Host "****Comic Register Automator work is finished!****" -ForegroundColor White -BackgroundColor Blue
#DCLogo
