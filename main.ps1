#Automatizador de quadrinhos

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
    #Write-Host "Analysing $nav" -BackgroundColor Cyan -ForegroundColor Black
    #Write-Host $excelFile.Cells.Item($excelLine,1).Value2
    #Write-Host $folder.Name
    $registeredComics = @()
    $checkregcomic = $excelLine
    $regcom = $testNav
    $boolean = $true
    $nextcomic = $excelFile.Cells.Item($checkregcomic,2).Value2
    while ($boolean) {
        $regcom = $excelFile.Cells.Item($checkregcomic,1).Value2
        $regbook = $excelFile.Cells.Item($checkregcomic,2).Value2
        if ($regcom -ne "$testNav" -and $regcom -ne $null) {
            $boolean = $false
            $checkregcomic++
        }
        elseif ($regcom -eq $null -and $regbook -eq $null) {
            $boolean = $false
        }

        if ($boolean) {
            $registeredComics += $excelFile.Cells.Item($checkregcomic,2).Value2
            $checkregcomic++
        }
    }
    Write-Output "Folder: $folder"
    foreach ($comic in $registeredComics) {
        Write-Output "-   $comic"
    }
    #Write-Output "Array: $registeredComics"
    if ($excelFile.Cells.Item($excelLine,1).Value2 -eq $folder.Name) {
        #se o folder ta registrado
        Write-Host "$nav already registered. Analysing Comics..." -BackgroundColor Yellow -ForegroundColor Black
        $initialRange = $excelLine
        $countComics = 0
        foreach($subfolder in $subfolderName) {
            $countComics++
        }
        $trigger = 0            
        foreach($subfolder in $subfolderName) {
            $comic=$subfolder.Name
            $directoryComic = Join-Path -Path $directoryChildPath -ChildPath $comic 
            $comictoremove = $excelLine
            if ($subfolder.Name -notin $registeredComics) {
                #se o comic registrado não existe no folder
                Write-Host "$comic registered is not in the folder." -BackgroundColor DarkYellow -ForegroundColor Black
                #procurar linha com comic não existente
                foreach ($c in $registeredComics) {
                    $f = $excelFile.Cells.Item($comictoremove,2).Value2
                    if ($c -eq $f) {
                        RemoveRow($comictoremove)
                    }
                    $comictoremove++
                }
                $trigger++
                RegisterFolder
            } elseif ($excelFile.Cells.Item($comictoremove,2).Value2 -notin $subfolderName.name) {
                $linn = $comictoremove
                foreach($subfolder in $excelFile.Cells.Item($comictoremove,2).Value2) {
                    $f = $excelFile.Cells.Item($linn,2).Value2
                    if ($f -eq $subfolder.name) {
                        RemoveRow($linn)
                    }
                    $linn++
                }
                $comictoremove++
            }
            if ($excelFile.Cells.Item($excelLine,2).Value2 -eq $subfolder.Name) {
                #se o comic ta registrado
                $directoryComic = Join-Path -Path $directoryChildPath -ChildPath $comic
                Write-Host "$comic already registered." -BackgroundColor DarkYellow -ForegroundColor Black
                $excelLine++
                <#if ($trigger -gt 0) {
                    $finalRange = $excelLine-1
                    $countComicsDIF = $finalRange - $initialRange
                    #ExplodeRange
                    #TableStyle
                }#>
            }
            else {
                #se o comic não ta registrado
                $trigger++
                Write-Host "Registering $comic" -BackgroundColor DarkGreen
                InsertRow
                RegisterComic
                $excelLine++
                $finalRange = $excelLine-1
                #ExplodeRange
                #TableStyle
            }
        }
        if ($trigger -gt 0) {
            $finalRange = $excelLine-1
            $countComicsDIF = $finalRange - $initialRange
            ExplodeRange
            TableStyle
        }
    }
    else {
        #se o folder não ta registrado
        Write-Host "Registering $nav" -BackgroundColor Green -ForegroundColor Black
        foreach($subfolder in $subfolderName) {
            InsertRow               
        }
        RegisterFolder
        $initialRange = $excelLine
        foreach($subfolder in $subfolderName) {
            $countComics++
        }
        Write-Host "Comic: $countComics antes"-BackgroundColor Cyan -ForegroundColor Black         
        foreach($subfolder in $subfolderName) {
            $comic=$subfolder.Name
            $directoryComic = Join-Path -Path $directoryChildPath -ChildPath $comic
            Write-Host "Registering $comic" -BackgroundColor DarkGreen
            RegisterComic
            $excelLine++
            $finalRange = $excelLine-1
            #TableStyle
            Write-Host "Comic: $countComics DEPOIS"-BackgroundColor Cyan -ForegroundColor Black
        }
        TableStyle
    }
}
        
$excelLine = 3
$lastFolderAnalyzed = $nav
$folder = $excelFile.Cells.Item($excelLine,1).Value2
$lineComicsRegistered = $initialCheckLine
Write-Host "****Comic Register Automator work is finished!****" -ForegroundColor White -BackgroundColor Blue
#DCLogo

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
    $rangeToInsert = $excelFile.Rows.Item($excelLine)
    $newRow = $rangeToInsert.Insert([System.Type]::Missing)
}

Function RemoveRow {
    param($line)
    $range0 = $excelFile.Range(("A{0}" -f $line),("F{0}" -f $line))
    $range0.Select()
    $rangeToRemove = $excelFile.Rows.Item($line)
    $revRow = $rangeToRemove.Delete([System.Type]::Missing)
}

Function RegisterFolder {
    $excelFile.Cells.Item($excelLine,1) = $folder.Name
    $excelFile.Cells.Item($excelLine,1).Font.Bold=$true
    $excelFile.Cells.Item($excelLine,5) = "Média Arco:"
    $excelFile.Cells.Item($excelLine,5).Font.Bold=$true
}

Function RegisterComic {
    $excelFile.Cells.Item($excelLine,2) = $subfolder.Name
    $excelFile.Cells.Item($excelLine,2).Font.Bold=$false
}

Function ExplodeRange {
    $range1 = $excelFile.Range(("A{0}" -f $initialRange),("A{0}" -f $finalRange))
        $range1.MergeCells = $false
    $range2 = $excelFile.Range(("E{0}" -f $initialRange),("E{0}" -f $finalRange))
        $range2.MergeCells = $false
    $range3 = $excelFile.Range(("F{0}" -f $initialRange),("F{0}" -f $finalRange))
        $range3.MergeCells = $false
}
    
Function TableStyle {
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

Function VerifyFolder {
    if (Test-Path -Path $directoryChildPath -PathType Container) {
        return $true
    } else {
        return $false
    }
}

Function VerifyComic {
    if (Test-Path -Path $directoryComic -PathType Container) {
        return $true
    } else {
        return $false
    }
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
        $directoryChildPathTest = Join-Path -Path $directoryPath -ChildPath $remNav
        if (Test-Path -Path $directoryChildPathTest -PathType Container) {
            Write-Host "$remNav exists." -BackgroundColor DarkYellow -ForegroundColor White                
            $countCoomicsToJump = 1
            $x = $excelLine
            <#while ($excelFile.Cells.Item($x,1).Value2 -eq $null) {
                $countCoomicsToJump++
                $x++
            }#>                
            $excelLine = $excelLine + $countCoomicsToJump
        } else {                
            $remNav = $excelFile.Cells.Item($excelLine,1).Value2
            Write-Host "$remNav does not exists." -BackgroundColor Magenta
            Write-Host "$remNav doesn't exists..." -BackgroundColor DarkRed
            $x = $excelLine + 1
            $countCoomicsToRemove = 1
            $boolean = $true
            <#while ($excelFile.Cells.Item($x,1).Value2 -eq $null) {
                $countCoomicsToRemove++
                $x++
            }#>
            while ($boolean) {
                if (($excelFile.Cells.Item($x,1).Value2 -eq $null) -and ($excelFile.Cells.Item($x,2).Value2 -eq $null)) {
                    $boolean = $false
                }
                else {
                    $countCoomicsToRemove++
                    $x++
                }
            }
            $rangeToRemove = $excelLine + $countCoomicsToRemove
            $y = 0
            do {
                RemoveRow($excelLine)
                $y++
                $lineComicsRegistered--
            } while ($y -lt $countCoomicsToRemove)
            Write-Host "$remNav deleted." -BackgroundColor DarkRed
        }
    }
    $excelLine = 3
}
