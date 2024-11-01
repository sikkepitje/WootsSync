<#
    WootsSync-ReadPhase.ps1
    Versie 20241101
    p.wiegmans@svok.nl

    WootsSync maakt programma's voor elk vak van leerlingen.
    Leest leerlingen, docenten en vakken uit Magister, 
    bepaalt de naam van programma's

    Aanpassing 2023-08-17 
    In de zomervakantie ontbreekt belangrijke informatie: aan docenten zijn geen groepen, 
    klassen of vakken toegekend. 
    Om het toch mogelijk te maken om Woots-programma's aan te maken, bekijkt het script nu
    alleen leerlingen en hun lidmaatschap van lesgroepen. 
    Voor een programma is nodig: leerjaar, studie en vak. Hieraan worden klassen of lesgroepen gekoppeld. 
    Voor onderbouw worden klassen gekoppeld.
    voor bovenbouw worden lesgroepen gekoppeld.

    Aanpassing 2023-10-04 
    Maakt programma's voor alle klassen in onderbouw, maar in de bovenbouw bovendien voor alle klassen 
    genoemd in het bestand WSKlassikaleVakken.txt.
    
    Aanpassing 2024-9-20 t.b.v. Castor College
    Veel berekeningen zijn JPT-specifiek. Structuur moet helemaal opnieuw geanalyseerd worden. 
    Gebruik docent.id direct, want in ScholenHB is magister.id gelijk aan e-mailadres. 
    Script gebruikt docentfilter in ini-file. 
    CAS groepnaam opbouw: {string[1-2]:afdeling}{int[1]:leerjaar}"."{string:vaknaam}{string[1]:volgletter}

    .INPUTS
    WootsSync.ini
    magister_docent.clixml
    magister_leerling.clixml
    magister_vak.clixml
    WsExcludeGroep.txt
    WsExcludeStudie.txt
    WsKlassikaleVakken.txt
    WsIncludeStudie.txt

    .OUTPUTS
    WootsSync.clixml
    WootsSync.csv

    WIJZIGINGEN
    20240920 Aanpassing t.b.v. Castor College: 
    * gebruiker docent.id direct, want is gelijk aan e-mailadres. Voeg geen '@jpthijsse.nl' toe!
    * ini-file 
#>
[CmdletBinding()]
param (
    [Parameter(
        HelpMessage = "Geef de naam van de te gebruiken INI-file, bij verstek 'WootsSync.ini'"
    )]
    [Alias('Inifile', 'Inibestandsnaam', 'Config', 'Configfile', 'Configuratiebestand')]
    [String]  $Inifilename = "WootsSync.ini"
)
# ========  VARIABELEN  =======
$stopwatch = [Diagnostics.Stopwatch]::StartNew()
$herePath = Split-Path -parent $MyInvocation.MyCommand.Definition
$inifile = "$herepath\$Inifilename"
$vaknaamtabel = @{}
$programmateller = 0
$programma = @{}

#region functies
# ========  FUNCTIES  ========
function Assertfile($Path) {
    if (-not (Test-Path -Path $path -ea:SilentlyContinue)) {
        Throw "FOUT: map of bestand niet gevonden: $path"
    }
}

function Capitalize($naam) {
    return $naam.substring(0, 1).toupper() + $naam.substring(1)
    # niet helemaal hetzelfde: (Get-Culture).TextInfo.ToTitleCase($string)
}

function GetAddProgram($key, $leerjaar, $studie, $vak) {
    # maak nieuw programma, maak aan indien nodig
    if (-not $programma.ContainsKey($key)) {
        $programma[$key] = [PSCustomObject]@{
            id       = ++$script:programmateller
            key      = $key
            naam     = 'onbekend'
            studie   = 'onbekend'
            leerjaar = 'onbekend'
            vak      = 'onbekend'
            classes  = @()
            docenten = @()
        }
        #write-host "Programma toevoegen [$key] $leerjaar $studie $vak"
        $p = $programma[$key]
        if ($studie) { $p.studie = $studie }
        if ($leerjaar) { $p.leerjaar = $leerjaar }
        if ($vak) { $p.vak = $vak }    
    }
    $p = $programma[$key]
    return $p
}
function Add-ProgramVakFiltered($leerjaar, $studie, $vak, $newclass) {
    # voeg newclass toe aan programma, maak programma indien nodig
    $key = "{0}{1}-{2}" -f ($studie, $leerjaar, $vak)
    $p = GetAddProgram -key $key -leerjaar $leerjaar -studie $studie -vak $vak
    if ($newclass -notin $p.classes) {
        $p.classes += $newclass
    }
}
function Add-ProgramDocent($studie, $leerjaar, $vak, $docent) {
    $key = "{0}{1}-{2}" -f ($studie, $leerjaar, $vak)
    if ($programma.ContainsKey($key)) {
        $p = $programma[$key]
        if ($docent -notin $p.docenten) {
            $p.docenten += $docent
        }
    }
}
function isNumeriek([string]$c) {
    # geef True als teken $c numeriek is
    return ($c[0] -ge '0' -and $c[0] -le '9')
}
function Splits($s) {
    # Splits een string in delen die elk geheel numeriek zijn of geheel alfa,
    # en retourneer deze delen als een arraylist.
    if (!$s) { return $null }
    $parts = @()
    $part = $s[0]
    $isnumlast = isNumeriek $s[0]
    $s = $s.substring(1)
    while ($s) {
        $current = $s[0]
        $s = $s.substring(1)
        $isnumcurrent = isNumeriek $current
        if ($isnumlast -ne $isnumcurrent) {
            $parts += $part # voeg een nieuw element toe aan de array
            $part = $current
        }
        else {
            $part += $current
        }
        $isnumlast = $isnumcurrent
    }
    $parts += $part # voeg de rest toe als element aan de array
    return $parts
}
function CfgValidateBoolean ($value) {
    if ($value -eq '0' -or $value -eq '1') { return } 
    Throw "Config variabele $value is niet een geldige boolean waarde"
}

#endregion functies
#region main
# ========  MAIN  ========
Write-Host "========  $(Split-Path -Leaf $MyInvocation.MyCommand.Definition)  ========"

# lees ini-file
AssertFile $inifile 
$cfg = Get-Content $inifile -Encoding UTF8 | ConvertFrom-StringData
if (!$cfg.school) { Throw "WootsSync.ini: school is verplicht" }
if (!$cfg.importmap) { Throw "WootsSync.ini: importmap is verplicht" }
if (!$cfg.datamap) { Throw "WootsSync.ini: datamap is verplicht" }
if (!$cfg.tempmap) { Throw "WootsSync.ini: tempmap is verplicht" }
if (!$cfg.magistersyncleerjaar) { Throw "WootsSync.ini: magistersyncleerjaar is verplicht" }
CfgValidateBoolean $cfg.gridview
$leerjaren = $cfg.magistersyncleerjaar -split ','
$usegridview = $cfg.gridview -eq '1'

# bestanden
$ImportFolder = "$herepath\$($cfg.importmap)"
$docentfile = "$ImportFolder\magister_docent.clixml"
$leerlingfile = "$ImportFolder\magister_leerling.clixml" 
$vaknaamfile = "$ImportFolder\Magister_vak.clixml"

$WSDataFolder = "$herepath\$($cfg.datamap)"
$out_clixml_file = "$WSDataFolder\WootsSync.clixml"
$out_csv_file = "$WSDataFolder\WootsSync.csv"
$out_txt_file = "$WSDataFolder\WootsSync.txt"
$filterInclStudiefile = "$WSDataFolder\WSIncludeStudie.txt"
$filterExclStudiefile = "$WSDataFolder\WSExcludeStudie.txt"
$filterExclVakFile = "$WSDataFolder\WSExcludeVakcode.txt"
$filterExclGroepFile = "$WSDataFolder\WSExcludeGroep.txt"
$klasvakkenFile = "$WSDataFolder\WSKlassikaleVakken.txt"

# lees importbestanden
Assertfile $docentfile
Assertfile $leerlingfile
Assertfile $vaknaamfile
$docent = Import-Clixml -Path $docentfile
Write-Host "Aantal docenten ingelezen: $($docent.count)"
if ($cfg.filterdocentlocatie) {
    $docent = $docent | Where-Object { $_.Locatie -eq $cfg.filterdocentlocatie }
    Write-Host "Aantal docenten na locatiefilter $($cfg.filterdocentlocatie): $($docent.count)"
}
$leerling = Import-Clixml -Path $leerlingfile
$vaknaamtabel = Import-Clixml -Path $vaknaamfile
Write-Host "Aantal leerlingen ingelezen: $($leerling.count)"

# lijst met klassikale vakken inlezen
$klassikaleVakken = @()
if (Test-path $klasvakkenFile) {
    $klassikaleVakken = Get-Content $klasvakkenFile 
    Write-Host "We hebben klassikale vakken: $klassikaleVakken" -ForegroundColor Yellow
}

# filter leerlingen insluitend op studie (alle matches worden meegenomen)
if (Test-Path $filterInclStudiefile) {
    $filter = $(Get-Content -path $filterInclStudiefile) -join '|'
    $filtleer = [System.Collections.Generic.List[object]]::new()
    foreach ($l in $leerling) {
        if ($l.Studie -match $filter) {
            $filtleer.Add($l) | Out-Null
        }
    }
    $leerling = $filtleer
    Write-Host "Aantal leerlingen na inclusief filteren op studie: $($leerling.count)"
}

# filter leerlingen uitsluitend op studie (alle matches worden weggelaten)
if (Test-Path $filterExclStudiefile) {
    $filter = $(Get-Content -path $filterExclStudiefile) -join '|'
    $filtleer = [System.Collections.Generic.List[object]]::new()
    foreach ($l in $leerling) {
        if ($l.Studie -notmatch $filter) {
            $filtleer.Add($l) | Out-Null
        }
    }
    $leerling = $filtleer
    Write-Host "Aantal leerlingen na exclusief filteren op studie: $($leerling.count)"
}
# laad vakuitsluitfilter
$ExclVakFilter = $null
if (Test-path $filterExclVakFile) {
    $ExclVakFilter = $(Get-Content -path $filterExclVakFile) -join '|'
}
# laad groepuitsluitfilter
$ExclGroepFilter = $null
if (Test-path $filterExclGroepFile) {
    $ExclGroepFilter = $(Get-Content -path $filterExclGroepFile) -join '|'
}

# Bewaar lijsten met alle klassen, lesgroepen, vakken, studies voor inspectie.
$klas = [System.Collections.Generic.List[string]]::new()
$groep = [System.Collections.Generic.List[string]]::new()
$vak = [System.Collections.Generic.List[string]]::new()
$studie = [System.Collections.Generic.List[string]]::new()
foreach ($l in $leerling) {
    $klas.add($l.klas)
    $studie.add($l.studie)
    foreach ($v in $l.Vakken) {
        $vak.add($v)
    }
    foreach ($g in $l.Groepen) {
        $groep.add($g)
    }
}
$klas = $klas | Sort-Object -Unique
$vak = $vak | Sort-Object -Unique
$groep = $groep | Sort-Object -Unique
$studie = $studie | Sort-Object -Unique
Write-Host "Klassen:" $klas.count
Write-Host "Vakken:" $vak.count
write-host "Groepen:" $groep.count
Write-Host "Studies:" $studie.count
$klas | Out-File -FilePath "$herepath\$($cfg.tempmap)\temp-klas.txt"
$vak | Out-File -FilePath "$herepath\$($cfg.tempmap)\temp-vak.txt"
$groep | Out-File -FilePath "$herepath\$($cfg.tempmap)\temp-groep.txt"
$studie | Out-File -FilePath "$herepath\$($cfg.tempmap)\temp-studie.txt"

#region verzamel
<# 
Maak programma's aan op afdeling en vak en voeg docenten en leerlingen toe.
De klas/lesgroep in Magister is opgebouwd uit afdeling, punt, vakcode, volgnummer.
Afdeling is opgebouwd uit studie (bijv v, h, m) en leerjaar.
#>

foreach ($l in $leerling) {
    switch ($cfg.school) {
        'JPT' {
            $wstudy = $l.Studie.substring(0, 1)
            $leerjaar = $l.Studie.substring(1, 1)
        }
        'CAS' {
            $wstudy = $l.Studie.substring(0, $l.Studie.Length - 1)
            $leerjaar = $l.Studie.substring($l.Studie.Length - 1, 1)
        }
    }
    
    if ($l.Groepen.count -gt 0) {
        foreach ($g in $l.groepen) {
            if ($g.contains('.')) {
                # skip raar geval waar leerling geen groepen heeft, en lijst heeft 1 element van lengte 1 met niets.
                switch ($cfg.school) {
                    'JPT' {
                        $vak = $g.split('.')[1]
                        $vak = $vak.substring(0, $vak.length - 1) 
                        Add-ProgramVakFiltered -leerjaar $leerjaar -studie $wstudy -vak $vak -newclass $g
                    }
                    'CAS' {
                        # hier moet ik groep "TOP.*" eruit filteren 
                        if ($g -notmatch $ExclGroepFilter) {
                            $vak = $g.split('.')[1]
                            $vak = $vak.substring(0, $vak.length - 1) 
                            Add-ProgramVakFiltered -leerjaar $leerjaar -studie $wstudy -vak $vak -newclass $g
                        }
                    }
                }
            }
        }
    }    
    if ($l.Vakken.count -gt 0) {
        foreach ($v in $l.Vakken) {
            switch ($cfg.school) {
                'JPT' { 
                    if ($leerjaar -le 3 -or $v -in $klassikaleVakken) {
                        # is onderbouw? dan maak klasvak-groepen
                        Add-ProgramVakFiltered -leerjaar $leerjaar -studie $wstudy -vak $v -newclass $l.Klas
                    }
                }
                'CAS' {
                    if ($leerjaar -le 3 -or $v -in $klassikaleVakken) {
                        # is onderbouw? dan maak klasvak-groepen
                        Add-ProgramVakFiltered -leerjaar $leerjaar -studie $wstudy -vak $v -newclass $l.Klas
                    }
                }
            }
        }
    }
    else {
        Write-Host "! Leerling heeft geen groepen:" $l.Email $l.Id $l.Achternaam
    }
}

Write-Host "Zoek programma's bij docenten..."
# zoek programma's bij docenten
foreach ($d in $docent) {
    foreach ($gv in $d.groepvakken) {
        if ($g.contains('.')) {
            switch ($cfg.school) {
                'JPT' {
                    $afdeling = $gv.Klas.split('.')[0]
                    $leerjaar, $wstudy, $rest = splits $afdeling
                    Add-ProgramDocent -studie $wstudy -leerjaar $leerjaar -vak $gv.vakcode -docent $d.id
                }
                'CAS' {
                    $afdeling = $gv.Klas.split('.')[0]
                    $wstudy, $leerjaar, $rest = splits $afdeling
                    Add-ProgramDocent -studie $wstudy -leerjaar $leerjaar -vak $gv.vakcode -docent $d.id
                }
            }
        }
    }
    # tel alle klasvakken en docentvakken af, test op lege lijst uitzondering 
    foreach ($kv in $d.klasvakken) {
        foreach ($dv in $d.docentvakken) {
            switch ($cfg.school) {
                'JPT' {
                    $leerjaar = $kv.substring(0, 1)
                    $wstudy = $kv.substring(1, 1)
                    Add-ProgramDocent -Studie $wstudy -leerjaar $leerjaar -vak $dv -docent $d.id
                }
                'CAS' {
                    $wstudy, $leerjaar, $rest = splits $kv
                    Add-ProgramDocent -Studie $wstudy -leerjaar $leerjaar -vak $dv -docent $d.id
                }
            }
        }
    }
}
#endregion verzamel

# genereer namen. Schooljaarstring "23/24" hoeft niet , wordt automatisch toegevoegd.
foreach ($p in $programma.Values) {
    $vaknaam = 'Onbekend-vak ' + $p.vak
    if ($vaknaamtabel.containskey($p.vak)) {
        $vaknaam = Capitalize $vaknaamtabel[$p.vak]
    }
    switch ($cfg.school) {
        'JPT' {    
            $club = "$($p.leerjaar)$($p.studie)" # op JPT: 4V
            if ($club -eq '1B') {
                $club = 'Brugklas'
            }
        }
        'CAS' {
            $club = "$($p.studie)$($p.leerjaar)"  # op CAS: V4
        }
    }
    $p.Naam = "$vaknaam $club"
}

# maak platte lijst
$programma = $programma.Values | Sort-Object naam
Write-host "Aantal programma's: $($programma.count)"

# filter op vakcode
if ($ExclVakFilter) {
    $programmafiltered = $programma | ForEach-Object {
        if ($_.vak -notmatch $ExclVakFilter) {
            $_
        }    
    }
    $programma = $programmafiltered
    Write-Host "Programma na exclusief filteren op vak:" $programma.count 
}

# filter leerjaren
$programmafiltered = $programma | ForEach-Object {
    if ($_.leerjaar -in $leerjaren) {
        $_
    }    
}
$programma = $programmafiltered
Write-Host "Programma na inclusief filteren op leerjaar:" $programma.count 

#region uitvoer
# uitvoeren

if ($usegridview) { $programma | Export-Clixml -Path $out_clixml_file }

# maak platte lijst voor CSV export 
foreach ($p in $programma) {
    $p.classes = $p.classes -join ','
    $p.docenten = $p.docenten -join ','
}
$programma | Export-csv -Path $out_csv_file -Delimiter ";" -Encoding UTF8 -NoTypeInformation
$programma.naam | Out-File -FilePath $out_txt_file -Encoding utf8 

# maak selectie voor interactieve viewer
$programma | Out-Gridview

$stopwatch.Stop()
Write-Host ("Klaar in " + $stopwatch.Elapsed.Hours + " uur " + $stopwatch.Elapsed.Minutes + " minuten " + $stopwatch.Elapsed.Seconds + " seconden ")    
#endregion uitvoer
#endregion main
