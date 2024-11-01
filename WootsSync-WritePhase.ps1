<#
    WootsSync-WritePhase.ps1
    Versie 20241101
    p.wiegmans@svok.nl

    Verzorgt synchronisatie van Woots met Magister. 

    * leest uitvoer van WootsSync-ReadPhase.ps1 uit WootsSync.clixml

    Dit script hanteert een cache-mechanisme voor het snel benaderen van users, courses en classes.

    .INPUTS 
    WootsSync.ini
    Woots-Syncdata.clixml

    .OUTPUTS
    programma's in Woots

    .NOTES
    Maakt de volgende bestanden en werkt ze bij:
    Woots-UserCache.clixml
    Woots-CourseCache.clixml
    Woots-ClassCache.clixml

    Documentatie vind je in Readme.adoc.
    
    Changes
    20231013 : 
    * koppelt klassen af indien nodig
#>
[CmdletBinding()]
param (
    [Parameter(
        HelpMessage = "Geef de naam van de te gebruiken INI-file, bij verstek 'WootsSync.ini'"
    )]
    [Alias('Inifile', 'Inibestandsnaam', 'Config', 'Configfile', 'Configuratiebestand')]
    [String]  $Inifilename = "WootsSync.ini"
)
# ========  MODULEN  ========
Import-Module Woots
$stopwatch = [Diagnostics.Stopwatch]::StartNew()
$herePath = Split-Path -parent $MyInvocation.MyCommand.Definition
$inifile = "$herepath\$Inifilename"

# ========  VARIABELEN  ========
if ($env:computername -eq "EDUCONNECTOR") {
    $automationDataFolder = "C:\Automation\_Data"
}
else {
    $automationDataFolder = "$([environment]::getfolderpath('mydocuments'))\Automation\_Data"
}
$wootsSyncDataFolder = "$herepath\Data"
$wootssyncdatafile = "$wootsSyncDataFolder\WootsSync.clixml"
$huidigjaar = (Get-Date).AddMonths(-7).Year  # vanaf 1 aug, gebruik huidig jaar in formaat "2023"
$wcourses = @()
$coursecacheinvalidated = $False 
#Remove-Item -Path $coursecachefile -ea:SilentlyContinue  # forceer altijd ophalen
#Remove-Item -Path $classcachefile -ea:SilentlyContinue
#Remove-Item -Path $usercachefile -ea:SilentlyContinue
$filename_log = $null       # we loggen alleen naar console
$numbertoprocess = 999999
$verwerkingslimiet = 999999

#region functies
# ========  FUNCTIES  ========
function Assertfile($Path) {
    if (-not (Test-Path -Path $path -ea:SilentlyContinue)) {
        Throw "FOUT: bestand niet gevonden: $path"
    }
}
Function Sync-Log ($Path) {
    # logfilename instellen en sync 
    if ($Path) {
        $script:filename_log = $Path
    }
    if ($script:filename_log) {
        # hebben we de naam van het logbestand?
        $loglines | Out-File -FilePath $filename_log -Append
        $script:loglines = @()
    }
}
Function Write-Log {
    # Schrijf tekst  naar console en logbestand
    Param (
        [Parameter(Mandatory, Position = 0)][string]$Level,
        [Parameter(Position = 1)][Alias('Message', 'Text')][string]$Tekst
    )
    $kleur = 'Gray'
    if ($Level -match 'info') { $kleur = 'White' }
    if ($Level -match 'whatif') { $kleur = 'Blue' }
    if ($Level -match 'waarschuwing') { $kleur = 'Yellow' }
    if ($Level -match 'warn') { $kleur = 'Yellow' }
    if ($Level -match 'notice') { $kleur = 'Magenta' }
    if ($Level -match 'fout') { $kleur = 'Red' }
    if ($Level -match 'error') { $kleur = 'Red' }

    $tijd = Get-Date -f 's' # of als je milliseconden wilt 'yyyy-MM-ddTHH:mm:ss.fff'
    Write-Host ("{0} " -f $tijd) -ForegroundColor Blue -NoNewline
    Write-Host "[$level] $Tekst" -ForegroundColor $kleur

    $logline = "{0} {1} " -f ($tijd, $Tekst)
    $script:loglines += $logline
    if ($script:loglines.count -ge 100) {
        Sync-Log
    }
}
Function Clear-CourseCache {
    if (!$coursecacheinvalidated) {
        if (Test-Path $coursecachefile) {
            Remove-Item -Path $coursecachefile 
        }
        $coursecacheinvalidated = $True
    }
}

function Test-NeedsUpdate($Path) {
    #return $true
    if (Test-Path -Path $Path) {
        $status = (Get-ChildItem -Path $Path).LastWriteTime -lt (Get-Date).AddDays(-1)
        return $status 
    }
    return $True 
}
function CfgValidateBoolean ($value) {
    if ($value -eq '0' -or $value -eq '1') { return} 
    Throw "Config variabele $value is niet een geldige boolean waarde"
}

#endregion functies
#region main
# ========  MAIN  ========
Sync-Log -path $MyInvocation.MyCommand.Path.replace(".ps1", ".log")
Write-Log info "========  $(Split-Path -Leaf $MyInvocation.MyCommand.Definition)  ========"

# lees ini-file
AssertFile $inifile 
$cfg = Get-Content $inifile -Encoding UTF8 | ConvertFrom-StringData
if (!$cfg.hostname) { Throw "WootsSync.ini: hostname is verplicht" }
if (!$cfg.school_id) { Throw "WootsSync.ini: school_id is verplicht" }
if (!$cfg.token) { Throw "WootsSync.ini: token is verplicht" }
if (!$cfg.wootsinstantie) { Throw "WootsSync.ini: wootsinstantie is verplicht" }
CfgValidateBoolean $cfg.whatif
CfgValidateBoolean $cfg.whatif
$whatif = $cfg.whatif -eq '1'
$do_remove_instructors = $cfg.do_remove_instructors -eq '1'
if ($whatif) {Write-Log Notice "Whatif: $whatif"}

Initialize-Woots -hostname $cfg.hostname -school_id $cfg.school_id -token $cfg.token
Write-Log Notice "Verbonden met Woots $($cfg.wootsinstantie)"

$coursecachefile = "$automationDataFolder\Woots-$($cfg.wootsinstantie)-CourseCache.clixml"
$classcachefile = "$automationDataFolder\Woots-$($cfg.wootsinstantie)-ClassCache.clixml"
$usercachefile = "$automationDataFolder\Woots-$($cfg.wootsinstantie)-UserCache.clixml"

Assertfile $wootssyncdatafile
Write-Log info "Huidig jaar: $huidigjaar"
$programma = Import-Clixml -Path $wootssyncdatafile 
foreach ($prog in $programma) { $prog.id = $prog.id -as [int] } # maak numeriek sorteren mogelijk
Write-Log info ("Programma's totaal {0}" -f $programma.count)
$programma = $programma | Select-Object -First $numbertoprocess
Write-Log info ("Programma's te verwerken {0}" -f $programma.count)
$totaal = $programma.count
$teller = 0  # aantal verwerkte items
$aangepast = 0  # aantal programma's waarin wijzigingen zijn verwerkte

if (Test-NeedsUpdate $coursecachefile) {
    Write-Log info "Course cache vernieuwen..." 
    $wcourses = Get-WootsAllCourse
    $wcourses | Export-Clixml -Path $coursecachefile -Force -Depth 6
}
else {
    $wcourses = Import-Clixml -Path $coursecachefile
}
if (Test-NeedsUpdate $classcachefile) {
    Write-Log info "Class cache vernieuwen... "
    $wclasses = Get-WootsAllClass
    $wclasses | Export-Clixml -Path $classcachefile -Force -Depth 6
}
else {
    $wclasses = Import-Clixml -Path $classcachefile
}
if (Test-NeedsUpdate $usercachefile) {
    Write-Log info "User cache vernieuwen... "
    $wusers = Get-WootsAllUser
    $wusers | Export-Clixml -Path $usercachefile -Force -Depth 6
}
else {
    $wusers = Import-Clixml -Path $usercachefile
}

# maak course van dit jaar opzoekbaar op naam
$namecourse = @{}   # courses op naam
foreach ($c in ($wcourses | Where-Object { $_.year -eq $huidigjaar })) {
    $namecourse[$c.name] = $c
}

# maak class van dit jaar opzoekbaar op naam
$nameclass = @{}    # classes op naam
foreach ($c in ($wclasses | Where-Object { $_.year -eq $huidigjaar })) {
    $nameclass[$c.name] = $c
}

$upnuser = @{}  # maak users opzoekbaar op upn
$iduser = @{} # maak users opzoekbaar op id
foreach ($u in $wusers) {
    if ($u.email) {
        $upnuser[$u.email] = $u
    }
    $iduser[$u.id] = $u
}

#region loop
foreach ($prog in $programma) {
    # Stel de programmanaam opnieuw samen
    $teller += 1
    $studieletter = $prog.studie.substring(0, 1)
    $prognaam = $prog.naam
    #Write-Host "Programma ($($teller)/$($totaal)): $prognaam"
    # maak programma

    $course = $null
    if ($namecourse.ContainsKey($prognaam)) { 
        $course = Get-WootsCourse -id $namecourse[$prognaam].id  # haal het hele plaatje
    }
    else {
        # TO DO : labels meegeven vak, schoolniveau (vwo,..), leerjaar
        if (!$whatif) {
            $parms = @{
                name = $prognaam
                year = $huidigjaar 
            }
            Write-Log info "($($teller)/$($totaal)) +Programma: $prognaam"
            $course = Add-WootsCourse -Parameter $parms
            if ($course) {
                $aangepast += 1
                $namecourse[$course.name] = $course
                Clear-CourseCache # invalidate cache on disk
            }
        }
        else {
            Write-Log Whatif "($($teller)/$($totaal)) +Programma: $prognaam"
        }
    }
    if ($course) {
        # klassen aankoppelen
        foreach ($klasnaam in $prog.classes) {
            $class = $nameclass[$klasnaam]
            if ($class) {
                # test of klas niet al in course is gelinkt
                if ($class.id -notin $course.class_ids) {
                    if (!$whatif) {
                        Write-Log info "($($teller)/$($totaal)) +Klas $prognaam : $klasnaam"
                        $result = Add-WootsCourseCoursesClass -id $course.id -Parameter @{class_id = $class.id }
                        if (!$result) {
                            Write-Log warn "$(Get-WootsLastError)"
                        }
                    }
                    else {
                        Write-Log Whatif "($($teller)/$($totaal)) +Klas $prognaam : $klasnaam"
                    }
                }
            }
            else {
                Write-Log warn "    Klas niet gevonden in Woots $klasnaam"
            }
        }    

        # Klassen afkoppelen
        $courseclasses = Get-WootsCourseCoursesClass -id $course.id
        foreach ($c in $courseclasses) {
            $klas = $wclasses | Where-Object { $_.id -eq $c.class_id }
            if ($klas.name) {
                if ($klas.name -notin $prog.classes) {
                    if (!$whatif) {
                        Write-Log info "($($teller)/$($totaal)) -Klas $prognaam : $($klas.name)"
                        $result = Remove-WootsCoursesClass -id $c.id
                        if (!$result) {
                            Write-Log warn "$(Get-WootsLastError)"
                        }
                    }
                    else {
                        Write-Log whatif "($($teller)/$($totaal)) -Klas $prognaam : $($klas.name)"
                    }
                }
            }
        }
        # Koppel docenten
        if ($prog.docenten) {
            # synchroniseer docenten met de instructors
            $progdocentid = @()
            # voeg docenten toe die geen instructor zijn
            foreach ($upn in ($prog.docenten)) { 
                if ($upnuser.ContainsKey($upn)) {
                    $docent = $upnuser[$upn] # de docent bestaat in Woots
                    if ($docent.id -notin $course.instructors.user_id) {
                        $naam = "{0} {1} {2} ({3})" -f ($docent.first_name, $docent.middle_name, $docent.last_name, $upn)
                        if (!$whatif) {
                            Write-Log info ("($($teller)/$($totaal)) +Docent $prognaam : {0} {1}" -f ($docent.id, $naam))
                            $result = Add-WootsCourseCoursesUser -id $course.id -Parameter @{user_id = $docent.id; role = "instructor" }
                            if (!$result) {
                                Write-Log warn "$(Get-WootsLastError)"
                            }
                        }
                        else {
                            Write-Log whatif "($($teller)/$($totaal)) +Docent $prognaam : $naam"
                        }
                    }
                    $progdocentid += $docent.id
                }
                else {
                    Write-Log warn "    User bestaat niet $upn"
                }
            }
            # verwijder instructors die geen docent zijn
            if ($do_remove_instructors) {
                $instructors = Get-WootsCourseCoursesUser -id $course.id | Where-Object { $_.role -eq "instructor" }
                foreach ($docent in $instructors) {
                    # instructors aftellen
                    if ($docent.user_id -notin $progdocentid) {
                        $naam = "{0} {1} {2} " -f ($docent.first_name, $docent.middle_name, $docent.last_name)
                        if ($iduser.ContainsKey($docent.user_id)) {
                            $user = $iduser[$docent.user_id]
                            $naam = "{0} {1} {2}" -f ($user.first_name, $user.middle_name, $user.last_name)
                        }
                        if (!$whatif) {
                            Write-Log info ("($($teller)/$($totaal)) -Instructor: $prognaam : {0} {1}" -f ($docent.user_id, $naam))
                            $result = Remove-WootsCoursesUser -Id $docent.id
                            if (!$result) {
                                Write-Log warn "$(Get-WootsLastError)"
                            }
                        }
                        else {
                            Write-Log Whatif "($($teller)/$($totaal)) -Instructor $prognaam : $naam"
                        }
                    }
                }
            }
        }
    }

    if ($aangepast -ge $verwerkingslimiet) {
        Write-Host "Bewerkingslimiet bereikt!" -ForegroundColor Magenta; break
    }
    Sync-Log
}
#endregion loop

$stopwatch.Stop()
Write-Log info ("Klaar in " + $stopwatch.Elapsed.Hours + " uur " + $stopwatch.Elapsed.Minutes + " minuten " + $stopwatch.Elapsed.Seconds + " seconden ")    
Sync-Log
#endregion main