function Test-Admin {
    $currentUser = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    $isAdmin = $currentUser.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    return $isAdmin
}

if (-not (Test-Admin)) {
    Write-Host "Ce script nécessite des privilèges d'administrateur. Veuillez exécuter PowerShell en tant qu'administrateur." -ForegroundColor Red
    exit
}

function Rearm-Windows {
    Write-Host "Réarmement de la licence Windows..." -ForegroundColor Yellow
    try {
        $process = Start-Process -FilePath "cscript.exe" -ArgumentList "$env:SystemRoot\System32\slmgr.vbs /rearm" -Wait -PassThru -NoNewWindow
        if ($process.ExitCode -eq 0) {
            Write-Host "Réarmement de Windows réussi!" -ForegroundColor Green
        } else {
            Write-Host "Échec du réarmement de Windows. Code d'erreur: $($process.ExitCode)" -ForegroundColor Red
        }
    } catch {
        Write-Host "Une erreur s'est produite lors du réarmement de Windows: $_" -ForegroundColor Red
    }
}

function Rearm-Office {
    Write-Host "Recherche des installations d'Office..." -ForegroundColor Yellow
    
    $officePaths = @(
        "C:\Program Files\Microsoft Office\Office16",
        "C:\Program Files (x86)\Microsoft Office\Office16",
        "C:\Program Files\Microsoft Office\Office15",
        "C:\Program Files (x86)\Microsoft Office\Office15",
        "C:\Program Files\Microsoft Office\Office14",
        "C:\Program Files (x86)\Microsoft Office\Office14"
    )
    
    $officeFound = $false
    
    foreach ($path in $officePaths) {
        $ospprearmPath = Join-Path -Path $path -ChildPath "OSPPREARM.EXE"
        
        if (Test-Path $ospprearmPath) {
            $officeFound = $true
            Write-Host "Installation d'Office trouvée à: $path" -ForegroundColor Cyan
            
            Write-Host "Fermeture de toutes les applications Office..." -ForegroundColor Yellow
            $officeApps = @("WINWORD", "EXCEL", "POWERPNT", "OUTLOOK", "ONENOTE", "MSACCESS", "MSPUB", "GROOVE", "INFOPATH")
            foreach ($app in $officeApps) {
                $running = Get-Process -Name $app -ErrorAction SilentlyContinue
                if ($running) {
                    Stop-Process -Name $app -Force
                    Write-Host "Application $app fermée." -ForegroundColor Cyan
                }
            }
            
            Write-Host "Réarmement de la licence Office..." -ForegroundColor Yellow
            try {
                $process = Start-Process -FilePath $ospprearmPath -Wait -PassThru -NoNewWindow
                if ($process.ExitCode -eq 0) {
                    Write-Host "Réarmement d'Office réussi!" -ForegroundColor Green
                } else {
                    Write-Host "Échec du réarmement d'Office. Code d'erreur: $($process.ExitCode)" -ForegroundColor Red
                }
            } catch {
                Write-Host "Une erreur s'est produite lors du réarmement d'Office: $_" -ForegroundColor Red
            }
        }
    }
    
    if (-not $officeFound) {
        Write-Host "Aucune installation compatible de Microsoft Office n'a été trouvée sur cet ordinateur." -ForegroundColor Red
    }
}

function Show-Menu {
    Clear-Host
    Write-Host "========== SCRIPT DE RÉARMEMENT DE LICENCES ==========" -ForegroundColor Cyan
    Write-Host "1: Réarmer Windows uniquement" -ForegroundColor White
    Write-Host "2: Réarmer Office uniquement" -ForegroundColor White
    Write-Host "3: Réarmer Windows et Office" -ForegroundColor White
    Write-Host "Q: Quitter" -ForegroundColor White
    Write-Host "=====================================================" -ForegroundColor Cyan
    
    $choice = Read-Host "Veuillez faire votre choix"
    
    switch ($choice) {
        '1' {
            Rearm-Windows
            Write-Host "`nUn redémarrage de l'ordinateur est nécessaire pour finaliser le réarmement de Windows." -ForegroundColor Yellow
            $restart = Read-Host "Voulez-vous redémarrer maintenant? (O/N)"
            if ($restart -eq 'O' -or $restart -eq 'o') {
                Restart-Computer -Force
            }
        }
        '2' {
            Rearm-Office
        }
        '3' {
            Rearm-Windows
            Rearm-Office
            Write-Host "`nUn redémarrage de l'ordinateur est nécessaire pour finaliser le réarmement de Windows." -ForegroundColor Yellow
            $restart = Read-Host "Voulez-vous redémarrer maintenant? (O/N)"
            if ($restart -eq 'O' -or $restart -eq 'o') {
                Restart-Computer -Force
            }
        }
        'Q' {
            exit
        }
        'q' {
            exit
        }
        default {
            Write-Host "Choix non valide. Veuillez réessayer." -ForegroundColor Red
            Start-Sleep -Seconds 2
            Show-Menu
        }
    }
}

Show-Menu
