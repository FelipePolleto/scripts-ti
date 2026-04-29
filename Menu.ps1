function Executar-Script {
    param (
        [string]$url,
        [string]$nome
    )
    $destino = "$env:TEMP\$nome"
    Write-Host "Baixando e executando $nome..." -ForegroundColor Cyan

    try {
        Invoke-WebRequest -Uri $url -OutFile $destino -UseBasicParsing
        & powershell -ExecutionPolicy Bypass -NoExit -File $destino
    }
    catch {
        Write-Host "ERRO: $_" -ForegroundColor Red
        Write-Host "Pressione qualquer tecla para voltar ao menu..."
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    }
}

while ($true) {
    Clear-Host
    Write-Host "================================" -ForegroundColor Yellow
    Write-Host "        MENU DE SCRIPTS TI      " -ForegroundColor Yellow
    Write-Host "================================" -ForegroundColor Yellow
    Write-Host "1. Dock"
    Write-Host "2. Limpeza V3.2"
    Write-Host "3. Logs"
    Write-Host "4. Monitor"
    Write-Host "5. Monitor Dock"
    Write-Host "0. Sair"
    Write-Host "================================" -ForegroundColor Yellow

    $op = Read-Host "Digite a opcao"

    switch ($op) {
        1 { Executar-Script "https://raw.githubusercontent.com/FelipePolleto/scripts-ti/main/Dock.ps1" "Dock.ps1" }
        2 { Executar-Script "https://raw.githubusercontent.com/FelipePolleto/scripts-ti/main/LimpezaV3.2.ps1" "LimpezaV3.2.ps1" }
        3 { Executar-Script "https://raw.githubusercontent.com/FelipePolleto/scripts-ti/main/Logs.ps1" "Logs.ps1" }
        4 { Executar-Script "https://raw.githubusercontent.com/FelipePolleto/scripts-ti/main/Monitor.ps1" "Monitor.ps1" }
        5 { Executar-Script "https://raw.githubusercontent.com/FelipePolleto/scripts-ti/main/Monitor-Dock.ps1" "Monitor-Dock.ps1" }
        0 { exit }
        default { Write-Host "Opcao invalida!" -ForegroundColor Red; Start-Sleep -Seconds 2 }
    }
}
