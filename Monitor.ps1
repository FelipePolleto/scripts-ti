# ============================================================
# CONFIGURAÇÕES GLOBAIS
# ============================================================
$DCU_PATH          = "C:\Program Files\Dell\CommandUpdate\dcu-cli.exe"
$DCU_LOG           = "$env:TEMP\DCU_Output.log"

$TEAMS_WEBHOOK_URL = "https://trten.webhook.office.com/webhookb2/9eccfb36-63e5-4e42-9877-c51fdd1ddf8d@62ccb864-6a1a-4b5d-8e1c-397dec1a8258/IncomingWebhook/3086f82d534f43148039cc95fc598513/a1b3c23d-9ac8-4bd5-aa6b-a32e395db538/V2kRIUB3N8-Shxz5FR5oXdz7BJREWC-hrIAeB95ZDBBQs1"

# ============================================================
# FUNÇÕES AUXILIARES
# ============================================================
function Write-Log {
    param([string]$Mensagem, [string]$Nivel = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $linha = "[$timestamp][$Nivel] $Mensagem"
    Write-Host $linha -ForegroundColor $(switch ($Nivel) {
        "INFO"    { "Cyan" }
        "OK"      { "Green" }
        "AVISO"   { "Yellow" }
        "ERRO"    { "Red" }
        default   { "White" }
    })
}

function Verificar-DCU {
    if (-not (Test-Path $DCU_PATH)) {
        $alt = "C:\Program Files (x86)\Dell\CommandUpdate\dcu-cli.exe"
        if (Test-Path $alt) {
            $script:DCU_PATH = $alt
            return $true
        }
        Write-Log "Dell Command Update (dcu-cli.exe) nao encontrado." "ERRO"
        return $false
    }
    return $true
}

# ============================================================
# FUNÇÃO DE NOTIFICAÇÃO TEAMS (Incoming Webhook)
# ============================================================
function Enviar-TeamsNotificacao {
    param(
        [string]$Resultado,
        [string]$Usuario,
        [string]$Computador,
        [string]$MonitoresAntes,
        [string]$MonitoresDepois,
        [string]$DriverStatus,
        [string]$ErrosEventViewer,
        [string]$DCUExecutado,
        [string]$CorrecaoAplicada,
        [string]$RebootNecessario,
        [string]$PortaUtilizada
    )

    $emoji = switch ($Resultado) {
        "RESOLVIDO"          { "" }
        "SEM PROBLEMAS"      { "" }
        "REBOOT_NECESSARIO"  { "" }
        "PROVÁVEL HARDWARE"  { "" }
        "INCONCLUSIVO"       { "️" }
        default              { "" }
    }

    $msg = @"
🖥️ Diagnostico TI — Monitor Direto no Notebook Dell

/Resultado: $Resultado $emoji
/Usuario: $Usuario
/Computador: $Computador
/Data/Hora: $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')
/Porta utilizada: $PortaUtilizada
/Driver de video: $DriverStatus
/Monitores antes: $MonitoresAntes
/Monitores depois: $MonitoresDepois
/Erros Event Viewer: $ErrosEventViewer
/DCU executado: $DCUExecutado
/Correcao aplicada: $CorrecaoAplicada
/Reboot necessario: $RebootNecessario
"@

    $body = @{ text = $msg } | ConvertTo-Json

    try {
        Invoke-RestMethod -Uri $TEAMS_WEBHOOK_URL -Method Post -Body $body -ContentType "application/json" -ErrorAction Stop
        Write-Log "Notificacao enviada ao Teams com sucesso." "OK"
    } catch {
        Write-Log "Falha ao enviar Notificacao ao Teams: $($_.Exception.Message)" "AVISO"
        Write-Log "O diagnostico foi concluido normalmente. Apenas a notificacao falhou." "AVISO"
    }
}

# ============================================================
# ETAPA 1 — COLETA DE CONTEXTO
# ============================================================
function Coletar-Contexto {
    Write-Log "=== INICIANDO COLETA DE CONTEXTO ===" "INFO"

    $pc = Get-WmiObject Win32_ComputerSystem
    Write-Log "Computador : $($pc.Manufacturer) $($pc.Model)" "INFO"
    Write-Log "Usuário    : $($env:USERNAME)" "INFO"
    Write-Log "SO         : $((Get-WmiObject Win32_OperatingSystem).Caption)" "INFO"

    # Verifica drivers de vídeo — ignora Unknown (comportamento normal em alguns adaptadores)
    $driversVideo = Get-WmiObject Win32_VideoController |
                    Sort-Object Name -Unique |
                    Select-Object Name, DriverVersion, Status
    $script:DriverStatus = "OK"
    foreach ($d in $driversVideo) {
        Write-Log "Driver video: $($d.Name) | Versao: $($d.DriverVersion) | Status: $($d.Status)" "INFO"
        if ($d.Status -ne "OK" -and $d.Status -ne "Unknown") {
            Write-Log "Driver com problema detectado: $($d.Name) — Status: $($d.Status)" "AVISO"
            $script:DriverStatus = "PROBLEMA"
        }
    }

    # Verifica portas de vídeo disponíveis no notebook — ignora Unknown e duplicatas
    Write-Log "Verificando portas de video disponiveis no notebook..." "INFO"
    $pnpDisplays = Get-PnpDevice -Class Display -ErrorAction SilentlyContinue |
                   Sort-Object FriendlyName -Unique |
                   Where-Object { $_.Status -ne "Unknown" }

    $script:PortaUtilizada = "Nao identificada"
    foreach ($pnp in $pnpDisplays) {
        Write-Log "Dispositivo de display: $($pnp.FriendlyName) | Status: $($pnp.Status)" "INFO"
        if ($pnp.FriendlyName -match "HDMI")                   { $script:PortaUtilizada = "HDMI" }
        elseif ($pnp.FriendlyName -match "DisplayPort")        { $script:PortaUtilizada = "DisplayPort" }
        elseif ($pnp.FriendlyName -match "USB-C|Thunderbolt")  { $script:PortaUtilizada = "USB-C/Thunderbolt" }
    }
    Write-Log "Porta utilizada detectada: $($script:PortaUtilizada)" "INFO"

    # Conta monitores ativos
    $monitores = Get-WmiObject WmiMonitorID -Namespace root\wmi -ErrorAction SilentlyContinue
    $script:QtdMonitores = ($monitores | Measure-Object).Count
    Write-Log "Monitores detectados pelo Windows: $($script:QtdMonitores)" "INFO"

    foreach ($m in $monitores) {
        $nome = ($m.UserFriendlyName | Where-Object { $_ -ne 0 } | ForEach-Object { [char]$_ }) -join ""
        Write-Log "  Monitor: $nome" "INFO"
    }

    # Verifica resolução e frequência — sem duplicatas
    $configs = Get-WmiObject Win32_VideoController |
               Sort-Object Name -Unique |
               Select-Object Name, CurrentHorizontalResolution, CurrentVerticalResolution, CurrentRefreshRate
    foreach ($c in $configs) {
        Write-Log "Resolucao: $($c.CurrentHorizontalResolution)x$($c.CurrentVerticalResolution) @ $($c.CurrentRefreshRate)Hz | $($c.Name)" "INFO"
    }

    # Verifica erros no Event Viewer
    $eventos = Get-WinEvent -FilterHashtable @{
        LogName   = 'System'
        Level     = 2
        StartTime = (Get-Date).AddHours(-24)
    } -ErrorAction SilentlyContinue |
    Where-Object { $_.Message -match "display|monitor|video|HDMI|DisplayPort|Thunderbolt" } |
    Select-Object -First 5

    if ($eventos) {
        Write-Log "Erros recentes no Event Viewer relacionados a display/vídeo:" "AVISO"
        foreach ($e in $eventos) {
            Write-Log "  [$($e.TimeCreated)] $($e.Message.Substring(0,[Math]::Min(120,$e.Message.Length)))" "AVISO"
        }
        $script:ErrosEventViewer = $true
    } else {
        Write-Log "Nenhum erro critico recente no Event Viewer para display/video." "OK"
        $script:ErrosEventViewer = $false
    }
}

# ============================================================
# ETAPA 2 — TENTATIVAS DE CORREÇÃO AUTOMATICA
# ============================================================
function Tentar-Correcoes {
    Write-Log "=== INICIANDO TENTATIVAS DE CORRECAO AUTOMATICA ===" "INFO"
    $script:CorrecaoAplicada  = $false
    $script:RebootNecessario  = $false
    $script:DCUExecutado      = $false

    # Tentativa 1: Forçar deteccao de displays
    Write-Log "Tentativa 1: Forcando deteccao de displays..." "INFO"
    try {
        Start-Process "DisplaySwitch.exe" -ArgumentList "/extend" -Wait -ErrorAction Stop
        Start-Sleep -Seconds 3
        Write-Log "Deteccao de display forçada." "OK"
        $script:CorrecaoAplicada = $true
    } catch {
        Write-Log "Nao foi possivel forcar deteccao via DisplaySwitch." "AVISO"
    }

    # Tentativa 2: Reiniciar apenas dispositivos de display com erro real (ignora Unknown e duplicatas)
    Write-Log "Tentativa 2: Reiniciando dispositivos de display via PnP..." "INFO"
    try {
        $pnpDisplays = Get-PnpDevice -Class Display -ErrorAction Stop |
                       Sort-Object FriendlyName -Unique |
                       Where-Object { $_.Status -ne "Unknown" }

        foreach ($pnp in $pnpDisplays) {
            Disable-PnpDevice -InstanceId $pnp.InstanceId -Confirm:$false -ErrorAction SilentlyContinue
            Start-Sleep -Seconds 2
            Enable-PnpDevice  -InstanceId $pnp.InstanceId -Confirm:$false -ErrorAction SilentlyContinue
            Write-Log "Dispositivo reiniciado: $($pnp.FriendlyName)" "OK"
        }
        $script:CorrecaoAplicada = $true
    } catch {
        Write-Log "Nao foi possível reiniciar dispositivos de display via PnP." "AVISO"
    }

    # Tentativa 3: Dell Command Update
    if (Verificar-DCU) {
        Write-Log "Tentativa 3: Executando scan via Dell Command Update..." "INFO"
        Write-Log "Isso pode levar alguns minutos. Aguarde..." "INFO"
        $script:DCUExecutado = $true

        $scanResult = Start-Process -FilePath $DCU_PATH `
                                    -ArgumentList "/scan -silent -outputlog=`"$DCU_LOG`"" `
                                    -Wait -PassThru -NoNewWindow

        if ($scanResult.ExitCode -eq 0) {
            Write-Log "Scan DCU concluído. Aplicando atualizacoes..." "OK"

            $updateResult = Start-Process -FilePath $DCU_PATH `
                                          -ArgumentList "/applyUpdates -silent -reboot=disable -outputlog=`"$DCU_LOG`"" `
                                          -Wait -PassThru -NoNewWindow

            switch ($updateResult.ExitCode) {
                0   { Write-Log "DCU: Atualizacoes aplicadas com sucesso." "OK"; $script:CorrecaoAplicada = $true }
                1   { Write-Log "DCU: Reboot necessario para concluir." "AVISO"; $script:RebootNecessario = $true; $script:CorrecaoAplicada = $true }
                5   { Write-Log "DCU: Drivers já estao na versao mais recente." "OK" }
                500 { Write-Log "DCU: Nenhuma atualizacao disponível no momento." "OK" }
                default { Write-Log "DCU: Codigo inesperado ($($updateResult.ExitCode)). Verifique o log DCU." "AVISO" }
            }
        } elseif ($scanResult.ExitCode -eq 500) {
            Write-Log "DCU: Nenhuma atualizacao disponivel (código 500). Drivers OK." "OK"
        } else {
            Write-Log "DCU scan falhou com codigo $($scanResult.ExitCode)." "ERRO"
        }
    }

    # Tentativa 4: Reiniciar serviço Plug and Play
    Write-Log "Tentativa 4: Reiniciando serviço Plug and Play..." "INFO"
    try {
        Restart-Service -Name "PlugPlay" -Force -ErrorAction Stop
        Start-Sleep -Seconds 3
        Write-Log "Serviço Plug and Play reiniciado." "OK"
        $script:CorrecaoAplicada = $true
    } catch {
        Write-Log "Nao foi possível reiniciar o serviço Plug and Play." "AVISO"
    }
}

# ============================================================
# ETAPA 3 — VALIDAÇÃO PÓS-CORREÇÃO
# ============================================================
function Validar-Resultado {
    Write-Log "=== VALIDANDO RESULTADO ===" "INFO"
    Start-Sleep -Seconds 5

    $monitoresApos = Get-WmiObject WmiMonitorID -Namespace root\wmi -ErrorAction SilentlyContinue
    $script:QtdMonitoresApos = ($monitoresApos | Measure-Object).Count

    # Re-verifica drivers de vídeo — ignora Unknown
    $driversApos = Get-WmiObject Win32_VideoController |
                   Sort-Object Name -Unique |
                   Where-Object { $_.Status -ne "OK" -and $_.Status -ne "Unknown" }
    $driverStatusApos = if ($driversApos) { "PROBLEMA" } else { "OK" }

    Write-Log "Monitores ANTES      : $($script:QtdMonitores)" "INFO"
    Write-Log "Monitores APOS       : $($script:QtdMonitoresApos)" "INFO"
    Write-Log "Driver Status ANTES  : $($script:DriverStatus)" "INFO"
    Write-Log "Driver Status APOS   : $driverStatusApos" "INFO"

    if ($script:QtdMonitoresApos -gt $script:QtdMonitores) {
        $script:ResultadoFinal = "RESOLVIDO"
    } elseif ($script:QtdMonitoresApos -ge 1 -and -not $script:ErrosEventViewer -and $driverStatusApos -eq "OK") {
        $script:ResultadoFinal = "SEM PROBLEMAS"
    } elseif ($script:CorrecaoAplicada -and $script:RebootNecessario) {
        $script:ResultadoFinal = "REBOOT_NECESSARIO"
    } elseif ($script:QtdMonitoresApos -lt 1 -and $driverStatusApos -eq "PROBLEMA") {
        $script:ResultadoFinal = "PROVAVEL HARDWARE"
    } else {
        $script:ResultadoFinal = "INCONCLUSIVO"
    }
}

# ============================================================
# ETAPA 4 — RELATÓRIO FINAL + NOTIFICAÇÃO TEAMS
# ============================================================
function Exibir-Relatorio {
    Write-Log "============================================" "INFO"
    Write-Log "        RESULTADO DO DIAGNOSTICO            " "INFO"
    Write-Log "============================================" "INFO"

    switch ($script:ResultadoFinal) {
        "SEM PROBLEMAS" {
            Write-Log "✅ SEM PROBLEMAS — Monitor estavel, driver OK, nenhum erro detectado." "OK"
        }
        "RESOLVIDO" {
            Write-Log "✅ RESOLVIDO — Monitor detectado apos correções automáticas." "OK"
        }
        "REBOOT_NECESSARIO" {
            Write-Log "🔄 REBOOT NECESSÁRIO — Atualizacoes aplicadas, aguardando reinicio." "AVISO"
        }
        "PROVÁVEL HARDWARE" {
            Write-Log "🔴 PROVAVEL HARDWARE — Monitor nao detectado e driver com problema." "ERRO"
            Write-Log "   Próximos passos:" "ERRO"
            Write-Log "     1. Testar monitor com outro cabo (HDMI/DisplayPort/USB-C)" "ERRO"
            Write-Log "     2. Testar monitor em outro computador Dell" "ERRO"
            Write-Log "     3. Testar outra porta de video disponivel no notebook" "ERRO"
            Write-Log "     4. Se falhar → acionar troca de hardware" "ERRO"
        }
        "INCONCLUSIVO" {
            Write-Log "⚠️  INCONCLUSIVO — Nao foi possível determinar a causa." "AVISO"
            Write-Log "   Verifique os logs e escale para analise manual." "AVISO"
        }
    }

    Write-Log "============================================" "INFO"

    Enviar-TeamsNotificacao `
        -Resultado        $script:ResultadoFinal `
        -Usuario          $env:USERNAME `
        -Computador       $env:COMPUTERNAME `
        -MonitoresAntes   "$($script:QtdMonitores)" `
        -MonitoresDepois  "$($script:QtdMonitoresApos)" `
        -DriverStatus     $script:DriverStatus `
        -ErrosEventViewer $(if ($script:ErrosEventViewer) { "Sim" } else { "Nao" }) `
        -DCUExecutado     $(if ($script:DCUExecutado)     { "Sim" } else { "Nao" }) `
        -CorrecaoAplicada $(if ($script:CorrecaoAplicada) { "Sim" } else { "Nao" }) `
        -RebootNecessario $(if ($script:RebootNecessario) { "Sim" } else { "Nao" }) `
        -PortaUtilizada   $script:PortaUtilizada
}

# ============================================================
# EXECUÇÃO PRINCIPAL
# ============================================================
Write-Log "Script iniciado por: $($env:USERNAME) em $(Get-Date)" "INFO"

Coletar-Contexto
Tentar-Correcoes
Validar-Resultado
Exibir-Relatorio