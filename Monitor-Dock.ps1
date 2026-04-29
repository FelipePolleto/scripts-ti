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
        [string]$DockDetectada,
        [string]$ErrosEventViewer,
        [string]$DCUExecutado,
        [string]$CorrecaoAplicada,
        [string]$RebootNecessario
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
🖥️ Diagnostico TI — Monitor via Dock Dell

Resultado: $Resultado $emoji
/Usuario: $Usuario
/Computador: $Computador
/Data/Hora: $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')
/Dock detectada: $DockDetectada
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
        Write-Log "Falha ao enviar notificacao ao Teams: $($_.Exception.Message)" "AVISO"
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

    try {
        $dock = Get-WmiObject -Namespace "root\dell\sysinv" `
                              -Class "dell_softwareidentity" `
                              -ErrorAction Stop |
                Where-Object { $_.ElementName -like "*Dock*" } |
                Select-Object ElementName, VersionString

        if ($dock) {
            Write-Log "Dock detectada: $($dock.ElementName) | Firmware: $($dock.VersionString)" "OK"
            $script:DockDetectada = $true
        } else {
            Write-Log "Nenhuma dock identificada via WMI Dell (pode ser normal em alguns modelos)." "AVISO"
            $script:DockDetectada = $false
        }
    } catch {
        Write-Log "Namespace WMI Dell nao disponível. Continuando via metodos alternativos." "AVISO"
        $script:DockDetectada = $false
    }

    $monitores = Get-WmiObject WmiMonitorID -Namespace root\wmi -ErrorAction SilentlyContinue
    $script:QtdMonitores = ($monitores | Measure-Object).Count
    Write-Log "Monitores detectados pelo Windows: $($script:QtdMonitores)" "INFO"

    $driversVideo = Get-WmiObject Win32_VideoController |
                    Select-Object Name, DriverVersion, Status
    foreach ($d in $driversVideo) {
        Write-Log "Driver vídeo: $($d.Name) | Versao: $($d.DriverVersion) | Status: $($d.Status)" "INFO"
    }

    $eventos = Get-WinEvent -FilterHashtable @{
        LogName   = 'System'
        Level     = 2
        StartTime = (Get-Date).AddHours(-24)
    } -ErrorAction SilentlyContinue |
    Where-Object { $_.Message -match "display|monitor|dock|USB|DisplayPort|Thunderbolt" } |
    Select-Object -First 5

    if ($eventos) {
        Write-Log "Erros recentes no Event Viewer relacionados a display/USB:" "AVISO"
        foreach ($e in $eventos) {
            Write-Log "  [$($e.TimeCreated)] $($e.Message.Substring(0,[Math]::Min(120,$e.Message.Length)))" "AVISO"
        }
        $script:ErrosEventViewer = $true
    } else {
        Write-Log "Nenhum erro crítico recente no Event Viewer para display/USB." "OK"
        $script:ErrosEventViewer = $false
    }

    $usbSeletivo = powercfg /query SCHEME_CURRENT SUB_USB 2>$null
    if ($usbSeletivo -match "AC Power Setting Index\s*:\s*0x00000001") {
        Write-Log "Suspensao seletiva de USB esta ATIVADA." "AVISO"
        $script:UsbSeletivoAtivo = $true
    } else {
        Write-Log "Suspensao seletiva de USB: OK." "OK"
        $script:UsbSeletivoAtivo = $false
    }
}

# ============================================================
# ETAPA 2 — TENTATIVAS DE CORREÇÃO AUTOMÁTICA
# ============================================================
function Tentar-Correcoes {
    Write-Log "=== INICIANDO TENTATIVAS DE CORRECAO AUTOMATICA ===" "INFO"
    $script:CorrecaoAplicada  = $false
    $script:RebootNecessario  = $false
    $script:DCUExecutado      = $false

    Write-Log "Tentativa 1: Forcando deteccao de displays..." "INFO"
    try {
        Start-Process "DisplaySwitch.exe" -ArgumentList "/extend" -Wait -ErrorAction Stop
        Start-Sleep -Seconds 3
        Write-Log "Deteccao de display forcada." "OK"
        $script:CorrecaoAplicada = $true
    } catch {
        Write-Log "Nao foi possível forcar deteccao via DisplaySwitch." "AVISO"
    }

    if ($script:UsbSeletivoAtivo) {
        Write-Log "Tentativa 2: Desativando suspensao seletiva de USB..." "INFO"
        try {
            powercfg /setacvalueindex SCHEME_CURRENT SUB_USB USBSELECTIVESUSPEND 0
            powercfg /setdcvalueindex SCHEME_CURRENT SUB_USB USBSELECTIVESUSPEND 0
            powercfg /setactive SCHEME_CURRENT
            Write-Log "Suspensao seletiva de USB desativada." "OK"
            $script:CorrecaoAplicada = $true
        } catch {
            Write-Log "Falha ao desativar suspensao seletiva de USB." "AVISO"
        }
    }

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
                1   { Write-Log "DCU: Reboot necessário para concluir." "AVISO"; $script:RebootNecessario = $true; $script:CorrecaoAplicada = $true }
                5   { Write-Log "DCU: Drivers já estao na versao mais recente." "OK" }
                500 { Write-Log "DCU: Nenhuma atualizacao disponivel no momento." "OK" }
                default { Write-Log "DCU: Codigo inesperado ($($updateResult.ExitCode)). Verifique o log DCU." "AVISO" }
            }
        } elseif ($scanResult.ExitCode -eq 500) {
            Write-Log "DCU: Nenhuma atualizacao disponivel (codigo 500). Drivers OK." "OK"
        } else {
            Write-Log "DCU scan falhou com codigo $($scanResult.ExitCode)." "ERRO"
        }
    }

    Write-Log "Tentativa 4: Reiniciando servico Plug and Play..." "INFO"
    try {
        Restart-Service -Name "PlugPlay" -Force -ErrorAction Stop
        Start-Sleep -Seconds 3
        Write-Log "Servico Plug and Play reiniciado." "OK"
        $script:CorrecaoAplicada = $true
    } catch {
        Write-Log "Nao foi possível reiniciar o servico Plug and Play." "AVISO"
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

    Write-Log "Monitores ANTES : $($script:QtdMonitores)" "INFO"
    Write-Log "Monitores APOS  : $($script:QtdMonitoresApos)" "INFO"

    if ($script:QtdMonitoresApos -gt $script:QtdMonitores) {
        $script:ResultadoFinal = "RESOLVIDO"
    } elseif ($script:QtdMonitoresApos -ge 2 -and -not $script:ErrosEventViewer -and -not $script:RebootNecessario) {
        $script:ResultadoFinal = "SEM PROBLEMAS"
    } elseif ($script:CorrecaoAplicada -and $script:RebootNecessario) {
        $script:ResultadoFinal = "REBOOT_NECESSARIO"
    } elseif ($script:QtdMonitoresApos -lt 2 -and -not $script:DockDetectada) {
        $script:ResultadoFinal = "PROVÁVEL HARDWARE"
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
            Write-Log "✅ SEM PROBLEMAS — Monitores estaveis, nenhum erro detectado." "OK"
        }
        "RESOLVIDO" {
            Write-Log "✅ RESOLVIDO — Monitor detectado apos correcoes automaticas." "OK"
        }
        "REBOOT_NECESSARIO" {
            Write-Log "🔄 REBOOT NECESSÁRIO — Atualizacoes aplicadas, aguardando reinicio." "AVISO"
        }
        "PROVÁVEL HARDWARE" {
            Write-Log "🔴 PROVÁVEL HARDWARE — Dock nao detectada e poucos monitores ativos." "ERRO"
            Write-Log "   Proximos passos:" "ERRO"
            Write-Log "     1. Testar dock em outro computador Dell" "ERRO"
            Write-Log "     2. Testar monitor direto no computador (sem dock)" "ERRO"
            Write-Log "     3. Testar com outro cabo USB-C/Thunderbolt" "ERRO"
            Write-Log "     4. Se falhar → acionar troca de hardware" "ERRO"
        }
        "INCONCLUSIVO" {
            Write-Log "⚠️  INCONCLUSIVO — Nao foi possível determinar a causa." "AVISO"
            Write-Log "   Verifique os logs e escale para análise manual." "AVISO"
        }
    }

    Write-Log "============================================" "INFO"

    Enviar-TeamsNotificacao `
        -Resultado        $script:ResultadoFinal `
        -Usuario          $env:USERNAME `
        -Computador       $env:COMPUTERNAME `
        -MonitoresAntes   "$($script:QtdMonitores)" `
        -MonitoresDepois  "$($script:QtdMonitoresApos)" `
        -DockDetectada    $(if ($script:DockDetectada)    { "Sim" } else { "Nao" }) `
        -ErrosEventViewer $(if ($script:ErrosEventViewer) { "Sim" } else { "Nao" }) `
        -DCUExecutado     $(if ($script:DCUExecutado)     { "Sim" } else { "Nao" }) `
        -CorrecaoAplicada $(if ($script:CorrecaoAplicada) { "Sim" } else { "Nao" }) `
        -RebootNecessario $(if ($script:RebootNecessario) { "Sim" } else { "Nao" })
}

# ============================================================
# EXECUÇÃO PRINCIPAL
# ============================================================
Write-Log "Script iniciado por: $($env:USERNAME) em $(Get-Date)" "INFO"

Coletar-Contexto
Tentar-Correcoes
Validar-Resultado
Exibir-Relatorio