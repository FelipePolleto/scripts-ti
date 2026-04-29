# ============================================================
# CONFIGURAÇÕES GLOBAIS
# ============================================================
$DCU_PATH          = "C:\Program Files\Dell\CommandUpdate\dcu-cli.exe"
$DCU_LOG           = "$env:TEMP\DCU_Output.log"
$DISPLAYLINK_LOG   = "$env:TEMP\DisplayLink_Reinstall.log"

$TEAMS_WEBHOOK_URL = "https://trten.webhook.office.com/webhookb2/9eccfb36-63e5-4e42-9877-c51fdd1ddf8d@62ccb864-6a1a-4b5d-8e1c-397dec1a8258/IncomingWebhook/3086f82d534f43148039cc95fc598513/a1b3c23d-9ac8-4bd5-aa6b-a32e395db538/V2kRIUB3N8-Shxz5FR5oXdz7BJREWC-hrIAeB95ZDBBQs1"

# ============================================================
# FUNCOES AUXILIARES
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
# FUNÇÃO: VERIFICAR E CORRIGIR DISPLAYLINK
# ============================================================
function Verificar-DisplayLink {
    Write-Log "=== VERIFICANDO DRIVER DISPLAYLINK ===" "INFO"
    $script:DisplayLinkProblema = $false
    $script:DisplayLinkCorrigido = $false

    # Verifica se ha erros do DisplayLink no Event Viewer
    $eventosDisplayLink = Get-WinEvent -FilterHashtable @{
        LogName   = 'System'
        Level     = 2
        StartTime = (Get-Date).AddHours(-24)
    } -ErrorAction SilentlyContinue |
    Where-Object { $_.Message -match "DisplayLink" } |
    Select-Object -First 5

    # Verifica adaptadores de rede DisplayLink com problema
    $adapterDisplayLink = Get-PnpDevice -ErrorAction SilentlyContinue |
                          Sort-Object FriendlyName -Unique |
                          Where-Object {
                              $_.FriendlyName -match "DisplayLink" -and
                              $_.Status -ne "OK" -and
                              $_.Status -ne "Unknown"
                          }

    if ($eventosDisplayLink -or $adapterDisplayLink) {
        $script:DisplayLinkProblema = $true

        if ($eventosDisplayLink) {
            Write-Log "Erros DisplayLink detectados no Event Viewer:" "AVISO"
            foreach ($e in $eventosDisplayLink) {
                Write-Log "  [$($e.TimeCreated)] $($e.Message.Substring(0,[Math]::Min(120,$e.Message.Length)))" "AVISO"
            }
        }

        if ($adapterDisplayLink) {
            foreach ($a in $adapterDisplayLink) {
                Write-Log "Adaptador DisplayLink com problema: $($a.FriendlyName) | Status: $($a.Status)" "AVISO"
            }
        }

        # Tenta corrigir reiniciando todos os dispositivos DisplayLink
        Write-Log "Tentando reiniciar dispositivos DisplayLink..." "INFO"
        try {
            $dispositivosDisplayLink = Get-PnpDevice -ErrorAction Stop |
                                       Sort-Object FriendlyName -Unique |
                                       Where-Object { $_.FriendlyName -match "DisplayLink" }

            foreach ($d in $dispositivosDisplayLink) {
                Write-Log "  Reiniciando: $($d.FriendlyName) | Status atual: $($d.Status)" "INFO"
                Disable-PnpDevice -InstanceId $d.InstanceId -Confirm:$false -ErrorAction SilentlyContinue
                Start-Sleep -Seconds 2
                Enable-PnpDevice  -InstanceId $d.InstanceId -Confirm:$false -ErrorAction SilentlyContinue
                Write-Log "  Dispositivo reiniciado: $($d.FriendlyName)" "OK"
            }
            Start-Sleep -Seconds 3
            $script:DisplayLinkCorrigido = $true
            $script:CorrecaoAplicada = $true
        } catch {
            Write-Log "Nao foi possível reiniciar dispositivos DisplayLink via PnP." "AVISO"
        }

        # Verifica se o problema persiste após o reinício
        $adapterApos = Get-PnpDevice -ErrorAction SilentlyContinue |
                       Sort-Object FriendlyName -Unique |
                       Where-Object {
                           $_.FriendlyName -match "DisplayLink" -and
                           $_.Status -ne "OK" -and
                           $_.Status -ne "Unknown"
                       }

        if ($adapterApos) {
            Write-Log "Problema DisplayLink persiste apos reinício. Recomendada reinstalacao manual do driver." "AVISO"
            Write-Log "  → Acesse: dell.com/support e busque pelo modelo da Dock (ex: UD22)" "AVISO"
            Write-Log "  → Desinstale o driver DisplayLink atual pelo Gerenciador de Dispositivos" "AVISO"
            Write-Log "  → Instale o driver baixado e reinicie o notebook" "AVISO"
            $script:DisplayLinkCorrigido = $false
            $script:RebootNecessario = $true
        } else {
            Write-Log "Dispositivos DisplayLink normalizados após reinício." "OK"
        }

    } else {
        Write-Log "Nenhum problema detectado no driver DisplayLink." "OK"
    }
}

# ============================================================
# FUNÇÃO DE NOTIFICAÇÃO TEAMS
# ============================================================
function Enviar-TeamsNotificacao {
    param(
        [string]$Resultado,
        [string]$Usuario,
        [string]$Computador,
        [string]$DockDetectada,
        [string]$FirmwareDock,
        [string]$PortasUSB,
        [string]$PortasVideo,
        [string]$AlimentacaoDock,
        [string]$ErrosEventViewer,
        [string]$DisplayLinkStatus,
        [string]$DCUExecutado,
        [string]$CorrecaoAplicada,
        [string]$RebootNecessario
    )

    $emoji = switch ($Resultado) {
        "RESOLVIDO"          { "" }
        "SEM PROBLEMAS"      { "" }
        "REBOOT_NECESSARIO"  { "" }
        "PROVAVEL HARDWARE"  { "" }
        "INCONCLUSIVO"       { "" }
        default              { "" }
    }

    $msg = @"
-->> Diagnostico TI — Dock Station Dell <<--

Resultado: $Resultado $emoji
/Usuario: $Usuario
/Computador: $Computador
/Data/Hora: $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')
/Dock detectada: $DockDetectada
/Firmware da Dock: $FirmwareDock
/Portas USB: $PortasUSB
/Portas de video (DP/HDMI): $PortasVideo
/Alimentacao da Dock: $AlimentacaoDock
/Driver DisplayLink: $DisplayLinkStatus
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
    Write-Log "Usuario    : $($env:USERNAME)" "INFO"
    Write-Log "SO         : $((Get-WmiObject Win32_OperatingSystem).Caption)" "INFO"

    # Verifica Dock via WMI Dell
    $script:DockDetectada = $false
    $script:FirmwareDock  = "Nao identificado"
    try {
        $dock = Get-WmiObject -Namespace "root\dell\sysinv" `
                              -Class "dell_softwareidentity" `
                              -ErrorAction Stop |
                Where-Object { $_.ElementName -like "*Dock*" } |
                Select-Object ElementName, VersionString

        if ($dock) {
            Write-Log "Dock detectada: $($dock.ElementName) | Firmware: $($dock.VersionString)" "OK"
            $script:DockDetectada = $true
            $script:FirmwareDock  = $dock.VersionString
        } else {
            Write-Log "Dock nao identificada via WMI Dell." "AVISO"
        }
    } catch {
        Write-Log "Namespace WMI Dell nao disponível. Tentando metodo alternativo..." "AVISO"
    }

    # Método alternativo: detectar dock via dispositivos USB/Thunderbolt
    if (-not $script:DockDetectada) {
        $dockUSB = Get-PnpDevice -ErrorAction SilentlyContinue |
                   Where-Object { $_.FriendlyName -match "Dell.*Dock|Thunderbolt|WD19|WD22|TB16|TB18|UD22|D6000" } |
                   Sort-Object FriendlyName -Unique |
                   Where-Object { $_.Status -eq "OK" }

        if ($dockUSB) {
            foreach ($d in $dockUSB) {
                Write-Log "Dock detectada via PnP: $($d.FriendlyName) | Status: $($d.Status)" "OK"
            }
            $script:DockDetectada = $true
        } else {
            Write-Log "Nenhuma Dock Dell detectada via PnP." "ERRO"
        }
    }

    # Verifica portas USB da Dock — ignora Unknown
    Write-Log "Verificando portas USB da Dock..." "INFO"
    $script:PortasUSB = "OK"
    $usbDispositivos = Get-PnpDevice -Class USB -ErrorAction SilentlyContinue |
                       Sort-Object FriendlyName -Unique |
                       Where-Object { $_.Status -ne "OK" -and $_.Status -ne "Unknown" }

    if ($usbDispositivos) {
        foreach ($usb in $usbDispositivos) {
            Write-Log "  Problema USB: $($usb.FriendlyName) | Status: $($usb.Status)" "AVISO"
        }
        $script:PortasUSB = "PROBLEMA"
    } else {
        Write-Log "Todas as portas USB estao OK." "OK"
    }

    # Verifica portas de video da Dock — ignora Unknown e duplicatas
    Write-Log "Verificando portas de vídeo da Dock (DisplayPort/HDMI)..." "INFO"
    $script:PortasVideo = "OK"
    $videoDispositivos = Get-PnpDevice -Class Display -ErrorAction SilentlyContinue |
                         Sort-Object FriendlyName -Unique |
                         Where-Object { $_.Status -ne "Unknown" -or $_.FriendlyName -notmatch "Dock" }

    foreach ($v in $videoDispositivos) {
        Write-Log "  Display: $($v.FriendlyName) | Status: $($v.Status)" "INFO"
        if ($v.Status -ne "OK") {
            $script:PortasVideo = "PROBLEMA"
            Write-Log "  Problema detectado em porta de vídeo: $($v.FriendlyName)" "AVISO"
        }
    }

    # Verifica alimentação da Dock
    Write-Log "Verificando alimentacao da Dock..." "INFO"
    $script:AlimentacaoDock = "Nao verificado"
    try {
        $bateria = Get-WmiObject Win32_Battery -ErrorAction SilentlyContinue
        if ($bateria) {
            $statusCarga = switch ($bateria.BatteryStatus) {
                1       { "Descarregando (sem alimentacao externa)" }
                2       { "AC conectado — Carregando" }
                3       { "AC conectado — Carga completa" }
                default { "Status desconhecido ($($bateria.BatteryStatus))" }
            }
            Write-Log "Status de energia: $statusCarga" "INFO"
            $script:AlimentacaoDock = $statusCarga
            if ($bateria.BatteryStatus -eq 1) {
                Write-Log "Dock pode nao estar fornecendo energia ao notebook." "AVISO"
            }
        } else {
            Write-Log "Informacao de bateria nao disponivel." "AVISO"
            $script:AlimentacaoDock = "Nao disponível"
        }
    } catch {
        Write-Log "Nao foi possivel verificar alimentacao." "AVISO"
    }

    # Verifica suspensão seletiva de USB
    $usbSeletivo = powercfg /query SCHEME_CURRENT SUB_USB 2>$null
    if ($usbSeletivo -match "AC Power Setting Index\s*:\s*0x00000001") {
        Write-Log "Suspensao seletiva de USB esta ATIVADA." "AVISO"
        $script:UsbSeletivoAtivo = $true
    } else {
        Write-Log "Suspensao seletiva de USB: OK." "OK"
        $script:UsbSeletivoAtivo = $false
    }

    # Verifica erros no Event Viewer
    $eventos = Get-WinEvent -FilterHashtable @{
        LogName   = 'System'
        Level     = 2
        StartTime = (Get-Date).AddHours(-24)
    } -ErrorAction SilentlyContinue |
    Where-Object { $_.Message -match "dock|USB|DisplayPort|Thunderbolt|HDMI|display|DisplayLink" } |
    Select-Object -First 5

    if ($eventos) {
        Write-Log "Erros recentes no Event Viewer relacionados à Dock/USB/Display/DisplayLink:" "AVISO"
        foreach ($e in $eventos) {
            Write-Log "  [$($e.TimeCreated)] $($e.Message.Substring(0,[Math]::Min(120,$e.Message.Length)))" "AVISO"
        }
        $script:ErrosEventViewer = $true
    } else {
        Write-Log "Nenhum erro critico recente no Event Viewer." "OK"
        $script:ErrosEventViewer = $false
    }
}

# ============================================================
# ETAPA 2 — TENTATIVAS DE CORRECAO AUTOMATICA
# ============================================================
function Tentar-Correcoes {
    Write-Log "=== INICIANDO TENTATIVAS DE CORRECAO AUTOMATICA ===" "INFO"
    $script:CorrecaoAplicada  = $false
    $script:RebootNecessario  = $false
    $script:DCUExecutado      = $false

    # Tentativa 1: Forçar deteccao de displays
    Write-Log "Tentativa 1: Forçando deteccao de displays via Dock..." "INFO"
    try {
        Start-Process "DisplaySwitch.exe" -ArgumentList "/extend" -Wait -ErrorAction Stop
        Start-Sleep -Seconds 3
        Write-Log "Deteccao de display forçada." "OK"
        $script:CorrecaoAplicada = $true
    } catch {
        Write-Log "Nao foi possível forçar deteccao via DisplaySwitch." "AVISO"
    }

    # Tentativa 2: Desativar suspensão seletiva de USB
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

    # Tentativa 3: Reiniciar dispositivos USB com erro real (ignora Unknown)
    if ($script:PortasUSB -eq "PROBLEMA") {
        Write-Log "Tentativa 3: Reiniciando dispositivos USB com problema..." "INFO"
        try {
            $usbProblema = Get-PnpDevice -Class USB -ErrorAction Stop |
                           Sort-Object FriendlyName -Unique |
                           Where-Object { $_.Status -ne "OK" -and $_.Status -ne "Unknown" }

            foreach ($usb in $usbProblema) {
                Disable-PnpDevice -InstanceId $usb.InstanceId -Confirm:$false -ErrorAction SilentlyContinue
                Start-Sleep -Seconds 2
                Enable-PnpDevice  -InstanceId $usb.InstanceId -Confirm:$false -ErrorAction SilentlyContinue
                Write-Log "Dispositivo USB reiniciado: $($usb.FriendlyName)" "OK"
            }
            $script:CorrecaoAplicada = $true
        } catch {
            Write-Log "Nao foi possivel reiniciar dispositivos USB." "AVISO"
        }
    }

    # Tentativa 4: Verificar e corrigir DisplayLink
    Write-Log "Tentativa 4: Verificando e corrigindo driver DisplayLink..." "INFO"
    Verificar-DisplayLink

    # Tentativa 5: Dell Command Update (firmware da Dock)
    if (Verificar-DCU) {
        Write-Log "Tentativa 5: Executando scan via Dell Command Update (inclui firmware Dock)..." "INFO"
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
                5   { Write-Log "DCU: Drivers/Firmware ja estao na versao mais recente." "OK" }
                500 { Write-Log "DCU: Nenhuma atualizacao disponivel no momento." "OK" }
                default { Write-Log "DCU: Codigo inesperado ($($updateResult.ExitCode)). Verifique o log DCU." "AVISO" }
            }
        } elseif ($scanResult.ExitCode -eq 500) {
            Write-Log "DCU: Nenhuma atualizacao disponível (código 500). Firmware OK." "OK"
        } else {
            Write-Log "DCU scan falhou com codigo $($scanResult.ExitCode)." "ERRO"
        }
    }

    # Tentativa 6: Reiniciar serviço Plug and Play
    Write-Log "Tentativa 6: Reiniciando serviço Plug and Play..." "INFO"
    try {
        Restart-Service -Name "PlugPlay" -Force -ErrorAction Stop
        Start-Sleep -Seconds 3
        Write-Log "Serviço Plug and Play reiniciado." "OK"
        $script:CorrecaoAplicada = $true
    } catch {
        Write-Log "Nao foi possivel reiniciar o serviço Plug and Play." "AVISO"
    }
}

# ============================================================
# ETAPA 3 — VALIDACAO POS-CORRECAO
# ============================================================
function Validar-Resultado {
    Write-Log "=== VALIDANDO RESULTADO ===" "INFO"
    Start-Sleep -Seconds 5

    # Re-verifica dock após correções
    $dockApos = $false
    try {
        $dock = Get-WmiObject -Namespace "root\dell\sysinv" `
                              -Class "dell_softwareidentity" `
                              -ErrorAction Stop |
                Where-Object { $_.ElementName -like "*Dock*" }
        if ($dock) { $dockApos = $true }
    } catch {}

    if (-not $dockApos) {
        $dockUSBApos = Get-PnpDevice -ErrorAction SilentlyContinue |
                       Where-Object { $_.FriendlyName -match "Dell.*Dock|Thunderbolt|WD19|WD22|TB16|TB18|UD22|D6000" } |
                       Sort-Object FriendlyName -Unique |
                       Where-Object { $_.Status -eq "OK" }
        if ($dockUSBApos) { $dockApos = $true }
    }

    # Re-verifica portas USB — ignora Unknown
    $usbProblemaApos = Get-PnpDevice -Class USB -ErrorAction SilentlyContinue |
                       Sort-Object FriendlyName -Unique |
                       Where-Object { $_.Status -ne "OK" -and $_.Status -ne "Unknown" }
    $portasUSBApos = if ($usbProblemaApos) { "PROBLEMA" } else { "OK" }

    # Re-verifica DisplayLink
    $displayLinkApos = Get-PnpDevice -ErrorAction SilentlyContinue |
                       Sort-Object FriendlyName -Unique |
                       Where-Object {
                           $_.FriendlyName -match "DisplayLink" -and
                           $_.Status -ne "OK" -and
                           $_.Status -ne "Unknown"
                       }
    $displayLinkStatusApos = if ($displayLinkApos) { "PROBLEMA" } else { "OK" }

    Write-Log "Dock detectada ANTES     : $(if ($script:DockDetectada) { 'Sim' } else { 'Nao' })" "INFO"
    Write-Log "Dock detectada APÓS      : $(if ($dockApos) { 'Sim' } else { 'Nao' })" "INFO"
    Write-Log "Portas USB ANTES         : $($script:PortasUSB)" "INFO"
    Write-Log "Portas USB APÓS          : $portasUSBApos" "INFO"
    Write-Log "DisplayLink ANTES        : $(if ($script:DisplayLinkProblema) { 'PROBLEMA' } else { 'OK' })" "INFO"
    Write-Log "DisplayLink APÓS         : $displayLinkStatusApos" "INFO"

    if ($dockApos -and $portasUSBApos -eq "OK" -and $displayLinkStatusApos -eq "OK" -and -not $script:ErrosEventViewer) {
        $script:ResultadoFinal = "SEM PROBLEMAS"
    } elseif ($dockApos -and ($script:DisplayLinkProblema -and $script:DisplayLinkCorrigido)) {
        $script:ResultadoFinal = "RESOLVIDO"
    } elseif ($script:CorrecaoAplicada -and $script:RebootNecessario) {
        $script:ResultadoFinal = "REBOOT_NECESSARIO"
    } elseif (-not $dockApos) {
        $script:ResultadoFinal = "PROVAVEL HARDWARE"
    } else {
        $script:ResultadoFinal = "INCONCLUSIVO"
    }
}

# ============================================================
# ETAPA 4 — RELATORIO FINAL + NOTIFICAÇÃO TEAMS
# ============================================================
function Exibir-Relatorio {
    Write-Log "============================================" "INFO"
    Write-Log "        RESULTADO DO DIAGNÓSTICO            " "INFO"
    Write-Log "============================================" "INFO"

    switch ($script:ResultadoFinal) {
        "SEM PROBLEMAS" {
            Write-Log "✅ SEM PROBLEMAS — Dock estavel, USB, video e DisplayLink OK." "OK"
        }
        "RESOLVIDO" {
            Write-Log "✅ RESOLVIDO — Problema no DisplayLink corrigido automaticamente." "OK"
        }
        "REBOOT_NECESSARIO" {
            Write-Log "🔄 REBOOT NECESSARIO — Firmware/drivers atualizados, aguardando reinicio." "AVISO"
            Write-Log "   Após reiniciar, reconecte o cabo de rede na Dock." "AVISO"
        }
        "PROVÁVEL HARDWARE" {
            Write-Log "🔴 PROVAVEL HARDWARE — Dock nao detectada apos todas as tentativas." "ERRO"
            Write-Log "   Proximos passos:" "ERRO"
            Write-Log "     1. Verificar cabo USB-C/Thunderbolt da Dock" "ERRO"
            Write-Log "     2. Testar Dock em outro computador Dell" "ERRO"
            Write-Log "     3. Verificar fonte de alimentacao da Dock" "ERRO"
            Write-Log "     4. Se falhar → acionar troca de hardware" "ERRO"
        }
        "INCONCLUSIVO" {
            Write-Log "⚠️  INCONCLUSIVO — Nao foi possível determinar a causa." "AVISO"
            if ($script:DisplayLinkProblema -and -not $script:DisplayLinkCorrigido) {
                Write-Log "   Driver DisplayLink com problema persistente." "AVISO"
                Write-Log "   → Acesse: dell.com/support e busque pelo modelo da Dock (ex: UD22)" "AVISO"
                Write-Log "   → Desinstale o driver DisplayLink atual pelo Gerenciador de Dispositivos" "AVISO"
                Write-Log "   → Instale o driver baixado e reinicie o notebook" "AVISO"
                Write-Log "   → Após reiniciar, reconecte o cabo de rede na Dock" "AVISO"
            } else {
                Write-Log "   Verifique os logs e escale para analise manual." "AVISO"
            }
        }
    }

    Write-Log "============================================" "INFO"

    $displayLinkStatusFinal = if ($script:DisplayLinkProblema) {
        if ($script:DisplayLinkCorrigido) { "Corrigido" } else { "Problema persistente" }
    } else { "OK" }

    Enviar-TeamsNotificacao `
        -Resultado        $script:ResultadoFinal `
        -Usuario          $env:USERNAME `
        -Computador       $env:COMPUTERNAME `
        -DockDetectada    $(if ($script:DockDetectada)    { "Sim" } else { "Nao" }) `
        -FirmwareDock     $script:FirmwareDock `
        -PortasUSB        $script:PortasUSB `
        -PortasVideo      $script:PortasVideo `
        -AlimentacaoDock  $script:AlimentacaoDock `
        -DisplayLinkStatus $displayLinkStatusFinal `
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