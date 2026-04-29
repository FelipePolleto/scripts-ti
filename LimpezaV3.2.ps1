# Garante codificação UTF8 e pede privilégios de Admin automaticamente
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
 Start-Process PowerShell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
 exit
}

# ===============================================
# VARIÁVEIS GERAIS DO DELL COMMAND UPDATE
# ===============================================
$dcuPaths = @(
    "C:\Program Files\Dell\CommandUpdate\dcu-cli.exe",
    "C:\Program Files (x86)\Dell\CommandUpdate\dcu-cli.exe"
)
$dcuCli = $dcuPaths | Where-Object { Test-Path $_ } | Select-Object -First 1

$dcuArgs = "/applyUpdates -updateType=driver,firmware -silent"

# ===============================================
# FUNÇÕES DO SISTEMA
# ===============================================

# Função para parar serviços
function Parar-ServicoUpdate {
 Write-Host "`nEncerrando servicos conflitantes..." -ForegroundColor Red
 $processos = @("TiWorker", "TrustedInstaller")
 foreach ($proc in $processos) {
  Stop-Process -Name $proc -Force -ErrorAction SilentlyContinue
 }
 $servicos = @("wuauserv", "bits", "cryptsvc")
 foreach ($servico in $servicos) {
  $svc = Get-CimInstance Win32_Service -Filter "Name='$servico'" -ErrorAction SilentlyContinue
  if ($svc -and $svc.State -eq 'Running') {
   if ($svc.ProcessId -gt 0) {
    Stop-Process -Id $svc.ProcessId -Force -ErrorAction SilentlyContinue
   } else {
    sc.exe stop $servico | Out-Null
   }
  }
 }
 Start-Sleep -Seconds 3
}

# Função de Status (linha da tabela)
function Show-TableRow {
 param($Label, $Value, $IsCritical)
 $Color = if ($IsCritical) { "Red" } else { "Green" }
 $Icon = if ($IsCritical) { "[!!]" } else { "[OK]" }
 Write-Host "`t" -NoNewline -ForegroundColor Cyan
 Write-Host ("{0,-20}" -f $Label) -NoNewline
 Write-Host " `t" -NoNewline -ForegroundColor Cyan
 Write-Host ("{0,-6}" -f $Icon) -ForegroundColor $Color -NoNewline
 Write-Host " `t" -NoNewline -ForegroundColor Cyan
 Write-Host ("{0,-20}" -f $Value) -NoNewline
 Write-Host " " -ForegroundColor Cyan
}

# Função de Tabela de Status do Sistema
function Show-StatusTable {
    Write-Host "`n=========================================" -ForegroundColor Cyan
    Write-Host "       STATUS ATUAL DO SISTEMA           " -ForegroundColor White
    Write-Host "=========================================" -ForegroundColor Cyan

    # --- UPTIME ---
    $os = Get-CimInstance Win32_OperatingSystem
    $uptime = (Get-Date) - $os.LastBootUpTime
    $uptimeStr = "{0}d {1}h {2}m" -f $uptime.Days, $uptime.Hours, $uptime.Minutes
    $uptimeCritico = $uptime.Days -ge 7

    # --- TOP CPU ---
    $topCpu = Get-Process | Sort-Object CPU -Descending | Select-Object -First 1
    $topCpuStr = "$($topCpu.ProcessName) ($([math]::Round($topCpu.CPU, 1))s)"
    $cpuCritico = $topCpu.CPU -gt 300

    # --- TOP RAM (processo) ---
    $topRam = Get-Process | Sort-Object WorkingSet64 -Descending | Select-Object -First 1
    $topRamMB = [math]::Round($topRam.WorkingSet64 / 1MB, 1)
    $topRamStr = "$($topRam.ProcessName) ($topRamMB MB)"
    $ramCritico = $topRamMB -gt 1500

    # --- USO TOTAL DE RAM (apenas %) ---
    $totalRam  = $os.TotalVisibleMemorySize
    $freeRam   = $os.FreePhysicalMemory
    $ramPct    = [math]::Round((($totalRam - $freeRam) / $totalRam) * 100, 0)
    $ramUsoStr = "$ramPct% em uso"
    $ramUsoCritico = $ramPct -gt 85

    # --- TOP DISCO (via Get-Counter, mais confiável) ---
    try {
        $diskCounters = Get-Counter '\Process(*)\IO Data Bytes/sec' -ErrorAction Stop
        $topDiskSample = $diskCounters.CounterSamples |
            Where-Object { $_.InstanceName -notin @('_total', 'idle') -and $_.CookedValue -gt 0 } |
            Sort-Object CookedValue -Descending |
            Select-Object -First 1

        if ($topDiskSample) {
            $topDiskName = $topDiskSample.InstanceName
            $topDiskMB   = [math]::Round($topDiskSample.CookedValue / 1MB, 2)
            $topDiskStr  = "$topDiskName ($topDiskMB MB/s)"
        } else {
            $topDiskFallback = Get-Process |
                Sort-Object { $_.ReadOperationCount + $_.WriteOperationCount } -Descending |
                Select-Object -First 1
            $topDiskStr = if ($topDiskFallback) { "$($topDiskFallback.ProcessName) (I/O acumulado)" } else { "N/D" }
        }
    } catch {
        $topDiskFallback = Get-Process |
            Sort-Object { $_.ReadOperationCount + $_.WriteOperationCount } -Descending |
            Select-Object -First 1
        $topDiskStr = if ($topDiskFallback) { "$($topDiskFallback.ProcessName) (I/O acumulado)" } else { "N/D" }
    }

    # --- ESPAÇO EM DISCO C: (apenas %) ---
    $disco        = Get-PSDrive -Name C
    $discoTotalGB = $disco.Used + $disco.Free
    $discoPct     = [math]::Round(($disco.Used / $discoTotalGB) * 100, 0)
    $discoStr     = "$discoPct% em uso"
    $discoCritico = $discoPct -gt 85

    # --- SAÚDE DA BATERIA ---
    $bateria = Get-CimInstance Win32_Battery -ErrorAction SilentlyContinue
    if ($bateria) {
        $cargaAtual = $bateria.EstimatedChargeRemaining
        $statusBat  = switch ($bateria.BatteryStatus) {
            1 { "Descarregando" }
            2 { "Conectada (CA)" }
            3 { "Carregando" }
            4 { "Baixa" }
            5 { "Critica" }
            default { "Desconhecido" }
        }
        $saudeStr   = "$cargaAtual% — $statusBat"
        $batCritico = $cargaAtual -lt 20 -and $bateria.BatteryStatus -eq 1
    } else {
        $saudeStr   = "Sem bateria / Desktop"
        $batCritico = $false
    }

    # --- EXIBIÇÃO DA TABELA ---
    $sep = "-" * 65
    Write-Host "`t$sep" -ForegroundColor DarkCyan
    Write-Host ("`t{0,-20} {1,-6}  {2}" -f "ITEM", "STATUS", "DETALHE") -ForegroundColor White
    Write-Host "`t$sep" -ForegroundColor DarkCyan

    Show-TableRow "Uptime"             $uptimeStr    $uptimeCritico
    Show-TableRow "Top CPU"            $topCpuStr    $cpuCritico
    Show-TableRow "Maior uso de RAM"   $topRamStr    $ramCritico
    Show-TableRow "Uso Total de RAM"   $ramUsoStr    $ramUsoCritico
	Show-TableRow "Maior uso Disco"    $topDiskStr   $false
    Show-TableRow "Disco C: (Espaco)"  $discoStr     $discoCritico  
    Show-TableRow "Bateria"            $saudeStr     $batCritico

    Write-Host "`t$sep" -ForegroundColor DarkCyan
    Write-Host "`n`t[!!] = Atencao recomendada   [OK] = Normal`n" -ForegroundColor DarkGray
}

# Função Finalização
function Finalizar-Script {
 Write-Host "`n===============================================================" -ForegroundColor Cyan
 Write-Host "Procedimentos concluidos!" -ForegroundColor Green
 Write-Host "Recomenda-se reiniciar o computador para aplicar todas as alteracoes (incluindo atualizacoes de Firmware/Drivers)." -ForegroundColor Yellow
 $respReboot = Read-Host "Deseja reiniciar a maquina agora? (S/N)"
 if ($respReboot -match '^[sS]$') {
  Write-Host "Reiniciando o sistema..." -ForegroundColor Red
  Restart-Computer
 } else {
  Write-Host "Saindo do processo. Lembre-se de reiniciar manualmente mais tarde." -ForegroundColor Gray
  Start-Sleep -Seconds 3
 }
}


# ===============================================
# INÍCIO DO MENU
# ===============================================
do {
 Clear-Host
 Show-StatusTable
 Write-Host "=========================================" -ForegroundColor Cyan
 Write-Host " SCRIPT DE MANUTENCAO DO TI " -ForegroundColor White
 Write-Host "=========================================" -ForegroundColor Cyan
 Write-Host "1. Fazer APENAS Atualizacao (Dell)"
 Write-Host "2. Fazer APENAS Limpeza e Reparo"
 Write-Host "3. Fazer Limpeza E Atualizacao (Completo)"
 Write-Host "0. Sair"
 Write-Host "=========================================" -ForegroundColor Cyan
 $opcao = Read-Host "Escolha uma opcao (0-3)"

 switch ($opcao) {

# -------------------------------------------------------
# 1 - ATUALIZAÇÃO
# -------------------------------------------------------
 '1' {
 Write-Host "`n[+] Iniciando apenas Atualizacao..." -ForegroundColor Green

 if ($dcuCli) {
  Write-Host "`nIniciando Dell Command Update..." -ForegroundColor Cyan
  Write-Host "Buscando e instalando atualizacoes (APENAS Drivers e Firmwares). Aguarde..." -ForegroundColor Yellow

  $process = Start-Process -FilePath $dcuCli -ArgumentList $dcuArgs -Wait -PassThru

  switch ($process.ExitCode) {
      0 { Write-Host "[Ok] Atualizações de Driver/Firmware instaladas com sucesso!" -ForegroundColor Green }
      1 { Write-Host "[Ok] Atualizações instaladas, reinício necessário." -ForegroundColor Yellow }
      2 { Write-Host "[Info] Nenhuma atualização de Driver/Firmware pendente." -ForegroundColor Cyan }
      3 { Write-Host "[Ok] Atualizações aplicadas — reinício necessário." -ForegroundColor Yellow }
      default { Write-Host "[!] Código inesperado: $($process.ExitCode)" -ForegroundColor Red }
  }

  Write-Host "`nVerificando se há atualizações de outros tipos..." -ForegroundColor Cyan
  
  Start-Process -FilePath $dcuCli -ArgumentList "/scan -silent" -Wait -NoNewWindow
  
  $pendingXml = "C:\ProgramData\Dell\CommandUpdate\Status\Pending.xml"
  $outrasAtualizacoes = @()

  if (Test-Path $pendingXml) {
      try {
          [xml]$pend = Get-Content $pendingXml
          $outrasAtualizacoes = $pend.Updates.Component | Select-Object -ExpandProperty Name
      } catch {}
  }

  if ($outrasAtualizacoes.Count -gt 0) {
      Write-Host "`n[!] Atualizações pendentes não necessárias:" -ForegroundColor Yellow
      foreach ($p in $outrasAtualizacoes) {
          Write-Host " → $p" -ForegroundColor Gray
      }
  } else {
      Write-Host "`n[Ok] Nenhuma atualização restante." -ForegroundColor Green
  }

 } else {
  Write-Host "Dell Command Update nao encontrado neste computador." -ForegroundColor Red
 }

 Finalizar-Script
 }

# -------------------------------------------------------
# 2 - LIMPEZA
# -------------------------------------------------------
 '2' {
 Write-Host "`n[+] Iniciando apenas Limpeza..." -ForegroundColor Green
 Write-Host ">> Limpando temporarios e cache DNS..." -NoNewline
 $null = Remove-Item -Path "C:\Windows\Temp\*" -Recurse -Force -ErrorAction SilentlyContinue
 $null = Remove-Item -Path "$env:TEMP\*" -Recurse -Force -ErrorAction SilentlyContinue
 $null = Clear-RecycleBin -Confirm:$false -ErrorAction SilentlyContinue
 ipconfig /flushdns | Out-Null
 Write-Host " [Concluido]" -ForegroundColor Green

 $respReparo = Read-Host "`nExecutar reparos de sistema (SFC/DISM)? (S/N)"
 if ($respReparo -match '^[sS]$') {
     Parar-ServicoUpdate
     Write-Host "Iniciando SFC..." -ForegroundColor Yellow
     $sfcRun = sfc /scannow | Out-String
     Write-Host "Verificando imagem do Windows (DISM)..." -ForegroundColor Cyan
     $checkHealth = dism.exe /online /cleanup-image /checkhealth | Out-String

     if ($checkHealth -notmatch "Nenhuma corrupcao" -and $checkHealth -notmatch "No component store corruption") {
         Start-Process -FilePath "dism.exe" -ArgumentList "/online /cleanup-image /restorehealth" -Wait -NoNewWindow
     }

     Start-Service -Name "wuauserv" -ErrorAction SilentlyContinue
     Start-Service -Name "bits" -ErrorAction SilentlyContinue
     Start-Service -Name "cryptsvc" -ErrorAction SilentlyContinue
 }
 Finalizar-Script
 }

# -------------------------------------------------------
# 3 - LIMPEZA + ATUALIZAÇÃO
# -------------------------------------------------------
 '3' {
 Write-Host "`n[+] Iniciando Limpeza e Atualizacao Completa..." -ForegroundColor Green

 Write-Host "`n>> Limpando temporarios e cache DNS..." -NoNewline
 $null = Remove-Item -Path "C:\Windows\Temp\*" -Recurse -Force -ErrorAction SilentlyContinue
 $null = Remove-Item -Path "$env:TEMP\*" -Recurse -Force -ErrorAction SilentlyContinue
 $null = Clear-RecycleBin -Confirm:$false -ErrorAction SilentlyContinue
 ipconfig /flushdns | Out-Null
 Write-Host " [Concluido]" -ForegroundColor Green

 $respReparo = Read-Host "`nExecutar reparos de sistema (SFC/DISM)? (S/N)"
 if ($respReparo -match '^[sS]$') {
     Parar-ServicoUpdate
     sfc /scannow | Out-Null
     Start-Process -FilePath "dism.exe" -ArgumentList "/online /cleanup-image /restorehealth" -Wait -NoNewWindow
 }

 if ($dcuCli) {
  Write-Host "`nIniciando Dell Command Update..." -ForegroundColor Cyan
  Write-Host "Buscando e instalando atualizacoes (APENAS Drivers e Firmwares). Aguarde..." -ForegroundColor Yellow

  $process = Start-Process -FilePath $dcuCli -ArgumentList $dcuArgs -Wait -PassThru

  switch ($process.ExitCode) {
      0 { Write-Host "[Ok] Atualizações de Driver/Firmware instaladas com sucesso!" -ForegroundColor Green }
      1 { Write-Host "[Ok] Atualizações instaladas, reinício necessário." -ForegroundColor Yellow }
      2 { Write-Host "[Info] Nenhuma atualização de Driver/Firmware pendente." -ForegroundColor Cyan }
      3 { Write-Host "[Ok] Atualizações aplicadas — reinício necessário." -ForegroundColor Yellow }
      default { Write-Host "[!] Código inesperado: $($process.ExitCode)" -ForegroundColor Red }
  }

  Write-Host "`nVerificando se há atualizações de outros tipos..." -ForegroundColor Cyan
  
  Start-Process -FilePath $dcuCli -ArgumentList "/scan -silent" -Wait -NoNewWindow
  
  $pendingXml = "C:\ProgramData\Dell\CommandUpdate\Status\Pending.xml"
  $outrasAtualizacoes = @()

  if (Test-Path $pendingXml) {
      try {
          [xml]$pend = Get-Content $pendingXml
          $outrasAtualizacoes = $pend.Updates.Component | Select-Object -ExpandProperty Name
      } catch {}
  }

  if ($outrasAtualizacoes.Count -gt 0) {
      Write-Host "`n[!] Atualizações pendentes não necessárias:" -ForegroundColor Yellow
      foreach ($p in $outrasAtualizacoes) {
          Write-Host " → $p" -ForegroundColor Gray
      }
  } else {
      Write-Host "`n[Ok] Nenhuma atualização restante." -ForegroundColor Green
  }

 } else {
  Write-Host "Dell Command Update nao encontrado." -ForegroundColor Red
 }

 Finalizar-Script
 }

# -------------------------------------------------------
# 0 - SAIR
# -------------------------------------------------------
 '0' {
 Write-Host "`nSaindo..." -ForegroundColor Yellow
 Start-Sleep -Seconds 1
 }

# DEFAULT
 default {
 Write-Host "`nOpcao invalida! Digite 0 a 3." -ForegroundColor Red
 Start-Sleep -Seconds 2
 }
 }
} while ($opcao -ne '0')