param(
    [switch]$MostrarMedios
)

$inicio = (Get-Date).AddHours(-48)

$filtros = @(
    @{LogName='System'; Id=41},
    @{LogName='System'; Id=1001},
    @{LogName='System'; Id=7034},
    @{LogName='System'; Id=7031},
    @{LogName='System'; Id=51},
    @{LogName='System'; Id=219},
    @{LogName='System'; Id=157},
    @{LogName='System'; Id=9}
)

$descricoes = @{
    41   = "Desligamento Inesperado"
    1001 = "BSOD"
    7034 = "Servico Parou Inesperadamente"
    7031 = "Servico Falhou"
    51   = "Erro de Disco"
    219  = "Driver com Problema (Hardware Externo)"
    157  = "Disco Removido Inesperadamente"
    9    = "Timeout de Controlador de Dispositivo"
}

$sugestoes = @{
    41   = @(
        "1. Verifique estabilidade da fonte de alimentacao",
        "2. Checar logs de temperatura (superaquecimento)",
        "3. Executar script de limpeza e correcao da imagem",
        "4. Verificar atualizacoes pendentes do Windows"
    )
    1001 = @(
        "1. Analisar Minidump em: C:\WINDOWS\Minidump\ (via WinDbg)",
        "2. Executar script de limpeza e correcao da imagem",
        "3. Verificar drivers desatualizados",
        "4. Testar memoria RAM: mdsched.exe"
    )
    7034 = @(
        "1. Verificar qual servico falhou no Event Viewer",
        "2. Tentar reiniciar o servico: Restart-Service -Name <nome>",
        "3. Checar dependencias do servico",
        "4. Verificar permissoes da conta de servico"
    )
    7031 = @(
        "1. Verificar configuracao de recuperacao do servico",
        "2. Checar logs especificos do servico",
        "3. Reinstalar o servico se necessario"
    )
    51   = @(
        "1. Executar: chkdsk /f /r",
        "2. Verificar saude do disco: Get-PhysicalDisk | Get-StorageReliabilityCounter",
        "3. Fazer backup imediato dos dados",
        "4. Considerar substituicao do disco se erro persistir"
    )
    219  = @(
        "1. Verificar se o driver do dispositivo esta atualizado",
        "2. Reconectar o dispositivo externo",
        "3. Testar em outra porta USB/Thunderbolt",
        "4. Verificar compatibilidade do driver com a versao do Windows"
    )
    157  = @(
        "1. Verificar cabo e conexao do dispositivo",
        "2. Testar em outra porta",
        "3. Verificar se o disco externo esta com problemas fisicos",
        "4. Executar chkdsk no disco externo"
    )
    9    = @(
        "1. Verificar cabo e conexao da dock/dispositivo",
        "2. Atualizar firmware da dock station",
        "3. Testar o dispositivo em outro computador",
        "4. Verificar drivers do controlador USB/Thunderbolt"
    )
}

$criticos  = @(41, 1001, 51)
$medios    = @(7034, 7031)
$hardware  = @(219, 157, 9)

# ─── Funções de verificação de status ───────────────────────────────────────

function Verificar-Status41 {
    param ($dataEvento)

    $bootLimpo = Get-WinEvent -FilterHashtable @{
        LogName   = 'System'
        Id        = 12
        StartTime = $dataEvento
    } -ErrorAction SilentlyContinue | Select-Object -First 1

    if ($bootLimpo) {
        Write-Host "  Status atual: ✅ Resolvido" -ForegroundColor Green
    } else {
        Write-Host "  Status atual: ⚠️  Pendente - Nenhum boot limpo detectado apos o evento" -ForegroundColor Yellow
    }
}

function Verificar-Status1001 {
    param ($dataEvento)

    $minidumpPath = "C:\Windows\Minidump"
    if (Test-Path $minidumpPath) {
        $dumps = Get-ChildItem -Path $minidumpPath -Filter "*.dmp" |
                 Where-Object { $_.LastWriteTime -ge $dataEvento } |
                 Sort-Object LastWriteTime -Descending

        if ($dumps) {
            Write-Host "  Status atual: ⚠️  Pendente - $($dumps.Count) Minidump(s) aguardando analise via WinDbg:" -ForegroundColor Yellow
            foreach ($dump in $dumps) {
                Write-Host "     - $($dump.Name) ($($dump.LastWriteTime.ToString('dd/MM/yyyy HH:mm')))" -ForegroundColor Yellow
            }
        } else {
            Write-Host "  Status atual: ✅ Resolvido" -ForegroundColor Green
        }
    } else {
        Write-Host "  Status atual: ✅ Resolvido" -ForegroundColor Green
    }
}

function Verificar-Status51 {
    try {
        $discos = Get-PhysicalDisk | Get-StorageReliabilityCounter -ErrorAction Stop
        $comErro = $discos | Where-Object { $_.ReadErrorsTotal -gt 0 -or $_.WriteErrorsTotal -gt 0 }

        if ($comErro) {
            Write-Host "  Status atual: ⚠️  Pendente - Erros de disco ainda detectados:" -ForegroundColor Yellow
            foreach ($disco in $comErro) {
                Write-Host "     - Leitura: $($disco.ReadErrorsTotal) erro(s) | Escrita: $($disco.WriteErrorsTotal) erro(s)" -ForegroundColor Yellow
            }
        } else {
            Write-Host "  Status atual: ✅ Resolvido" -ForegroundColor Green
        }
    } catch {
        Write-Host "  Status atual: ⚠️  Pendente - Nao foi possivel verificar saude do disco (requer privilegios de administrador)" -ForegroundColor Yellow
    }
}

function Verificar-StatusHardware {
    param ($eventoID, $dataEvento)

    $eventoRecente = Get-WinEvent -FilterHashtable @{
        LogName   = 'System'
        Id        = $eventoID
        StartTime = (Get-Date).AddHours(-1)
    } -ErrorAction SilentlyContinue | Select-Object -First 1

    if ($eventoRecente) {
        Write-Host "  Status atual: ⚠️  Pendente - Problema ainda recorrente na ultima hora" -ForegroundColor Yellow
    } else {
        Write-Host "  Status atual: ✅ Resolvido" -ForegroundColor Green
    }
}

function Obter-StatusAtual {
    param ($eventoID, $dataEvento)

    switch ($eventoID) {
        41             { Verificar-Status41   -dataEvento $dataEvento }
        1001           { Verificar-Status1001 -dataEvento $dataEvento }
        51             { Verificar-Status51 }
        {$_ -in $hardware} { Verificar-StatusHardware -eventoID $eventoID -dataEvento $dataEvento }
    }
}

# ─── Coleta de eventos ───────────────────────────────────────────────────────

$resultados = @()

foreach ($filtro in $filtros) {
    try {
        $eventos = Get-WinEvent -FilterHashtable @{
            LogName   = $filtro.LogName
            Id        = $filtro.Id
            StartTime = $inicio
        } -ErrorAction SilentlyContinue

        foreach ($evento in $eventos) {
            $resultados += [PSCustomObject]@{
                Data      = $evento.TimeCreated.ToString("dd/MM/yyyy")
                Hora      = $evento.TimeCreated.ToString("HH:mm")
                DataHora  = $evento.TimeCreated
                EventoID  = $evento.Id
                Tipo      = $descricoes[$evento.Id]
                Descricao = $evento.Message.Split("`n")[0]
                Nivel     = if ($criticos -contains $evento.Id) { "CRITICO" }
                            elseif ($hardware -contains $evento.Id) { "HARDWARE" }
                            else { "MEDIO" }
            }
        }
    } catch {}
}

# ─── Exibição ────────────────────────────────────────────────────────────────

$hardwarePendente = $false

function Exibir-Evento {
    param ($grupo, $cor, $verificarStatus)

    $primeiro    = $grupo.Group[0]
    $ocorrencias = $grupo.Count
    $label       = if ($ocorrencias -gt 1) { "($ocorrencias x)" } else { "" }

    Write-Host "  [$($primeiro.Hora)] $label $($primeiro.Tipo) (ID: $($primeiro.EventoID))" -ForegroundColor $cor
    Write-Host "  $($primeiro.Descricao)" -ForegroundColor $cor
    Write-Host "  Sugestoes de Correcao:" -ForegroundColor Cyan
    foreach ($sugestao in $sugestoes[$primeiro.EventoID]) {
        Write-Host "     $sugestao" -ForegroundColor Cyan
    }

    if ($verificarStatus) {
        $statusOutput = Obter-StatusAtual -eventoID $primeiro.EventoID -dataEvento $primeiro.DataHora

        if ($hardware -contains $primeiro.EventoID -and $statusOutput -like "*Pendente*") {
            $script:hardwarePendente = $true
        }
    }

    Write-Host ""
}

$resultadosCriticos = $resultados | Where-Object { $_.Nivel -eq "CRITICO" }
$resultadosMedios   = $resultados | Where-Object { $_.Nivel -eq "MEDIO" }
$resultadosHardware = $resultados | Where-Object { $_.Nivel -eq "HARDWARE" }

if ($resultadosCriticos.Count -eq 0 -and $resultadosHardware.Count -eq 0 -and (-not $MostrarMedios -or $resultadosMedios.Count -eq 0)) {
    Write-Host "`nNenhum erro encontrado nas ultimas 48h.`n" -ForegroundColor Green
} else {

    if ($resultadosCriticos.Count -gt 0) {
        Write-Host "`n=== ERROS CRITICOS - ULTIMAS 48H ===`n" -ForegroundColor Red

        $porDia = $resultadosCriticos | Sort-Object Data | Group-Object -Property Data

        foreach ($dia in $porDia) {
            Write-Host "========================================" -ForegroundColor Cyan
            Write-Host "  $($dia.Name)" -ForegroundColor Cyan
            Write-Host "========================================" -ForegroundColor Cyan

            $dia.Group |
                Sort-Object Hora |
                Group-Object EventoID, Descricao |
                ForEach-Object { Exibir-Evento -grupo $_ -cor "Red" -verificarStatus $true }
        }
    }

    if ($resultadosHardware.Count -gt 0) {
        Write-Host "`n=== HARDWARE EXTERNO - ULTIMAS 48H ===`n" -ForegroundColor Magenta

        $porDia = $resultadosHardware | Sort-Object Data | Group-Object -Property Data

        foreach ($dia in $porDia) {
            Write-Host "========================================" -ForegroundColor Cyan
            Write-Host "  $($dia.Name)" -ForegroundColor Cyan
            Write-Host "========================================" -ForegroundColor Cyan

            $dia.Group |
                Sort-Object Hora |
                Group-Object EventoID, Descricao |
                ForEach-Object { Exibir-Evento -grupo $_ -cor "red" -verificarStatus $true }
        }
    }

    if ($MostrarMedios -and $resultadosMedios.Count -gt 0) {
        Write-Host "`n=== ERROS MEDIOS - ULTIMAS 48H ===`n" -ForegroundColor Yellow

        $porDia = $resultadosMedios | Sort-Object Data | Group-Object -Property Data

        foreach ($dia in $porDia) {
            Write-Host "========================================" -ForegroundColor Cyan
            Write-Host "  $($dia.Name)" -ForegroundColor Cyan
            Write-Host "========================================" -ForegroundColor Cyan

            $dia.Group |
                Sort-Object Hora |
                Group-Object EventoID, Descricao |
                ForEach-Object { Exibir-Evento -grupo $_ -cor "Yellow" -verificarStatus $false }
        }
    }

    $totalCriticos = $resultadosCriticos.Count
    $totalMedios   = $resultadosMedios.Count
    $totalHardware = $resultadosHardware.Count

    Write-Host "========================================" -ForegroundColor White
    Write-Host "  RESUMO" -ForegroundColor White
    Write-Host "========================================" -ForegroundColor White
    Write-Host "  CRITICOS         : $totalCriticos evento(s)" -ForegroundColor Red
    Write-Host "  HARDWARE EXTERNO : $totalHardware evento(s)" -ForegroundColor red
    if ($MostrarMedios) {
        Write-Host "  MEDIOS           : $totalMedios evento(s)" -ForegroundColor Yellow
    }
    Write-Host "========================================`n" -ForegroundColor White

    if ($hardwarePendente) {
        Write-Host "========================================" -ForegroundColor Yellow
        Write-Host "  ACAO RECOMENDADA" -ForegroundColor Yellow
        Write-Host "========================================" -ForegroundColor Yellow
        Write-Host "  ⚠️  Erros de hardware externo detectados e pendentes." -ForegroundColor Yellow
        Write-Host "  Execute o script de correcao de hardware:" -ForegroundColor Yellow
        Write-Host "  .\CorrecaoHardware.ps1" -ForegroundColor Cyan
        Write-Host "========================================`n" -ForegroundColor Yellow
    }
}