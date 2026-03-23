# =============================================================================
# DIAGNÓSTICO TÉCNICO: NDD
# Versão: 2.0
# Descrição: Testa conectividade TCP/UDP entre servidor NDD e impressoras. 
# =============================================================================

# ----------------------------- CONFIGURAÇÕES ---------------------------------

$ipDestino = "10.101.8.19"
$timeoutMs  = 2000  # Timeout por porta em milissegundos

# IP de origem: filtra loopback e endereços automáticos (APIPA/WellKnown)
$ipOrigem = (
    Get-NetIPAddress -AddressFamily IPv4 |
    Where-Object {
        $_.InterfaceAlias -notlike "*Loopback*" -and
        $_.PrefixOrigin    -ne "WellKnown"
    } |
    Select-Object -First 1
).IPAddress

# Pasta e arquivo de log (criado automaticamente se não existir)
$logDir  = "C:\Logs\NDD"
$logFile = Join-Path $logDir "diagnostico_ndd_$(Get-Date -Format 'yyyyMMdd_HHmm').txt"

if (-not (Test-Path $logDir)) {
    New-Item -ItemType Directory -Path $logDir | Out-Null
}

# ---------------------- MAPA DE PORTAS E PROTOCOLOS --------------------------

# Formato: "porta" = @{ Descricao = "..."; Protocolo = "TCP" | "UDP" }
$mapaPortas = [ordered]@{
    "80"    = @{ Descricao = "HTTP  - Acesso ao Web Image Monitor";                         Protocolo = "TCP" }
    "443"   = @{ Descricao = "HTTPS - Comunicação segura com o painel da impressora";       Protocolo = "TCP" }
    "161"   = @{ Descricao = "SNMP  - Coleta de contadores, suprimentos e status";          Protocolo = "UDP" }
    "9100"  = @{ Descricao = "RAW   - Porta padrão de envio de trabalhos de impressão";     Protocolo = "TCP" }
    "5656"  = @{ Descricao = "DCS   - Data Communication Service (protocolo NDD)";          Protocolo = "TCP" }
    "56562" = @{ Descricao = "NDD   - Monitoramento via porta alta (Inbound)";              Protocolo = "TCP" }
    "56563" = @{ Descricao = "NDD   - Monitoramento via porta alta (Outbound)";             Protocolo = "TCP" }
}

# ----------------------------- FUNÇÕES ---------------------------------------

function Escrever-Log {
    param([string]$Mensagem)
    $Mensagem | Out-File -FilePath $logFile -Append -Encoding UTF8
}

function Testar-PortaTCP {
    param([string]$IP, [int]$Porta, [int]$Timeout)

    try {
        $tcp = New-Object System.Net.Sockets.TcpClient
        $conectou = $tcp.ConnectAsync($IP, $Porta).Wait($Timeout)
        $tcp.Close()
        return $conectou
    } catch {
        return $false
    }
}

function Escrever-Linha {
    param(
        [string]$Porta,
        [string]$Status,
        [string]$Cor,
        [string]$Protocolo,
        [string]$Descricao
    )

    $prefixo = "  Porta $Porta".PadRight(12) + "[$Protocolo]".PadRight(7) + " : "

    Write-Host $prefixo -NoNewline
    Write-Host $Status.PadRight(12) -ForegroundColor $Cor -NoNewline
    Write-Host "| $Descricao"

    # Linha para log (sem cores ANSI)
    Escrever-Log ("  Porta $Porta".PadRight(12) + "[$Protocolo]".PadRight(7) + " : " + $Status.PadRight(12) + "| $Descricao")
}

# ----------------------------- CABEÇALHO ------------------------------------

$linha  = "=" * 70
$inicio = Get-Date -Format "dd/MM/yyyy HH:mm:ss"

$cabecalho = @"
$linha
  DIAGNÓSTICO TÉCNICO: NDD
  Autor: Rafael Zonta
  Data/Hora  : $inicio
  IP Origem  : $ipOrigem
  IP Destino : $ipDestino
  Log salvo  : $logFile
$linha
"@

Write-Host $cabecalho -ForegroundColor Cyan
Escrever-Log $cabecalho

# ----------------------------- TESTES ----------------------------------------

$abertas    = [System.Collections.Generic.List[string]]::new()
$bloqueadas = [System.Collections.Generic.List[string]]::new()
$ignoradas  = [System.Collections.Generic.List[string]]::new()  # UDP (não testável via TCP)

foreach ($porta in $mapaPortas.Keys) {

    $info      = $mapaPortas[$porta]
    $protocolo = $info.Protocolo
    $descricao = $info.Descricao

    # Portas UDP não podem ser testadas via TCP — registrar como N/T (não testável)
    if ($protocolo -eq "UDP") {
        Escrever-Linha -Porta $porta -Status "NÃO TESTÁVEL" -Cor "Yellow" -Protocolo $protocolo -Descricao $descricao
        $ignoradas.Add($porta)
        continue
    }

    $aberta = Testar-PortaTCP -IP $ipDestino -Porta ([int]$porta) -Timeout $timeoutMs

    if ($aberta) {
        Escrever-Linha -Porta $porta -Status "ABERTA" -Cor "Green" -Protocolo $protocolo -Descricao $descricao
        $abertas.Add($porta)
    } else {
        Escrever-Linha -Porta $porta -Status "BLOQUEADA" -Cor "Red" -Protocolo $protocolo -Descricao $descricao
        $bloqueadas.Add($porta)
    }
}

# ----------------------------- RESUMO ----------------------------------------

$rodape = @"
$linha
  RESUMO DOS TESTES
$linha
  Abertas      : $($abertas.Count)   [ $($abertas -join ", ") ]
  Bloqueadas   : $($bloqueadas.Count)   [ $($bloqueadas -join ", ") ]
  Não testadas : $($ignoradas.Count)   [ $($ignoradas -join ", ") ]  (UDP — verificar via SNMP Walk)
$linha
  OBSERVAÇÕES
$linha
  [!] Porta 161 usa UDP. Mesmo que o teste TCP falhe, o SNMP pode estar
      funcionando. Valide via SNMP Walk ou pelo Web Image Monitor da Ricoh.

  [!] Portas bloqueadas podem indicar regra de firewall, ACL no switch,
      ou serviço não iniciado na impressora/servidor NDD.

  [!] Timeout configurado: ${timeoutMs}ms por porta.
$linha
  Log completo salvo em: $logFile
$linha
"@

Write-Host $rodape -ForegroundColor Cyan
Escrever-Log $rodape

Write-Host "  Script finalizado em: $(Get-Date -Format 'HH:mm:ss')" -ForegroundColor White
Escrever-Log "  Script finalizado em: $(Get-Date -Format 'HH:mm:ss')"
