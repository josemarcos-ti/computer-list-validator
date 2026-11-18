<#
.SYNOPSIS
    Valida uma lista de computadores em paralelo: online, IP, dominio, usuario logado,
    sistema, tempo ligado e espaco livre - e gera HTML interativo + CSV.

.DESCRIPTION
    Para cada computador da lista (arquivo texto, um por linha, ou -ComputerName):
      - resolve o nome no DNS e faz ping
      - le pela rede (CIM/WinRM, com tentativa via DCOM): dominio, usuario logado,
        sistema operacional, ultimo boot e espaco livre no C:
    Os testes rodam em paralelo (runspaces), entao listas grandes terminam rapido.
    O resultado mantem a ordem original (numerada) e e classificado em:
        Online       - respondeu e foi lido
        Sem acesso   - responde na rede, mas a leitura remota falhou (permissao/firewall/WinRM)
        Offline      - nao respondeu
        Sem DNS      - o nome nao existe no DNS

.PARAMETER Arquivo
    Arquivo texto com um computador por linha (linhas vazias e # sao ignoradas).
.PARAMETER ComputerName
    Lista de computadores direto no comando (alternativa ao -Arquivo).
.PARAMETER Credential
    Credencial administrativa para a leitura remota (opcional; padrao: conta atual).
.PARAMETER Paralelo
    Quantos computadores testar ao mesmo tempo (padrao: 25).
.PARAMETER Saida
    Pasta dos arquivos gerados (padrao: pasta atual).
.PARAMETER Demo
    Gera o relatorio com dados ficticios (para testar/visualizar).
.PARAMETER NaoAbrir
    Nao abre o HTML ao terminar.

.EXAMPLE
    .\Test-ComputerList.ps1 -Arquivo .\computadores.txt
.EXAMPLE
    .\Test-ComputerList.ps1 -ComputerName PC01, PC02, SRV-ARQ01 -Credential (Get-Credential)
.EXAMPLE
    .\Test-ComputerList.ps1 -Demo
#>
[CmdletBinding(DefaultParameterSetName = 'Arquivo')]
param(
    [Parameter(ParameterSetName = 'Arquivo', Position = 0)][string]$Arquivo,
    [Parameter(ParameterSetName = 'Lista')][string[]]$ComputerName,
    [pscredential]$Credential,
    [ValidateRange(1, 100)][int]$Paralelo = 25,
    [string]$Saida = (Get-Location).Path,
    [Parameter(ParameterSetName = 'Demo')][switch]$Demo,
    [switch]$NaoAbrir
)

$ErrorActionPreference = 'Stop'
$agora = Get-Date

# ---------------- teste de UM computador (roda em paralelo) ----------------
$teste = {
    param($nome, $cred)
    $r = [ordered]@{ Nome = $nome; IP = ''; Ping = $false; Dominio = ''; Usuario = ''; Sistema = ''; LigadoDesde = ''; DiasLigado = ''; LivreC_GB = ''; Leitura = ''; Erro = '' }
    try { $ips = @([System.Net.Dns]::GetHostAddresses($nome) | Where-Object { $_.AddressFamily -eq 'InterNetwork' }) } catch { $ips = @() }
    if (-not $ips.Count) { $r.Erro = 'nome nao encontrado no DNS'; return [pscustomobject]$r }
    $r.IP = $ips[0].IPAddressToString
    try { $r.Ping = ((New-Object System.Net.NetworkInformation.Ping).Send($nome, 1500).Status -eq 'Success') } catch { }
    $erros = @()
    foreach ($proto in 'Wsman', 'Dcom') {
        $s = $null
        try {
            $prm = @{ ComputerName = $nome; SessionOption = (New-CimSessionOption -Protocol $proto); OperationTimeoutSec = 15; ErrorAction = 'Stop' }
            if ($cred) { $prm['Credential'] = $cred }
            $s = New-CimSession @prm
            $cs = Get-CimInstance -CimSession $s -ClassName Win32_ComputerSystem
            $os = Get-CimInstance -CimSession $s -ClassName Win32_OperatingSystem
            $dk = Get-CimInstance -CimSession $s -ClassName Win32_LogicalDisk -Filter "DeviceID='C:'"
            $r.Dominio = "$($cs.Domain)"; $r.Sistema = "$($os.Caption)".Replace('Microsoft ', '')
            $r.Usuario = "$($cs.UserName)"
            if (-not $r.Usuario) {
                # sem usuario no console: dono do explorer.exe (sessao remota / troca rapida)
                $donos = foreach ($p in @(Get-CimInstance -CimSession $s -ClassName Win32_Process -Filter "Name='explorer.exe'")) {
                    try { $o = Invoke-CimMethod -InputObject $p -MethodName GetOwner; if ($o.User) { "$($o.Domain)\$($o.User)" } } catch { }
                }
                $r.Usuario = (@($donos | Select-Object -Unique) -join ', ')
            }
            if ($os.LastBootUpTime) { $r.LigadoDesde = $os.LastBootUpTime.ToString('yyyy-MM-dd HH:mm'); $r.DiasLigado = [int]((Get-Date) - $os.LastBootUpTime).TotalDays }
            if ($dk) { $r.LivreC_GB = [math]::Round($dk.FreeSpace / 1GB, 1) }
            $r.Leitura = $proto
            Remove-CimSession $s
            return [pscustomobject]$r
        } catch {
            $erros += "${proto}: $($_.Exception.Message.Split("`n")[0])"
            if ($s) { Remove-CimSession $s -ErrorAction SilentlyContinue }
        }
    }
    $r.Erro = $erros -join ' | '
    [pscustomobject]$r
}

# ---------------- lista de computadores ----------------
$nomes = New-Object System.Collections.Generic.List[string]
if ($Demo) { }
elseif ($ComputerName) { foreach ($n in $ComputerName) { if ($n.Trim()) { $nomes.Add($n.Trim().ToUpper()) } } }
elseif ($Arquivo) {
    foreach ($l in (Get-Content -LiteralPath $Arquivo)) {
        $n = ($l.Trim() -split '[\s,;]+')[0]
        if ($n -and -not $n.StartsWith('#')) { $nomes.Add($n.ToUpper()) }
    }
}
else { throw 'Informe -Arquivo, -ComputerName ou -Demo.' }

# ---------------- execucao ----------------
$res = @{}
if ($Demo) {
    Write-Host 'Modo DEMO: gerando dados ficticios...' -ForegroundColor Yellow
    $rnd = New-Object System.Random 7
    $us = 'ana.souza', 'bruno.lima', 'carla.mendes', 'diego.rocha', 'elisa.prado', 'fabio.nunes', 'gabriela.reis', ''
    for ($i = 1; $i -le 30; $i++) {
        $n = '{0}-PC{1:000}' -f @('FIN', 'RH', 'COM', 'TI', 'JUR')[$rnd.Next(5)], $i; $nomes.Add($n)
        $k = $rnd.Next(10)
        $o = [ordered]@{ Nome = $n; IP = "10.0.$($rnd.Next(1,9)).$($rnd.Next(10,250))"; Ping = $true; Dominio = 'empresa.local'; Usuario = "EMPRESA\$($us[$rnd.Next($us.Count)])".TrimEnd('\'); Sistema = @('Windows 11 Pro', 'Windows 10 Pro')[$rnd.Next(2)]; LigadoDesde = ''; DiasLigado = $rnd.Next(0, 40); LivreC_GB = [math]::Round(5 + $rnd.NextDouble() * 180, 1); Leitura = 'Wsman'; Erro = '' }
        $o.LigadoDesde = $agora.AddDays(-$o.DiasLigado).ToString('yyyy-MM-dd HH:mm')
        if ($o.Usuario -eq 'EMPRESA') { $o.Usuario = '' }
        if ($k -eq 0) { $o.Ping = $false; foreach ($c in 'Dominio', 'Usuario', 'Sistema', 'LigadoDesde', 'DiasLigado', 'LivreC_GB', 'Leitura') { $o[$c] = '' }; $o.Erro = 'tempo esgotado' }
        elseif ($k -eq 1) { foreach ($c in 'Dominio', 'Usuario', 'Sistema', 'LigadoDesde', 'DiasLigado', 'LivreC_GB', 'Leitura') { $o[$c] = '' }; $o.Erro = 'Acesso negado' }
        elseif ($k -eq 2 -and $i % 3 -eq 0) { $o.IP = ''; $o.Ping = $false; foreach ($c in 'Dominio', 'Usuario', 'Sistema', 'LigadoDesde', 'DiasLigado', 'LivreC_GB', 'Leitura') { $o[$c] = '' }; $o.Erro = 'nome nao encontrado no DNS' }
        $res[$n] = [pscustomobject]$o
    }
}
else {
    $unicos = @($nomes | Select-Object -Unique)
    Write-Host "$($unicos.Count) computador(es), $Paralelo em paralelo..." -ForegroundColor Cyan
    $pool = [runspacefactory]::CreateRunspacePool(1, $Paralelo); $pool.Open()
    $jobs = New-Object System.Collections.Generic.List[object]
    foreach ($n in $unicos) {
        $ps = [powershell]::Create().AddScript($teste).AddArgument($n).AddArgument($Credential)
        $ps.RunspacePool = $pool
        $jobs.Add([pscustomobject]@{ Nome = $n; PS = $ps; H = $ps.BeginInvoke() })
    }
    $feitos = 0; $inicio = Get-Date
    while ($jobs.Count) {
        foreach ($j in $jobs.ToArray()) {
            if (-not $j.H.IsCompleted) { continue }
            try { $res[$j.Nome] = @($j.PS.EndInvoke($j.H))[0] } catch { $res[$j.Nome] = [pscustomobject]@{ Nome = $j.Nome; Erro = $_.Exception.Message } }
            $j.PS.Dispose(); [void]$jobs.Remove($j); $feitos++
            Write-Progress -Activity 'Testando computadores' -Status "$feitos de $($unicos.Count)" -PercentComplete ($feitos * 100 / $unicos.Count)
        }
        if (((Get-Date) - $inicio).TotalMinutes -gt (5 + $unicos.Count / 10)) {
            foreach ($j in $jobs.ToArray()) { try { $j.PS.Stop() } catch { }; $res[$j.Nome] = [pscustomobject]@{ Nome = $j.Nome; Erro = 'tempo esgotado' }; $j.PS.Dispose() }
            $jobs.Clear()
        }
        Start-Sleep -Milliseconds 200
    }
    $pool.Close(); $pool.Dispose()
    Write-Progress -Activity 'Testando computadores' -Completed
}

# ---------------- resultado na ordem da lista ----------------
$linhas = New-Object System.Collections.Generic.List[object]
$i = 0
foreach ($n in $nomes) {
    $i++; $r = $res[$n]
    $sit = if ($r.Leitura) { 'Online' } elseif ("$($r.Erro)" -match 'DNS') { 'Sem DNS' } elseif ($r.Ping) { 'Sem acesso' } else { 'Offline' }
    $linhas.Add([pscustomobject]@{
            N = $i; Computador = $n; Situacao = $sit; IP = "$($r.IP)"; Dominio = "$($r.Dominio)"; Usuario = "$($r.Usuario)"
            Sistema = "$($r.Sistema)"; LigadoDesde = "$($r.LigadoDesde)"; DiasLigado = "$($r.DiasLigado)"; LivreC_GB = "$($r.LivreC_GB)"
            Detalhe = $(if ($r.Erro) { "$($r.Erro)" } elseif ($r.Ping) { "lido via $($r.Leitura)" } else { "lido via $($r.Leitura) (sem resposta ao ping)" })
        })
}
$arr = $linhas.ToArray()
if (-not (Test-Path $Saida)) { New-Item -ItemType Directory -Path $Saida -Force | Out-Null }
$stamp = $agora.ToString('yyyyMMdd-HHmmss')
$base = Join-Path $Saida "Validacao-Computadores-$stamp"
$arr | Export-Csv -Path "$base.csv" -Delimiter ';' -NoTypeInformation -Encoding UTF8

$meta = [ordered]@{
    titulo      = 'Validacao de computadores'
    subtitulo   = "Gerado em $($agora.ToString('dd/MM/yyyy HH:mm')) - $($arr.Count) computador(es)$(if ($Demo) { ' - dados ficticios (modo demo)' }) - conta: $(if ($Credential) { $Credential.UserName } else { "$env:USERDOMAIN\$env:USERNAME" })"
    rodape      = 'N = posicao na lista original. Clique nos indicadores para filtrar e nos titulos das colunas para ordenar.'
    arquivo     = "Validacao-Computadores-$stamp"
    campoStatus = 'Situacao'
    ordemStatus = @('Online', 'Sem acesso', 'Offline', 'Sem DNS')
    cores       = @{ 'Online' = '#15803d'; 'Sem acesso' = '#b45309'; 'Offline' = '#64748b'; 'Sem DNS' = '#9f1239' }
    numericas   = @('N', 'DiasLigado', 'LivreC_GB')
    ordem       = 'N'
    colunas     = @(
        @{ k = 'N'; t = '#' }, @{ k = 'Computador'; t = 'Computador'; mono = $true }, @{ k = 'Situacao'; t = 'Situacao' }, @{ k = 'IP'; t = 'IP' },
        @{ k = 'Dominio'; t = 'Dominio' }, @{ k = 'Usuario'; t = 'Usuario logado' }, @{ k = 'Sistema'; t = 'Sistema' },
        @{ k = 'DiasLigado'; t = 'Dias ligado' }, @{ k = 'LivreC_GB'; t = 'Livre no C: (GB)' }, @{ k = 'Detalhe'; t = 'Detalhe' }
    )
}

# ---------------- modelo do relatorio HTML ----------------
$Modelo = @'
<!doctype html>
<html lang="pt-br"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Relatorio</title>
<style>
:root{--ink:#1f2933;--mut:#627d98;--line:#e4e9f0;--bg:#f2f5f9;--pri:#1d4ed8}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--ink);font-family:"Segoe UI",Roboto,Arial,sans-serif}
header{background:linear-gradient(120deg,#0f2b5b,#1d4ed8);color:#fff;padding:24px 36px}
header h1{margin:0;font-size:23px;font-weight:700}header .sub{opacity:.85;font-size:13px;margin-top:6px}
main{padding:22px 36px;max-width:1400px;margin:0 auto}
.cards{display:grid;grid-template-columns:repeat(auto-fit,minmax(170px,1fr));gap:12px;margin-bottom:18px}
.card{background:#fff;border-radius:12px;padding:14px 16px;cursor:pointer;box-shadow:0 1px 4px rgba(15,43,91,.08);border-left:6px solid #cbd5e1;transition:transform .1s}
.card:hover{transform:translateY(-2px)}.card.on{outline:3px solid var(--pri)}
.card .n{font-size:30px;font-weight:800;line-height:1}.card .l{font-size:11.5px;color:var(--mut);margin-top:7px;text-transform:uppercase;letter-spacing:.04em}
.panel{background:#fff;border-radius:12px;padding:16px 18px;box-shadow:0 1px 4px rgba(15,43,91,.06)}
.tools{display:flex;flex-wrap:wrap;gap:10px;align-items:center;margin-bottom:12px}
.tools input{flex:1;min-width:260px;max-width:420px;padding:9px 12px;border:1px solid #cbd5e1;border-radius:8px;font-size:14px}
button{background:var(--pri);color:#fff;border:0;border-radius:8px;padding:9px 14px;font-size:13px;cursor:pointer}button.sec{background:#64748b}
.cnt{color:var(--mut);font-size:13px;margin-left:auto}
.scroll{max-height:70vh;overflow:auto;border-radius:8px;border:1px solid var(--line)}
table{border-collapse:collapse;width:100%;font-size:13px}
th{background:#0f2b5b;color:#fff;text-align:left;padding:9px 10px;position:sticky;top:0;cursor:pointer;white-space:nowrap}
th:after{content:" \2195";opacity:.4;font-size:11px}
td{padding:8px 10px;border-top:1px solid var(--line);vertical-align:top}
tr:nth-child(even) td{background:#fafbfd}tr:hover td{background:#eef4ff}
td.mono{font-family:Consolas,monospace;font-weight:600}
.tag{display:inline-block;padding:2px 9px;border-radius:12px;font-size:11.5px;font-weight:700;color:#fff;white-space:nowrap}
.foot{color:var(--mut);font-size:12px;margin-top:14px}
@media print{.tools,.cards{display:none}.scroll{max-height:none;overflow:visible;border:0}th{position:static}}
</style></head><body>
<header><h1 id="tit"></h1><div class="sub" id="sub"></div></header>
<main>
<div class="cards" id="cards"></div>
<div class="panel">
<div class="tools"><input id="q" placeholder="Filtrar em qualquer coluna..."><button onclick="baixarCsv()">Baixar CSV</button><button class="sec" onclick="window.print()">Imprimir / PDF</button><span class="cnt" id="cnt"></span></div>
<div class="scroll"><table id="tb"><thead></thead><tbody></tbody></table></div>
</div>
<div class="foot" id="foot"></div>
</main>
<script>
var META = __META__;
var DATA = __DATA__;
var filtro = '', ord = { k: META.ordem || META.colunas[0].k, d: META.ordemDir || 1 };
function e(s) { return (s == null ? '' : String(s)).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }
function cor(s) { return (META.cores && META.cores[s]) || '#64748b'; }
function lista() {
  var q = document.getElementById('q').value.toLowerCase();
  var L = DATA.filter(function (r) {
    if (filtro && r[META.campoStatus] !== filtro) return false;
    if (!q) return true;
    return META.colunas.some(function (c) { return String(r[c.k] == null ? '' : r[c.k]).toLowerCase().indexOf(q) >= 0; });
  });
  var num = (META.numericas || []);
  L.sort(function (a, b) {
    var x = a[ord.k], y = b[ord.k];
    if (num.indexOf(ord.k) >= 0) { x = (x === '' || x == null) ? 1e9 : +x; y = (y === '' || y == null) ? 1e9 : +y; }
    else { x = String(x == null ? '' : x).toLowerCase(); y = String(y == null ? '' : y).toLowerCase(); }
    return x > y ? ord.d : x < y ? -ord.d : 0;
  });
  return L;
}
function cards() {
  var cont = {}; DATA.forEach(function (r) { var s = r[META.campoStatus]; cont[s] = (cont[s] || 0) + 1; });
  var h = '<div class="card on" data-s="" style="border-left-color:#1d4ed8"><div class="n" style="color:#1d4ed8">' + DATA.length + '</div><div class="l">Todos</div></div>';
  (META.ordemStatus || Object.keys(cont)).forEach(function (s) {
    h += '<div class="card" data-s="' + e(s) + '" style="border-left-color:' + cor(s) + '"><div class="n" style="color:' + cor(s) + '">' + (cont[s] || 0) + '</div><div class="l">' + e(s) + '</div></div>';
  });
  var el = document.getElementById('cards'); el.innerHTML = h;
  el.querySelectorAll('.card').forEach(function (c) {
    c.onclick = function () { el.querySelectorAll('.card').forEach(function (x) { x.classList.remove('on'); }); c.classList.add('on'); filtro = c.getAttribute('data-s'); desenha(); };
  });
}
function desenha() {
  var L = lista();
  document.querySelector('#tb thead').innerHTML = '<tr>' + META.colunas.map(function (c) { return '<th data-k="' + c.k + '">' + e(c.t) + '</th>'; }).join('') + '</tr>';
  document.querySelectorAll('#tb th').forEach(function (th) { th.onclick = function () { var k = th.getAttribute('data-k'); ord.d = (ord.k === k ? -ord.d : 1); ord.k = k; desenha(); }; });
  document.querySelector('#tb tbody').innerHTML = L.map(function (r) {
    return '<tr>' + META.colunas.map(function (c) {
      var v = r[c.k];
      if (c.k === META.campoStatus) return '<td><span class="tag" style="background:' + cor(v) + '">' + e(v) + '</span></td>';
      return '<td' + (c.mono ? ' class="mono"' : '') + '>' + e(v) + '</td>';
    }).join('') + '</tr>';
  }).join('') || '<tr><td colspan="' + META.colunas.length + '">Nenhum registro.</td></tr>';
  document.getElementById('cnt').textContent = L.length + ' de ' + DATA.length + ' registro(s)';
}
function baixarCsv() {
  var L = lista();
  var s = '﻿' + META.colunas.map(function (c) { return c.t; }).join(';') + '\r\n' + L.map(function (r) {
    return META.colunas.map(function (c) { return '"' + String(r[c.k] == null ? '' : r[c.k]).replace(/"/g, '""') + '"'; }).join(';');
  }).join('\r\n');
  var a = document.createElement('a'); a.href = URL.createObjectURL(new Blob([s], { type: 'text/csv;charset=utf-8' })); a.download = (META.arquivo || 'relatorio') + '.csv'; a.click();
}
document.getElementById('tit').textContent = META.titulo;
document.getElementById('sub').textContent = META.subtitulo;
document.getElementById('foot').textContent = META.rodape || '';
document.title = META.titulo;
document.getElementById('q').oninput = desenha;
cards(); desenha();
</script>
</body></html>

'@

# monta o HTML (modelo abaixo + dados em JSON)
$jsMeta = ConvertTo-Json -InputObject $meta -Depth 5 -Compress
$jsData = if ($arr.Count) { ConvertTo-Json -InputObject $arr -Depth 3 -Compress } else { '[]' }
$html = $Modelo.Replace('__META__', ($jsMeta -replace '</', '<\/')).Replace('__DATA__', ($jsData -replace '</', '<\/'))
[IO.File]::WriteAllText("$base.html", $html, (New-Object System.Text.UTF8Encoding($true)))

Write-Host ''
$arr | Group-Object -Property $meta.campoStatus | Sort-Object Count -Descending | ForEach-Object { Write-Host ("  {0,-14} {1}" -f $_.Name, $_.Count) }
Write-Host "`nHTML: $base.html" -ForegroundColor Green
Write-Host "CSV : $base.csv" -ForegroundColor Green
if (-not $NaoAbrir) { Start-Process "$base.html" }
