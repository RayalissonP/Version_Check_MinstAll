# ==========================
# CONFIGURAÇÃO
# ==========================
$IniFiles = @(
    "C:\Users\Rayalisson Pastor\Downloads\MInstall\minst.ini",
    "C:\Users\Rayalisson Pastor\Downloads\MInstall\profiles\z1.Portáteis.ini"
)

$AltUrlsPath = Join-Path $PSScriptRoot "alt_urls.json"
$AltUrls = @{}

if (Test-Path $AltUrlsPath) {
    $AltUrls = Get-Content $AltUrlsPath -Raw | ConvertFrom-Json
}

$Timeout = 20

# ==========================
# MAPA FIXO DE VERSÕES
# ==========================
$VersionMap = @{

    "github\.com"      = @{
        GetVersion = {
            param($url)
            $repo = ($url -replace 'https://github.com/', '').Split('/')[0..1] -join '/'
            $api = "https://api.github.com/repos/$repo/releases/latest"
            (Invoke-RestMethod $api).tag_name -replace '[^\d\.]'
        }
    }

    "fcportables\.com" = @{
        GetVersion = {
            param($html)
            if ($html -match '<title>.*?(\d+(\.\d+)+).*?</title>') {
                $matches[1]
            }
        }
    }

    "nirsoft\.net"     = @{
        GetVersion = {
            param($html)
            if ($html -match 'Version\s+([\d\.]+)') {
                $matches[1]
            }
        }
    }

    "sordum\.org"      = @{
        GetVersion = {
            param($html)
            if ($html -match 'v(\d+(\.\d+)+)') {
                $matches[1]
            }
        }
    }

    "majorgeeks\.com" = @{
        GetVersion = {
            param($html)

            if ($html -match '<h1[^>]*>[\s\S]*?([0-9]+(?:\.[0-9]+)+)[\s\S]*?</h1>') {
                return $matches[1]
            }
        }
    }

    "softwareok\.com" = @{
        GetVersion = {
            param($html)
            if ($html -match 'New in version\s+([\d\.]+)') {
                $matches[1]
            }
        }
    }

    "ccleaner\.com" = @{
        GetVersion = {
            param($html)
            if ($html -match 'v(\d+\.\d+\.\d+)') {
                $matches[1]
            }
        }
    }
}

# ==========================
# FUNÇÕES
# ==========================
function Normalize-AppName {
    param ($name)

    switch -Regex ($name) {
        'chrome' { return 'Google Chrome' }
        'firefox' { return 'Mozilla Firefox' }
        'adobe.*reader' { return 'Adobe Acrobat Reader DC MUI' }
        'anydesk' { return 'AnyDesk' }
        'k[- ]?lite' { return 'K-Lite Mega Codec Pack' }
        'java|jre' { return 'Java' }
        default { return $name }
    }
}

function Get-VersionFromAltUrls {
    param (
        [string]$Software
    )

    $key = $Software `
        -replace '®', '' `
        -replace '\(.*?\)', '' `
        -replace '\s{2,}', ' '

    $key = $key.Trim()

    if (-not $AltUrls.$key) {
        return $null
    }

    foreach ($src in $AltUrls.$key) {
        try {
            if ($src.Type -eq "text") {
                $r = Invoke-WebRequest $src.Url -UseBasicParsing -TimeoutSec $Timeout
                if ($r.Content.Trim()) {
                    return $r.Content.Trim()
                }
            }

            if ($src.Type -eq "html" -and $src.Regex) {
                $html = Invoke-WebRequest $src.Url -UseBasicParsing -TimeoutSec $Timeout
                if ($html.Content -match $src.Regex) {
                    return $matches[1]
                }
            }
        }
        catch { continue }
    }

    return $null
}

function Compare-VersionSafe {
    param ($local, $remote)
    try { [version]$local -lt [version]$remote }
    catch { $false }
}

function Get-IniApps {
    param ($path)

    $lines = Get-Content $path -Encoding Unicode
    $apps = @()
    $current = @{}

    foreach ($line in $lines) {
        if ($line -match '^\[\d+\]') {
            if ($current.Name) { $apps += [pscustomobject]$current }
            $current = @{}
        }

        if ($line -match '^Name=(.+)') { $current.Name = $matches[1] }
        if ($line -match '^Ver=(.+)') { $current.Ver = $matches[1] }
        if ($line -match '^URL=(.+)') { $current.URL = $matches[1] }
    }

    if ($current.Name) { $apps += [pscustomobject]$current }
    $apps
}

function Get-RemoteVersionFixed {
    param ($url)

    try {
        foreach ($key in $VersionMap.Keys) {
            if ($url -match $key) {
                if ($key -eq "github\.com") {
                    return & $VersionMap[$key].GetVersion $url
                }
                else {
                    $html = Invoke-WebRequest $url -UseBasicParsing -TimeoutSec $Timeout
                    return & $VersionMap[$key].GetVersion $html.Content
                }
            }
        }
        $null
    }
    catch {
        "ERRO"
    }
}

# ==========================
# PROCESSAMENTO
# ==========================
$result = @()

foreach ($ini in $IniFiles) {
    foreach ($app in (Get-IniApps $ini)) {

        if (-not $app.Ver -or -not $app.URL) { continue }

        $remote = Get-RemoteVersionFixed $app.URL
        if (-not $remote) {
            $normalized = Normalize-AppName $app.Name
            $remote = Get-VersionFromAltUrls $normalized
        }

        if (-not $remote) {
            $status = "VERIFICAÇÃO MANUAL"
        }
        elseif ($remote -eq "ERRO") {
            $status = "ERRO"
        }
        elseif (Compare-VersionSafe $app.Ver $remote) {
            $status = "DESATUALIZADO"
        }
        else {
            $status = "OK"
        }

        $result += [pscustomobject]@{
            Software      = $app.Name
            Versao_Local  = $app.Ver
            Versao_Remota = $remote
            Status        = $status
            URL           = $app.URL
            Origem_INI    = (Split-Path $ini -Leaf)
        }
    }
}

# ==========================
# LINHAS HTML
# ==========================
$rows = foreach ($r in $result) {

    $class = switch ($r.Status) {
        "OK" { "ok" }
        "DESATUALIZADO" { "old" }
        "VERIFICAÇÃO MANUAL" { "manual" }
        "ERRO" { "error" }
    }

    @"
<tr class="$class"
    data-status="$($r.Status)"
    data-app="$($r.Software)">
<td>$($r.Software)</td>
<td>$($r.Versao_Local)</td>
<td>$($r.Versao_Remota)</td>
<td>$($r.Status)</td>
<td><a href="$($r.URL)" target="_blank">Abrir</a></td>
<td>$($r.Origem_INI)</td>
<td><button onclick="toggleIgnore(this)">Ignorar</button></td>
</tr>
"@
}

# ==========================
# HTML FINAL
# ==========================
$html = @"
<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<title>Relatório MInstall</title>

<style>
body { font-family:Segoe UI,Arial; background:#f4f6f8; margin:20px }

.filters { margin-bottom:12px }
button { padding:6px 12px; margin:4px; border-radius:6px; border:0; cursor:pointer; font-weight:600 }
button.active { outline:2px solid #000 }

.ok{background:#d4edda}
.old{background:#fff3cd}
.manual{background:#d1ecf1}
.error{background:#f8d7da}
.ignored-btn{background:#ced4da}
.all{background:#e2e3e5}

table{width:100%; border-collapse:collapse; background:#fff}
th{background:#343a40; color:#fff; padding:8px}
td{padding:6px; border-bottom:1px solid #ddd}
.ignored{opacity:.4}
</style>
</head>

<body>

<h1>Relatório de Versões - MInstall</h1>

<div class="filters">
<button class="all active" onclick="setFilter('ALL',this)">
    Todos (<span id="count_all">0</span>)
</button>
<button class="ok" onclick="setFilter('OK',this)">
    OK (<span id="count_ok">0</span>)
</button>
<button class="old" onclick="setFilter('DESATUALIZADO',this)">
    Desatualizados (<span id="count_old">0</span>)
</button>
<button class="manual" onclick="setFilter('VERIFICAÇÃO MANUAL',this)">
    Manuais (<span id="count_manual">0</span>)
</button>
<button class="error" onclick="setFilter('ERRO',this)">
    Erros (<span id="count_error">0</span>)
</button>
<button class="ignored-btn" onclick="setFilter('IGNORED',this)">
    Ignorados (<span id="count_ignored">0</span>)
</button>
</div>

<table>
<tr>
<th>Software</th><th>Local</th><th>Remota</th>
<th>Status</th><th>URL</th><th>INI</th><th>Ação</th>
</tr>
$($rows -join "`n")
</table>

<script>
let currentFilter = 'ALL';
const ignored = new Set(JSON.parse(localStorage.getItem('ignoredApps') || '[]'));

function saveIgnored(){
 localStorage.setItem('ignoredApps', JSON.stringify([...ignored]));
}

function toggleIgnore(btn){
 const tr = btn.closest('tr');
 const app = tr.dataset.app;

 if(ignored.has(app)){
   ignored.delete(app);
   tr.classList.remove('ignored');
   btn.textContent='Ignorar';
 } else {
   ignored.add(app);
   tr.classList.add('ignored');
   btn.textContent='Ativar';
 }
 saveIgnored();
 updateCounters();
 applyFilter();
}

function setFilter(f,btn){
 currentFilter=f;
 document.querySelectorAll('.filters button').forEach(b=>b.classList.remove('active'));
 btn.classList.add('active');
 applyFilter();
}

function applyFilter(){
 document.querySelectorAll('tr[data-status]').forEach(tr=>{
   const ig = ignored.has(tr.dataset.app);
   const st = tr.dataset.status;

   let show =
     currentFilter==='ALL' ? !ig :
     currentFilter==='IGNORED' ? ig :
     (!ig && st===currentFilter);

   tr.style.display = show ? '' : 'none';
 });
}

function updateCounters(){
 let c={ALL:0,OK:0,DESATUALIZADO:0,"VERIFICAÇÃO MANUAL":0,ERRO:0};

 document.querySelectorAll('tr[data-status]').forEach(tr=>{
   if(!ignored.has(tr.dataset.app)){
     c.ALL++;
     c[tr.dataset.status]++;
   }
 });

 count_all.textContent=c.ALL;
 count_ok.textContent=c.OK;
 count_old.textContent=c.DESATUALIZADO;
 count_manual.textContent=c["VERIFICAÇÃO MANUAL"];
 count_error.textContent=c.ERRO;
 count_ignored.textContent=ignored.size;
}

document.querySelectorAll('tr[data-status]').forEach(tr=>{
 if(ignored.has(tr.dataset.app)){
   tr.classList.add('ignored');
   tr.querySelector('button').textContent='Ativar';
 }
});

updateCounters();
applyFilter();
</script>

</body>
</html>
"@

# ==========================
# SALVAR E ABRIR
# ==========================
$path = Join-Path $PSScriptRoot "Relatorio_MInstall.html"
$html | Set-Content $path -Encoding UTF8 -Force
Start-Process $path
