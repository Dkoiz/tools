#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Windows Server Post-Deployment Setup - HTML GUI
.DESCRIPTION
    Launches a local web UI for configuring Network, Hostname, and Domain Join
    after a fresh Windows Server install from a VMware template.
    Run as Administrator. Opens in your default browser.
.NOTES
    Listens on http://127.0.0.1:8199/ - only accessible locally.
    Close the browser tab and press Ctrl+C in the console to stop.
#>

$Port = 8199
$Prefix = "http://127.0.0.1:$Port/"

# ----------------------------------------------
#  HELPER FUNCTIONS
# ----------------------------------------------

function Get-AdapterListJson {
    $adapters = Get-NetAdapter | Where-Object { $_.Status -eq 'Up' -or $_.Status -eq 'Disconnected' } |
        Select-Object Name, InterfaceIndex, Status, MacAddress
    $result = @()
    foreach ($a in $adapters) {
        $ip  = Get-NetIPAddress -InterfaceIndex $a.InterfaceIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue | Select-Object -First 1
        $gw  = Get-NetRoute -InterfaceIndex $a.InterfaceIndex -DestinationPrefix '0.0.0.0/0' -ErrorAction SilentlyContinue | Select-Object -First 1
        $dns = Get-DnsClientServerAddress -InterfaceIndex $a.InterfaceIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue
        $result += @{
            name         = $a.Name
            index        = $a.InterfaceIndex
            status       = $a.Status
            mac          = $a.MacAddress
            ip           = if ($ip) { $ip.IPAddress } else { '' }
            prefix       = if ($ip) { $ip.PrefixLength } else { 24 }
            gateway      = if ($gw) { $gw.NextHop } else { '' }
            dns1         = if ($dns -and $dns.ServerAddresses.Count -gt 0) { $dns.ServerAddresses[0] } else { '' }
            dns2         = if ($dns -and $dns.ServerAddresses.Count -gt 1) { $dns.ServerAddresses[1] } else { '' }
            dhcp         = if ($ip) { $ip.PrefixOrigin -eq 'Dhcp' } else { $true }
        }
    }
    return ($result | ConvertTo-Json -Depth 3)
}

function Flush-AdapterNetwork ([int]$InterfaceIndex) {
    Get-NetIPAddress -InterfaceIndex $InterfaceIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue |
        Remove-NetIPAddress -Confirm:$false -ErrorAction SilentlyContinue
    Get-NetRoute -InterfaceIndex $InterfaceIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue |
        Remove-NetRoute -Confirm:$false -ErrorAction SilentlyContinue
    Set-DnsClientServerAddress -InterfaceIndex $InterfaceIndex -ResetServerAddresses -ErrorAction SilentlyContinue
    Set-NetIPInterface -InterfaceIndex $InterfaceIndex -Dhcp Disabled -ErrorAction SilentlyContinue
    Start-Sleep -Milliseconds 500
}

function MaskToPrefix ([string]$mask) {
    $bits = 0
    foreach ($o in ($mask -split '\.')) {
        $bits += [Convert]::ToString([int]$o, 2).Replace('0','').Length
    }
    return $bits
}

function PrefixToMask ([int]$prefix) {
    if ($prefix -le 0) { return '0.0.0.0' }
    if ($prefix -ge 32) { return '255.255.255.255' }
    [uint32]$raw = ([math]::Pow(2, 32) - [math]::Pow(2, 32 - $prefix))
    $o1 = [math]::Floor($raw / 16777216) -band 0xFF
    $o2 = [math]::Floor($raw / 65536) -band 0xFF
    $o3 = [math]::Floor($raw / 256) -band 0xFF
    $o4 = $raw -band 0xFF
    return ('{0}.{1}.{2}.{3}' -f $o1, $o2, $o3, $o4)
}

# ----------------------------------------------
#  HTML PAGE
# ----------------------------------------------

$HTML = @'
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Server Post-Deployment Setup</title>
<style>
  @import url('https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@400;500&family=IBM+Plex+Sans:wght@300;400;500;600&display=swap');

  * { margin: 0; padding: 0; box-sizing: border-box; }

  :root {
    --bg: #0c0e14;
    --surface: #141620;
    --surface2: #1a1d2c;
    --border: #2a2f44;
    --border-focus: #4a7cff;
    --text: #d8dce8;
    --text-dim: #6b7394;
    --accent: #4a7cff;
    --accent-glow: rgba(74, 124, 255, 0.15);
    --success: #34d399;
    --success-bg: rgba(52, 211, 153, 0.1);
    --warn: #fbbf24;
    --danger: #f87171;
    --danger-bg: rgba(248, 113, 113, 0.08);
    --mono: 'JetBrains Mono', 'Consolas', monospace;
    --sans: 'IBM Plex Sans', -apple-system, sans-serif;
  }

  body {
    font-family: var(--sans);
    background: var(--bg);
    color: var(--text);
    min-height: 100vh;
    line-height: 1.5;
  }

  /* --- Subtle grid background --- */
  body::before {
    content: '';
    position: fixed;
    inset: 0;
    background-image:
      linear-gradient(rgba(74, 124, 255, 0.03) 1px, transparent 1px),
      linear-gradient(90deg, rgba(74, 124, 255, 0.03) 1px, transparent 1px);
    background-size: 40px 40px;
    pointer-events: none;
    z-index: 0;
  }

  .container {
    max-width: 680px;
    margin: 0 auto;
    padding: 40px 24px 60px;
    position: relative;
    z-index: 1;
  }

  /* --- Header --- */
  .header {
    margin-bottom: 36px;
    padding-bottom: 24px;
    border-bottom: 1px solid var(--border);
  }
  .header h1 {
    font-size: 22px;
    font-weight: 600;
    letter-spacing: -0.3px;
    color: var(--text);
    margin-bottom: 6px;
  }
  .header h1 span { color: var(--accent); }
  .header p {
    font-size: 13px;
    color: var(--text-dim);
    font-weight: 300;
  }
  .hostname-badge {
    display: inline-block;
    margin-top: 10px;
    padding: 4px 12px;
    background: var(--surface2);
    border: 1px solid var(--border);
    border-radius: 4px;
    font-family: var(--mono);
    font-size: 12px;
    color: var(--warn);
  }

  /* --- Sections --- */
  .section {
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: 8px;
    padding: 24px;
    margin-bottom: 20px;
    transition: border-color 0.2s;
  }
  .section:hover { border-color: #353a54; }

  .section-head {
    display: flex;
    align-items: center;
    gap: 10px;
    margin-bottom: 20px;
  }
  .section-num {
    width: 28px;
    height: 28px;
    display: flex;
    align-items: center;
    justify-content: center;
    border-radius: 6px;
    background: var(--accent-glow);
    color: var(--accent);
    font-size: 13px;
    font-weight: 600;
    flex-shrink: 0;
  }
  .section-title {
    font-size: 15px;
    font-weight: 600;
    letter-spacing: -0.2px;
  }

  /* --- Form elements --- */
  .field {
    margin-bottom: 14px;
  }
  .field label {
    display: block;
    font-size: 12px;
    font-weight: 500;
    color: var(--text-dim);
    margin-bottom: 5px;
    text-transform: uppercase;
    letter-spacing: 0.5px;
  }
  .field input, .field select {
    width: 100%;
    padding: 10px 14px;
    background: var(--bg);
    border: 1px solid var(--border);
    border-radius: 6px;
    color: var(--text);
    font-family: var(--mono);
    font-size: 13px;
    transition: border-color 0.2s, box-shadow 0.2s;
    outline: none;
  }
  .field input:focus, .field select:focus {
    border-color: var(--border-focus);
    box-shadow: 0 0 0 3px var(--accent-glow);
  }
  .field input:disabled {
    opacity: 0.35;
    cursor: not-allowed;
  }
  .field select {
    cursor: pointer;
    -webkit-appearance: none;
    appearance: none;
    background-image: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='12' fill='%236b7394'%3E%3Cpath d='M6 8L1 3h10z'/%3E%3C/svg%3E");
    background-repeat: no-repeat;
    background-position: right 12px center;
    padding-right: 36px;
  }
  .field select option {
    background: var(--surface);
    color: var(--text);
  }
  .field .hint {
    font-size: 11px;
    color: var(--text-dim);
    margin-top: 4px;
    font-style: italic;
  }

  .row {
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 12px;
  }

  /* --- Toggle / Checkbox --- */
  .toggle-row {
    display: flex;
    align-items: center;
    gap: 10px;
    margin-bottom: 18px;
    padding: 10px 14px;
    background: var(--surface2);
    border-radius: 6px;
    cursor: pointer;
    user-select: none;
    transition: background 0.15s;
  }
  .toggle-row:hover { background: #1f2336; }
  .toggle-row input[type=checkbox] { display: none; }
  .toggle-track {
    width: 38px;
    height: 20px;
    background: var(--border);
    border-radius: 10px;
    position: relative;
    transition: background 0.2s;
    flex-shrink: 0;
  }
  .toggle-track::after {
    content: '';
    width: 16px;
    height: 16px;
    background: #555;
    border-radius: 50%;
    position: absolute;
    top: 2px;
    left: 2px;
    transition: transform 0.2s, background 0.2s;
  }
  .toggle-row input:checked + .toggle-track {
    background: var(--accent);
  }
  .toggle-row input:checked + .toggle-track::after {
    transform: translateX(18px);
    background: #fff;
  }
  .toggle-label {
    font-size: 13px;
    font-weight: 400;
  }

  /* --- Buttons --- */
  .btn {
    display: inline-flex;
    align-items: center;
    justify-content: center;
    gap: 6px;
    padding: 10px 22px;
    border: 1px solid transparent;
    border-radius: 6px;
    font-family: var(--sans);
    font-size: 13px;
    font-weight: 500;
    cursor: pointer;
    transition: all 0.15s;
    outline: none;
  }
  .btn:active { transform: scale(0.97); }
  .btn-primary {
    background: var(--accent);
    color: #fff;
  }
  .btn-primary:hover { background: #5d8aff; }
  .btn-primary:disabled {
    opacity: 0.5;
    cursor: not-allowed;
    transform: none;
  }
  .btn-success {
    background: var(--success);
    color: #0c0e14;
  }
  .btn-success:hover { background: #4ee3a9; }
  .btn-danger {
    background: var(--danger-bg);
    color: var(--danger);
    border-color: rgba(248, 113, 113, 0.2);
  }
  .btn-danger:hover { background: rgba(248, 113, 113, 0.15); }
  .btn-ghost {
    background: transparent;
    color: var(--text-dim);
    border-color: var(--border);
  }
  .btn-ghost:hover { border-color: #454b6a; color: var(--text); }

  .btn-row {
    display: flex;
    gap: 10px;
    margin-top: 20px;
  }

  /* --- Status toast --- */
  .toast {
    position: fixed;
    bottom: 24px;
    right: 24px;
    max-width: 420px;
    padding: 14px 20px;
    border-radius: 8px;
    font-size: 13px;
    font-weight: 500;
    box-shadow: 0 8px 32px rgba(0,0,0,0.5);
    transform: translateY(120%);
    opacity: 0;
    transition: transform 0.3s ease, opacity 0.3s ease;
    z-index: 999;
  }
  .toast.show {
    transform: translateY(0);
    opacity: 1;
  }
  .toast.ok {
    background: #0f2a1e;
    border: 1px solid rgba(52, 211, 153, 0.3);
    color: var(--success);
  }
  .toast.err {
    background: #2a0f0f;
    border: 1px solid rgba(248, 113, 113, 0.3);
    color: var(--danger);
  }
  .toast.info {
    background: #0f1a2a;
    border: 1px solid rgba(74, 124, 255, 0.3);
    color: var(--accent);
  }

  /* --- Spinner --- */
  .spinner {
    display: inline-block;
    width: 14px;
    height: 14px;
    border: 2px solid rgba(255,255,255,0.2);
    border-top-color: #fff;
    border-radius: 50%;
    animation: spin 0.6s linear infinite;
  }
  @keyframes spin { to { transform: rotate(360deg); } }

  /* --- Footer --- */
  .footer {
    display: flex;
    justify-content: space-between;
    align-items: center;
    margin-top: 12px;
    padding-top: 20px;
    border-top: 1px solid var(--border);
  }

  /* --- Fade in animation --- */
  @keyframes fadeUp {
    from { opacity: 0; transform: translateY(12px); }
    to { opacity: 1; transform: translateY(0); }
  }
  .section { animation: fadeUp 0.4s ease both; }
  .section:nth-child(2) { animation-delay: 0.05s; }
  .section:nth-child(3) { animation-delay: 0.1s; }
  .section:nth-child(4) { animation-delay: 0.15s; }
</style>
</head>
<body>

<div class="container">
  <div class="header">
    <h1><span>//</span> Server Post-Deployment Setup</h1>
    <p>Configure network, hostname, and domain membership. All changes are applied via PowerShell on the local machine.</p>
    <div class="hostname-badge" id="currentHost">Loading...</div>
  </div>

  <!-- SECTION 1: NETWORK -->
  <div class="section">
    <div class="section-head">
      <div class="section-num">1</div>
      <div class="section-title">Network Configuration</div>
    </div>

    <div class="field">
      <label>Network Adapter</label>
      <select id="adapterSelect" onchange="onAdapterChange()"></select>
    </div>

    <label class="toggle-row" for="dhcpToggle">
      <input type="checkbox" id="dhcpToggle" onchange="toggleDHCP()" checked>
      <div class="toggle-track"></div>
      <span class="toggle-label">Use DHCP (obtain IP automatically)</span>
    </label>

    <div id="staticFields">
      <div class="row">
        <div class="field">
          <label>IP Address</label>
          <input type="text" id="ipAddr" placeholder="192.168.1.10" disabled>
        </div>
        <div class="field">
          <label>Subnet Mask</label>
          <input type="text" id="subnetMask" placeholder="255.255.255.0" value="255.255.255.0" disabled>
        </div>
      </div>
      <div class="field">
        <label>Default Gateway</label>
        <input type="text" id="gateway" placeholder="192.168.1.1" disabled>
      </div>
      <div class="row">
        <div class="field">
          <label>Primary DNS</label>
          <input type="text" id="dns1" placeholder="8.8.8.8" disabled>
        </div>
        <div class="field">
          <label>Secondary DNS</label>
          <input type="text" id="dns2" placeholder="8.8.4.4" disabled>
        </div>
      </div>
    </div>

    <div class="btn-row">
      <button class="btn btn-primary" id="btnApplyNet" onclick="applyNetwork()">Apply Network</button>
      <button class="btn btn-ghost" onclick="loadAdapters()">Refresh Adapters</button>
    </div>
  </div>

  <!-- SECTION 2: HOSTNAME -->
  <div class="section">
    <div class="section-head">
      <div class="section-num">2</div>
      <div class="section-title">Computer Name</div>
    </div>

    <div class="row">
      <div class="field">
        <label>Current Hostname</label>
        <input type="text" id="currentHostInput" disabled>
      </div>
      <div class="field">
        <label>New Hostname</label>
        <input type="text" id="newHostname" placeholder="SRV-APP-01" maxlength="15">
        <div class="hint">1-15 chars: letters, numbers, hyphens</div>
      </div>
    </div>

    <div class="btn-row">
      <button class="btn btn-primary" onclick="applyHostname()">Apply Hostname</button>
    </div>
  </div>

  <!-- SECTION 3: DOMAIN JOIN -->
  <div class="section">
    <div class="section-head">
      <div class="section-num">3</div>
      <div class="section-title">Domain Join</div>
    </div>

    <div class="field">
      <label>Domain Name</label>
      <input type="text" id="domainName" placeholder="corp.local">
    </div>
    <div class="field">
      <label>OU Path (optional)</label>
      <input type="text" id="ouPath" placeholder="OU=Servers,DC=corp,DC=local">
      <div class="hint">Leave blank to use the default Computers container</div>
    </div>
    <div class="row">
      <div class="field">
        <label>Admin Username</label>
        <input type="text" id="domUser" placeholder="administrator">
      </div>
      <div class="field">
        <label>Admin Password</label>
        <input type="password" id="domPass" placeholder="********">
      </div>
    </div>

    <div class="btn-row">
      <button class="btn btn-success" onclick="joinDomain()">Join Domain</button>
    </div>
  </div>

  <!-- FOOTER -->
  <div class="footer">
    <button class="btn btn-danger" onclick="restartServer()">Restart Now</button>
    <button class="btn btn-ghost" onclick="window.close()">Close</button>
  </div>
</div>

<!-- Toast container -->
<div class="toast" id="toast"></div>

<script>
  let adapters = [];

  // --- Toast notification ---
  function toast(msg, type) {
    const t = document.getElementById('toast');
    t.textContent = msg;
    t.className = 'toast ' + type + ' show';
    clearTimeout(t._timer);
    t._timer = setTimeout(() => t.classList.remove('show'), 5000);
  }

  // --- API helper ---
  async function api(action, data) {
    const resp = await fetch('/api', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ action, ...data })
    });
    return resp.json();
  }

  // --- Set button loading state ---
  function setLoading(btn, loading) {
    if (loading) {
      btn.dataset.origText = btn.textContent;
      btn.innerHTML = '<span class="spinner"></span> Working...';
      btn.disabled = true;
    } else {
      btn.textContent = btn.dataset.origText || btn.textContent;
      btn.disabled = false;
    }
  }

  // --- Load adapters ---
  async function loadAdapters() {
    try {
      const res = await api('getAdapters');
      adapters = res.adapters || [];
      const sel = document.getElementById('adapterSelect');
      sel.innerHTML = '';
      adapters.forEach((a, i) => {
        const opt = document.createElement('option');
        opt.value = i;
        opt.textContent = a.name + '  [' + a.status + ']  ' + (a.mac || '');
        sel.appendChild(opt);
      });
      if (adapters.length > 0) onAdapterChange();
      toast('Adapters loaded', 'info');
    } catch (e) {
      toast('Failed to load adapters: ' + e.message, 'err');
    }
  }

  // --- On adapter selection change ---
  function onAdapterChange() {
    const idx = document.getElementById('adapterSelect').value;
    const a = adapters[idx];
    if (!a) return;
    document.getElementById('dhcpToggle').checked = a.dhcp;
    document.getElementById('ipAddr').value = a.ip || '';
    document.getElementById('subnetMask').value = a.mask || '255.255.255.0';
    document.getElementById('gateway').value = a.gateway || '';
    document.getElementById('dns1').value = a.dns1 || '';
    document.getElementById('dns2').value = a.dns2 || '';
    toggleDHCP();
  }

  // --- Toggle DHCP ---
  function toggleDHCP() {
    const dhcp = document.getElementById('dhcpToggle').checked;
    const fields = document.querySelectorAll('#staticFields input');
    fields.forEach(f => f.disabled = dhcp);
  }

  // --- Apply network ---
  async function applyNetwork() {
    const btn = document.getElementById('btnApplyNet');
    setLoading(btn, true);
    try {
      const idx = document.getElementById('adapterSelect').value;
      const a = adapters[idx];
      const data = {
        interfaceIndex: a.index,
        adapterName: a.name,
        dhcp: document.getElementById('dhcpToggle').checked,
        ip: document.getElementById('ipAddr').value,
        mask: document.getElementById('subnetMask').value,
        gateway: document.getElementById('gateway').value,
        dns1: document.getElementById('dns1').value,
        dns2: document.getElementById('dns2').value
      };
      const res = await api('applyNetwork', data);
      if (res.success) {
        toast(res.message, 'ok');
      } else {
        toast(res.message, 'err');
      }
    } catch (e) {
      toast('Error: ' + e.message, 'err');
    }
    setLoading(btn, false);
  }

  // --- Apply hostname ---
  async function applyHostname() {
    const name = document.getElementById('newHostname').value.trim();
    if (!name) { toast('Enter a hostname first', 'err'); return; }
    if (!/^[a-zA-Z0-9\-]{1,15}$/.test(name)) {
      toast('Invalid hostname: 1-15 chars, letters/numbers/hyphens only', 'err');
      return;
    }
    try {
      const res = await api('applyHostname', { hostname: name });
      if (res.success) toast(res.message, 'ok');
      else toast(res.message, 'err');
    } catch (e) {
      toast('Error: ' + e.message, 'err');
    }
  }

  // --- Join domain ---
  async function joinDomain() {
    const domain = document.getElementById('domainName').value.trim();
    const user = document.getElementById('domUser').value.trim();
    const pass = document.getElementById('domPass').value;
    const ou = document.getElementById('ouPath').value.trim();

    if (!domain) { toast('Enter a domain name', 'err'); return; }
    if (!user) { toast('Enter admin username', 'err'); return; }
    if (!pass) { toast('Enter admin password', 'err'); return; }

    if (!confirm('Join this computer to domain ' + domain + '?\n\nThis will require a restart.')) return;

    try {
      const res = await api('joinDomain', { domain, user, pass, ou });
      if (res.success) toast(res.message, 'ok');
      else toast(res.message, 'err');
    } catch (e) {
      toast('Error: ' + e.message, 'err');
    }
  }

  // --- Restart ---
  async function restartServer() {
    if (!confirm('Restart the server now?')) return;
    try {
      await api('restart');
      toast('Restarting...', 'info');
    } catch (e) {
      toast('Restart command sent', 'info');
    }
  }

  // --- Init ---
  async function init() {
    try {
      const res = await api('getInfo');
      document.getElementById('currentHost').textContent = res.hostname || '?';
      document.getElementById('currentHostInput').value = res.hostname || '?';
    } catch(e) {}
    loadAdapters();
  }
  init();
</script>
</body>
</html>
'@

# ----------------------------------------------
#  HTTP SERVER
# ----------------------------------------------

$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add($Prefix)

try {
    $listener.Start()
} catch {
    Write-Host "ERROR: Could not start listener on $Prefix" -ForegroundColor Red
    Write-Host $_.Exception.Message
    Write-Host ""
    Write-Host "Try: netsh http add urlacl url=$Prefix user=Everyone"
    exit 1
}

Write-Host ""
Write-Host "  Server Post-Deployment Setup" -ForegroundColor Cyan
Write-Host "  ----------------------------" -ForegroundColor DarkGray
Write-Host "  Listening on $Prefix" -ForegroundColor Green
Write-Host "  Press Ctrl+C to stop." -ForegroundColor DarkGray
Write-Host ""

# Open default browser
Start-Process $Prefix

# Serve requests
try {
    while ($listener.IsListening) {
        $context  = $listener.GetContext()
        $request  = $context.Request
        $response = $context.Response

        $path = $request.Url.AbsolutePath

        if ($path -eq '/' -and $request.HttpMethod -eq 'GET') {
            # Serve HTML page
            $buffer = [System.Text.Encoding]::UTF8.GetBytes($HTML)
            $response.ContentType = 'text/html; charset=utf-8'
            $response.ContentLength64 = $buffer.Length
            $response.OutputStream.Write($buffer, 0, $buffer.Length)
            $response.Close()
            continue
        }

        if ($path -eq '/api' -and $request.HttpMethod -eq 'POST') {
            # Read JSON body
            $reader = New-Object System.IO.StreamReader($request.InputStream, $request.ContentEncoding)
            $body = $reader.ReadToEnd()
            $reader.Close()
            $json = $body | ConvertFrom-Json

            $result = @{ success = $true; message = 'OK' }

            try {
                switch ($json.action) {

                    'getInfo' {
                        $result = @{ hostname = $env:COMPUTERNAME }
                    }

                    'getAdapters' {
                        $adapterList = @()
                        $rawAdapters = Get-NetAdapter | Where-Object { $_.Status -eq 'Up' -or $_.Status -eq 'Disconnected' }
                        foreach ($a in $rawAdapters) {
                            $ip  = Get-NetIPAddress -InterfaceIndex $a.InterfaceIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue | Select-Object -First 1
                            $gw  = Get-NetRoute -InterfaceIndex $a.InterfaceIndex -DestinationPrefix '0.0.0.0/0' -ErrorAction SilentlyContinue | Select-Object -First 1
                            $dns = Get-DnsClientServerAddress -InterfaceIndex $a.InterfaceIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue
                            $pfx = if ($ip) { $ip.PrefixLength } else { 24 }
                            $adapterList += @{
                                name    = $a.Name
                                index   = $a.InterfaceIndex
                                status  = [string]$a.Status
                                mac     = $a.MacAddress
                                ip      = if ($ip) { $ip.IPAddress } else { '' }
                                mask    = (PrefixToMask $pfx)
                                gateway = if ($gw) { $gw.NextHop } else { '' }
                                dns1    = if ($dns -and $dns.ServerAddresses.Count -gt 0) { $dns.ServerAddresses[0] } else { '' }
                                dns2    = if ($dns -and $dns.ServerAddresses.Count -gt 1) { $dns.ServerAddresses[1] } else { '' }
                                dhcp    = if ($ip) { $ip.PrefixOrigin -eq 'Dhcp' } else { $true }
                            }
                        }
                        $result = @{ adapters = $adapterList }
                    }

                    'applyNetwork' {
                        $idx = [int]$json.interfaceIndex
                        Write-Host "  [NET] Flushing adapter index $idx..." -ForegroundColor Yellow

                        # Aggressive flush
                        Flush-AdapterNetwork -InterfaceIndex $idx

                        if ($json.dhcp -eq $true) {
                            Set-NetIPInterface -InterfaceIndex $idx -Dhcp Enabled -ErrorAction Stop
                            Set-DnsClientServerAddress -InterfaceIndex $idx -ResetServerAddresses
                            Start-Sleep -Milliseconds 500
                            $aName = $json.adapterName
                            ipconfig /release $aName 2>$null | Out-Null
                            ipconfig /renew $aName 2>$null | Out-Null
                            Write-Host "  [NET] DHCP enabled on $aName" -ForegroundColor Green
                            $result = @{ success = $true; message = 'DHCP enabled and renewed. Old config cleared.' }
                        } else {
                            $ip   = $json.ip
                            $mask = $json.mask
                            $gw   = $json.gateway
                            $dns1 = $json.dns1
                            $dns2 = $json.dns2
                            $prefix = MaskToPrefix $mask

                            Set-NetIPInterface -InterfaceIndex $idx -Dhcp Disabled -ErrorAction SilentlyContinue
                            Start-Sleep -Milliseconds 300

                            if ($gw) {
                                New-NetIPAddress -InterfaceIndex $idx -AddressFamily IPv4 -IPAddress $ip -PrefixLength $prefix -DefaultGateway $gw -ErrorAction Stop | Out-Null
                            } else {
                                New-NetIPAddress -InterfaceIndex $idx -AddressFamily IPv4 -IPAddress $ip -PrefixLength $prefix -ErrorAction Stop | Out-Null
                            }

                            $dnsServers = @()
                            if ($dns1) { $dnsServers += $dns1 }
                            if ($dns2) { $dnsServers += $dns2 }
                            if ($dnsServers.Count -gt 0) {
                                Set-DnsClientServerAddress -InterfaceIndex $idx -ServerAddresses $dnsServers
                            }

                            Clear-DnsClientCache -ErrorAction SilentlyContinue

                            Write-Host "  [NET] Static $ip/$prefix applied" -ForegroundColor Green
                            $result = @{ success = $true; message = "Static IP $ip/$prefix applied. Old config cleared first." }
                        }
                    }

                    'applyHostname' {
                        $name = $json.hostname
                        Rename-Computer -NewName $name -Force -ErrorAction Stop
                        Write-Host "  [HOST] Renamed to $name" -ForegroundColor Green
                        $result = @{ success = $true; message = "Hostname set to $name. Restart required." }
                    }

                    'joinDomain' {
                        $domain = $json.domain
                        $secPass = ConvertTo-SecureString $json.pass -AsPlainText -Force
                        $cred = New-Object System.Management.Automation.PSCredential(($domain + '\' + $json.user), $secPass)
                        $params = @{
                            DomainName  = $domain
                            Credential  = $cred
                            Force       = $true
                            ErrorAction = 'Stop'
                        }
                        if ($json.ou) { $params['OUPath'] = $json.ou }
                        Add-Computer @params
                        Write-Host "  [DOMAIN] Joined $domain" -ForegroundColor Green
                        $result = @{ success = $true; message = "Joined domain $domain. Restart required." }
                    }

                    'restart' {
                        Write-Host "  [RESTART] Restarting..." -ForegroundColor Red
                        $result = @{ success = $true; message = 'Restarting...' }
                        # Send response first, then restart
                        $jsonOut = $result | ConvertTo-Json -Compress
                        $buffer = [System.Text.Encoding]::UTF8.GetBytes($jsonOut)
                        $response.ContentType = 'application/json'
                        $response.ContentLength64 = $buffer.Length
                        $response.OutputStream.Write($buffer, 0, $buffer.Length)
                        $response.Close()
                        Start-Sleep -Seconds 1
                        Restart-Computer -Force
                        continue
                    }

                    default {
                        $result = @{ success = $false; message = 'Unknown action: ' + $json.action }
                    }
                }
            } catch {
                $result = @{ success = $false; message = $_.Exception.Message }
                Write-Host "  [ERROR] $($_.Exception.Message)" -ForegroundColor Red
            }

            $jsonOut = $result | ConvertTo-Json -Depth 5 -Compress
            $buffer = [System.Text.Encoding]::UTF8.GetBytes($jsonOut)
            $response.ContentType = 'application/json'
            $response.ContentLength64 = $buffer.Length
            $response.OutputStream.Write($buffer, 0, $buffer.Length)
            $response.Close()
            continue
        }

        # 404
        $response.StatusCode = 404
        $response.Close()
    }
} finally {
    $listener.Stop()
    Write-Host "  Listener stopped." -ForegroundColor DarkGray
}
