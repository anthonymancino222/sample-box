# Builds test/index.html from index.html: same app, but wired to the TEST tracker storage + the TEST copy of the Google Sheet,
# with its own browser storage keys so it can never mix with the live app.  Re-run after every change to index.html.
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$u8 = New-Object Text.UTF8Encoding $false
$t = [IO.File]::ReadAllText("$root\index.html", [Text.Encoding]::UTF8)

$TEST_BACKEND = 'https://script.google.com/macros/s/AKfycbwXhblrMm9LcpQITxHextfT4kp9q_rWzsdCfhi6nKOCqpTIOt1tuPSHR4h8TNgP93CU/exec'
$TEST_READER  = 'https://script.google.com/macros/s/AKfycbzIrfIWqrBjre1Ebq4pkD04BtK6sUbcz_ZzRcTT9mD6YxHytE1UdNaibXFg7nWKPPxI/exec'
$TEST_WRITER  = 'https://script.google.com/macros/s/AKfycbwinuDckO0GUAjaOquTfN0JWI5vfP5JCT0eBp7VykxuT_0uzk-GBoJJmNdszOttjD0J/exec'

function Rep($rx, $new) { $m = [regex]::Match($script:t, $rx); if (-not $m.Success) { throw "not found: $rx" }; $script:t = $script:t.Substring(0, $m.Index) + $new + $script:t.Substring($m.Index + $m.Length) }

Rep "const BACKEND_URL = '[^']*';" "const BACKEND_URL = '$TEST_BACKEND';"
Rep "const SOURCE_URL = '[^']*';"  "const SOURCE_URL = '$TEST_READER';"
Rep "const WRITER_URL = '[^']*';" "const WRITER_URL = '$TEST_WRITER';"
$t = $t.Replace('sampleBoxTracker', 'sampleBoxTrackerTEST')                       # separate browser storage
$t = $t.Replace('href="icons/', 'href="../icons/').Replace('url(icons/marble.png)', 'url(../icons/marble.png)')
$t = [regex]::Replace($t, '<link rel="manifest" href="manifest.json">\r?\n?', '')   # the test page is not an installable app
$t = $t.Replace('<title>Sample Box Tracker</title>', '<title>Sample Box Tracker (TEST)</title>')
$banner = '<div class="testbar">TEST VERSION &middot; connected to the TEST copy of the Google Sheet &middot; nothing here touches the live app or the real sheet</div>'
Rep '<body>' ("<body>`n" + $banner)
$css = '.testbar{background:#a8442e;color:#fff8ef;text-align:center;padding:7px 12px;font:600 12.5px var(--sans);letter-spacing:.05em}'
$i = $t.LastIndexOf('</style>'); $t = $t.Substring(0, $i) + $css + "`n" + $t.Substring($i)

New-Item -ItemType Directory -Force "$root\test" | Out-Null
[IO.File]::WriteAllText("$root\test\index.html", $t, $u8)
"built test/index.html ($($t.Length) chars)"
