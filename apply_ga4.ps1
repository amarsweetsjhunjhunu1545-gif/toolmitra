# apply_ga4.ps1 – inject Google Analytics 4 tag into every HTML page in the frontend folder

$frontendPath = "C:\Users\risha\OneDrive\Desktop\orbixa final website\orbixa final website\frontend"
# Replace with your real GA4 Measurement ID
$ga4Id = "G-XXXXXXXXXX"

Write-Host "Injecting GA4 tag into all HTML files..."

Get-ChildItem -Path $frontendPath -Filter *.html | ForEach-Object {
    $file = $_.FullName
    $content = Get-Content $file -Raw
    $ga4Tag = @"
    <!-- 👁️ Google Analytics 4 -->
    <script async src="https://www.googletagmanager.com/gtag/js?id=$ga4Id"></script>
    <script>
      window.dataLayer = window.dataLayer || [];
      function gtag(){dataLayer.push(arguments);}
      gtag('js', new Date());
      gtag('config', '$ga4Id');
    </script>
"@
    if ($content -notmatch [regex]::Escape($ga4Id)) {
        $newContent = $content -replace "(</head>)", "`n$ga4Tag`n`$1"
        Set-Content -Path $file -Value $newContent -Encoding UTF8
        Write-Host "Updated $($_.Name)"
    } else {
        Write-Host "Skipped $($_.Name) – GA4 already present"
    }
}
Write-Host "GA4 injection completed."
