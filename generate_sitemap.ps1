# generate_sitemap.ps1 – create a sitemap.xml for all HTML pages in the frontend folder
$frontendPath = "C:\Users\risha\OneDrive\Desktop\orbixa final website\orbixa final website\frontend"
$baseUrl = "https://orbixapdftool.in"

$xmlHeader = '<?xml version="1.0" encoding="UTF-8"?>'
$urlsetOpen = '<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">'

$entries = Get-ChildItem -Path $frontendPath -Filter *.html | ForEach-Object {
    $loc = "$baseUrl/$($_.Name)"
    $lastMod = (Get-Item $_.FullName).LastWriteTime.ToString('yyyy-MM-dd')
    @"  <url>
    <loc>$loc</loc>
    <lastmod>$lastMod</lastmod>
    <changefreq>weekly</changefreq>
    <priority>0.8</priority>
  </url>"@
}

$sitemapContent = $xmlHeader + "`n" + $urlsetOpen + "`n" + ($entries -join "`n") + "`n</urlset>"
$sitemapPath = Join-Path $frontendPath "sitemap.xml"
Set-Content -Path $sitemapPath -Value $sitemapContent -Encoding UTF8
Write-Host "Sitemap generated at $sitemapPath"
