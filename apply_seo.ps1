Write-Host "Applying SEO meta tags to HTML files..."
$frontendPath = "C:\Users\risha\OneDrive\Desktop\orbixa final website\orbixa final website\frontend"
Get-ChildItem -Path $frontendPath -Filter *.html | ForEach-Object {
    $file = $_.FullName
    $content = Get-Content $file -Raw
    # Determine page name without extension
    $pageName = $_.BaseName -replace "-", " "
    $title = "$($pageName) Free Online | OrbixaPDFTool"
    $description = "Free $($pageName) tool online. No sign‑up, no watermark, fast and secure."
    # SEO block to insert
    $seoBlock = @"
    <!-- ✅ SEO TITLE - Primary keyword first -->
    <title>$title</title>
    <!-- ✅ SEO META DESCRIPTION -->
    <meta name="description" content="$description">
    <!-- ✅ LONG-TAIL KEYWORDS -->
    <meta name="keywords" content="free $($pageName) online, $($pageName) tool, $($pageName) without watermark, $($pageName) free, pdf tools" />
    <meta name="author" content="OrbixaPDFTool" />
    <meta name="robots" content="index, follow" />
    <meta name="theme-color" content="#6366f1" />
    <link rel="canonical" href="https://orbixapdftool.in/$($_.Name)" />
    <!-- ✅ OPEN GRAPH -->
    <meta property="og:type" content="website" />
    <meta property="og:url" content="https://orbixapdftool.in/$($_.Name)" />
    <meta property="og:title" content="$title" />
    <meta property="og:description" content="$description" />
    <meta property="og:image" content="https://orbixapdftool.in/assets/favicon.png" />
    <meta property="og:site_name" content="OrbixaPDFTool" />
    <!-- ✅ TWITTER CARD -->
    <meta name="twitter:card" content="summary_large_image" />
    <meta name="twitter:title" content="$title" />
    <meta name="twitter:description" content="$description" />
    <meta name="twitter:image" content="https://orbixapdftool.in/assets/favicon.png" />
"@
    # Insert SEO block after the opening <head> tag if not already present
    if ($content -notmatch "<title>") {
        $newContent = $content -replace "(<head>\s*)", "`$1`n$seoBlock"
        Set-Content -Path $file -Value $newContent -Encoding UTF8
        Write-Host "Updated $($_.Name)"
    } else {
        Write-Host "Skipped $($_.Name) (already has title)"
    }
}
Write-Host "SEO meta tags applied."
