# Run both GA4 injection and sitemap generation
powershell -NoProfile -ExecutionPolicy Bypass -File "C:\Users\risha\OneDrive\Desktop\orbixa final website\orbixa final website\apply_ga4.ps1"
if ($LASTEXITCODE -ne 0) { Write-Error "apply_ga4 failed"; exit $LASTEXITCODE }

powershell -NoProfile -ExecutionPolicy Bypass -File "C:\Users\risha\OneDrive\Desktop\orbixa final website\orbixa final website\generate_sitemap.ps1"
if ($LASTEXITCODE -ne 0) { Write-Error "generate_sitemap failed"; exit $LASTEXITCODE }

Write-Host "All SEO tasks completed successfully."
