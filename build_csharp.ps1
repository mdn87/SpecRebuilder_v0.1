Write-Host "Building Word Numbering Rebuilder..." -ForegroundColor Green

Set-Location src
dotnet restore
dotnet build --configuration Release
dotnet publish --configuration Release --output ../output

Write-Host "Build complete!" -ForegroundColor Green
Write-Host "Executable location: output/WordNumberingRebuilder.exe" -ForegroundColor Yellow 