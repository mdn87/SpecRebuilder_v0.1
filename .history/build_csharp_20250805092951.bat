@echo off
echo Building Word Numbering Rebuilder...

cd src
dotnet restore
dotnet build --configuration Release
dotnet publish --configuration Release --output ../output

echo Build complete!
echo Executable location: output/WordNumberingRebuilder.exe 