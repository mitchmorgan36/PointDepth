param(
    [string]$Configuration = "Release",
    [string]$AutoCADInstallDir = "C:\Program Files\Autodesk\AutoCAD 2026\"
)

$ErrorActionPreference = "Stop"

dotnet build .\PointDepth.csproj `
    --configuration $Configuration `
    -p:AutoCADInstallDir="$AutoCADInstallDir"
