# PointDepth

PointDepth is a Civil 3D .NET command that writes each selected COGO point group's vertical difference to a selected Civil 3D surface into a point UDP named `Depth_To_Surface`.

The sign convention is:

- Positive: point elevation is above the selected surface.
- Negative: point elevation is below the selected surface.
- Zero: point elevation matches the selected surface.

## Use

1. Build `PointDepth.dll`.
2. In Civil 3D 2026, run `NETLOAD` and load the built DLL.
3. Run `AddPointDepth`.
4. Select an existing point group from the numbered command-line prompt.
5. Select the existing surface to compare against by typing its number or picking it in the drawing.

PointDepth creates a numeric `Depth_To_Surface` UDP when needed. Existing numeric `Depth_To_Surface` UDPs are reused. After writing depths, PointDepth creates or updates two sign point groups:

- `PointDepth_Positive`: points where `Depth_To_Surface > 0`
- `PointDepth_Negative`: points where `Depth_To_Surface < 0`

PointDepth first attempts to define those groups with Civil 3D custom queries against the `Depth_To_Surface` UDP. If Civil 3D rejects the UDP query through the .NET API, PointDepth rolls that group update back and populates the groups with point-number include queries based on the depths written in the current run.

Points outside the selected surface are skipped and reported at the command line.

## Build

This project targets Civil 3D 2026 / AutoCAD 2026 on .NET 8.

From this repo:

```powershell
.\build.ps1
```

If Windows blocks the script because of the local execution policy, run:

```powershell
powershell -ExecutionPolicy Bypass -File .\build.ps1
```

Or directly:

```powershell
dotnet build .\PointDepth.csproj --configuration Release
```

If the machine has the .NET 8 runtime but not the .NET 8 reference packs, install the .NET 8 SDK or build once with a NuGet source available:

```powershell
dotnet build .\PointDepth.csproj --configuration Release --source https://api.nuget.org/v3/index.json
```

If AutoCAD is installed somewhere other than `C:\Program Files\Autodesk\AutoCAD 2026\`, pass the install folder:

```powershell
.\build.ps1 -AutoCADInstallDir "D:\Autodesk\AutoCAD 2026\"
```

The release DLL is written to `bin\Release\PointDepth.dll`.

Civil 3D keeps a loaded .NET DLL locked for the rest of that session. Close Civil 3D before rebuilding over `bin\Release\PointDepth.dll`, or build to a separate output folder while testing:

```powershell
dotnet build .\PointDepth.csproj --configuration Release --no-restore -p:OutputPath=bin\Release-test\
```
