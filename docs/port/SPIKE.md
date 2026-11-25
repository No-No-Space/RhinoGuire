# WP-0 spike results

Date: 2026-11-25. Rhino 8.37.26307.15001 on .NET 10.0.15, Windows 11. Started by Antigravity (Gemini), finished by Claude.

All automated checks pass. Three checks need a person at the keyboard; they are listed under [Manual checks for Gate 0](#manual-checks-for-gate-0).

## Results

| Plan item | Check | Result | Evidence |
| --- | --- | --- | --- |
| 1 | Plugin loads in Rhino | Pass | `PlugIn.LoadPlugIn` → `Success`. `RhinoGuire.rhp` and `RhinoGuire.UI.dll` load into the **Default** load context. |
| 1 | Rebuild loop | Pass | Close Rhino → change `SpikeRunner.BuildMarker` → `dotnet build` → new Rhino → `RGSpike` loads the new build by itself (`wp0-build-2`). A build while any Rhino has the plugin loaded fails with MSB3027 (file locked). |
| 2 | A command opens a dockable WPF panel | Pass | `RGSpike` → `Panels.IsPanelVisible` true; the panel is hosted in an `HwndSource`. Registered as a `UserControl` with `[Guid]` and a `uint` constructor. |
| 3 | Typing in a panel textbox reaches the textbox | **Manual** | Needs real keystrokes, so it can't be automated. See below. |
| 4 | Embedded dictionaries load by pack URI | Pass | Tokens, theme and styles resolve; `Application.Current` exists in Rhino (`System.Windows.Application`) and the `pack` scheme is registered, so no workaround is needed. |
| 4 | Embedded fonts resolve | Pass | Gully Bold, Moderustic Regular/Bold, Fliege Mono Light/Regular load from the assembly's resources, real faces, no synthetic bold. |
| 4 | Light/dark swap re-themes an open panel | Pass | Panel background and text follow `{DynamicResource}` on swap and back. |
| 4 | Loose external skin overrides the embedded skin | Pass | A loose `Tokens.Light.xaml` overrides one key while embedded keys still resolve; a font shipped in the skin folder loads by relative path (`./Fonts/#…`); a style in the loose file targets `RhinoGuire.UI.SpikeProbe` and resolves to the loaded type. |
| 5 | `RhinoGet.GetOneObject` from a panel button returns | Pass (preselected) + **Manual** | The button's real `Click` handler returns with the preselected object. Picking in a viewport after the click needs a person. |
| 6 | ClosedXML inside Rhino | Pass | ClosedXML 0.104.2 / DocumentFormat.OpenXml 3.1.1 write and re-read a workbook with non-ASCII text. `System.IO.Packaging` resolves once, to WPF's copy in the .NET 10 Windows Desktop runtime; pack URIs keep working afterwards. |
| 7 | Panel renders to PNG | Pass | Docked panel (light and dark) and an off-screen instance that was never docked both render with `RenderTargetBitmap`. |

Raw output of the last run: `%TEMP%\RhinoGuire\spike\spike-results.json` and the PNGs beside it (not in git).

## Decisions for Gate 0

| Topic | WP-0 choice | Notes |
| --- | --- | --- |
| Target framework | `net7.0-windows` (UI, Plugin), `net7.0` (Core, Engine, tests) | The McNeel template for Rhino 8 + WPF gives `net7.0-windows;net48`; `net48` is dropped as planned. Rhino 8.37 runs .NET 10 and loads it without issue. The build shows no warnings. |
| RhinoCommon floor | `8.21.25188.17001` in `src/Directory.Build.props` | The plan proposed 8.20; WP-0 pinned 8.21, which is also the Rhino 8 version in the local NuGet cache. It sets the oldest Rhino 8 the package installs on. |
| Plugin load time | On demand (Rhino's default), changed from `AtStartup` | See the dev loop below. |
| AAJ fonts | Gully Bold; Moderustic Regular, Bold; Fliege Mono Light, Regular, Medium, Bold; OFL text beside each | The faces listed in AAJ's `tokens.json`. Protest Strike is left out until a panel needs it. Extra weights fetched earlier were removed. |

## Dev loop (read before WP-1)

- **Loading once registers the plugin with Rhino.** `PlugIn.LoadPlugIn` writes `HKCU\Software\McNeel\Rhinoceros\8.0\Plug-Ins\{b1162b4d-b97a-429b-aac6-d4262e92fa12}`. From then on, every Rhino 8 session on this machine, including the user's own, knows the `RG*` commands and loads `src\RhinoGuire.Plugin\bin\Debug\net7.0-windows\RhinoGuire.rhp` when one is typed.
- **A loaded plugin locks the build output.** Close every Rhino that has used an `RG*` command or panel before building. If the build fails with MSB3027, ask the user whether their own Rhino has it loaded; do not kill their process.
- **Rhino MCP slots share the user's Rhino profile.** A panel left open when a slot closes is saved into the user's window layout and opens in their next Rhino session. Close spike and tool panels (`RGSpike` toggles) before `close_slot`.
- **Loop used in WP-0:** `close_slot` → `dotnet build src/RhinoGuire.sln -c Debug` → `spawn_slot 8` → run any `RG*` command (the registered plugin loads) → check. The first time on a machine, load it once with `run_csharp`: `Rhino.PlugIns.PlugIn.LoadPlugIn(@"<repo>\src\RhinoGuire.Plugin\bin\Debug\net7.0-windows\RhinoGuire.rhp", out Guid id);`
- **Without registration:** the trimmed launch profiles (`.vscode/launch.json`, `Properties/launchSettings.json`) start Rhino 8 with `RHINO_PACKAGE_DIRS` pointing at `bin\Debug`, McNeel's usual debugging route. Not used in WP-0.
- **Agents without Rhino MCP tools** can drive Rhino with `tools/rhino_client.py` (stdio client for McNeel's `rhino-mcp-router`, written by Antigravity; it uses the router installed under `%APPDATA%\McNeel\Rhinoceros\ai\bin`). Do not copy router binaries into the repo.

To remove the dev registration later: close Rhino, then delete the registry key above, or turn the plugin off in Rhino's **Options → Plug-ins**.

## Re-running the spike

1. Build, then in a spawned slot: `RGSpike` (opens the panel), then `RGSpikeTest`.
2. `RGSpikeTest` runs on the next idle tick, outside any command, and writes `%TEMP%\RhinoGuire\spike\spike-results.json` after every check (`"complete": true` at the end).
3. `RGSpike` again to close the panel, then `close_slot`.

## Findings for later work packages

- **WP-1A.** WPF's default chrome is unthemed: the scrollbar stays light grey on the dark theme, and the default button template draws its own rounded blue hover (the spike button needed its own template). Every control in the contract needs a template, not only setters.
- **WP-1A / RGUiSnapshot.** `RenderTargetBitmap` renders at the monitor's scale (1.5× on this machine). For comparable PNGs across machines, render snapshots at a fixed scale.
- **WP-1C.** Off-screen rendering works without docking, so `RGUiSnapshot` can render any panel type without opening it in Rhino.
- **WP-1D.** External skins can ship fonts and style `RhinoGuire.UI` types from loose XAML. Loose XAML can't have code-behind (`x:Class`), as planned.
- **WP-4A.** The plugin's `bin` contains a `System.IO.Packaging.dll` from ClosedXML's dependency graph that Rhino never loads (the runtime's copy wins). It can be left out of the Yak package once the release build is checked.
- The spawned Rhino docks new panels small (190 × 173 px); size panels with a minimum width rather than relying on the dock.

## Manual checks for Gate 0

In your own Rhino 8 (the plugin is already registered):

1. Type `RGSpike`. The "RhinoGuire spike" panel opens.
2. **Typing:** click in the box under *1 · TYPING*, type a few letters, press Enter. Pass if the letters appear in the box and nothing appears on Rhino's command line or runs.
3. **Pick:** click *Pick object*, then click any object in a viewport. Pass if its id appears under the button and Rhino stays responsive. Also try pressing Esc instead of picking: the text should read "Pick ended: Cancel".
4. **Theme:** click *Toggle dark / light* twice.
5. Type `RGSpike` again to close the panel.

| Manual check | Result | Notes |
| --- | --- | --- |
| Typing reaches the textbox | | |
| Pick from a panel button | | |
| Esc cancels the pick | | |
