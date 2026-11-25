# RhinoGuire → C# Rhino plugin with skinnable WPF UI (AAJ public, White internal)

Plan for executing agents. Each work package (WP) is sized for one agent in one session and names its inputs, owned files and acceptance checks.

## Context

RhinoGuire is ten Rhino 8 tools written as CPython scripts (~11,900 lines) with Eto.Forms UIs and a neo-brutalist theme (`ui/theme.py`). Aksel wants:

1. The UI moved from Eto to **WPF** with a better user experience. Mac support is knowingly dropped.
2. The **White Arkitekter Design System** applied for use at White, and the **AAJ Design System** (Aksel's own studio identity) as the public face of the repo.
3. Unfinished tool development completed.
4. Project documentation brought up to date.

Decisions taken with Aksel on 2026-11-25:

| Decision | Choice |
| --- | --- |
| UI architecture | **Full C# Rhino plugin** (`.rhp`): real Rhino commands, dockable WPF panels. This is a port of the Python code, not only a UI swap. |
| Public look | **AAJ skin**, compiled into the plugin as the default. Everything public (repo, Food4Rhino package, screenshots, icon, docs) uses AAJ. |
| White look | **White skin in a gitignored folder**, loaded at runtime from a skins folder next to the plugin. No White material (tokens, fonts, logos, PDFs, hex values) is committed or packaged. This supersedes the earlier "commit White tokens" answer. |
| AAJ fonts | Agents fetch desktop TTF/OTF files from the upstream OFL sources (`_Skills/AAJ Design System/fonts/` holds `.woff2` only, which WPF cannot load) and stop to ask if a family cannot be found. |
| Finish-development scope | All four tracks: TerrainTools reliability, Trocha deferred items, Lindero to-do list, release and packaging. |

Machine facts (checked): Rhino **8.37** only, running on **.NET 10.0.15** (measured in WP-0; this plan first assumed .NET 8), .NET SDK **9.0.100**, Visual Studio 2022 Community, `Yak.exe` present, RhinoCommon and CommunityToolkit.Mvvm in the NuGet cache. McNeel `dotnet new` templates installed during WP-0. `_Skills/` is gitignored. Agents spawn their own Rhino MCP slots. WP-0 findings, including the dev loop, are in `docs/port/SPIKE.md`.

## Ground rules for every agent

- **Python stays runnable on `main` until release.** Aksel uses these tools daily on live project models. All C# work lands on the integration branch `feat/csharp-plugin` under `src/`; each WP works on its own branch or worktree and merges into the integration branch. Do not edit, move or delete the Python tools before WP-4C.
- **Python is the behavioural spec.** Port logic 1:1 first and prove parity. Fix bugs and add features only in Phase 3, each with a regression test.
- **Nothing from White reaches git.** `_Skills/` and `skins/White/` are gitignored. Before every commit, check that no staged path or content comes from the White system. The public code speaks only of "external skins".
- **Tool XAML is skin-agnostic.** Tools reference only the keys in `docs/port/SKIN_CONTRACT.md` through `{DynamicResource}`: semantic brushes, metric tokens and named styles. No literal colors, font names, font sizes, border widths or corner radii in tool XAML.
- **Rhino sessions.** Use only Rhino MCP slots you spawned (`spawn_slot` version `8`). Never run commands in an adopted (user-started) slot. Never read or copy from the project paths in `_prefs.json` (`P:\…`). Test only on fixtures in `tests/fixtures/`.
- **Owned files.** A tool WP edits only its own `Tools/<Tool>/` folders. Changes to `RhinoGuire.UI`, the skin contract, shared services or `Directory.Build.props` after Phase 1 go through a note in `docs/port/STATUS.md` and a separate small commit.
- **Status tracking.** Update `docs/port/STATUS.md` (WP table: owner, state, branch, open questions) at the start and end of each session.
- Things not yet verified on this machine are marked **(verify)**. Confirm them in Phase 0 and record the outcome in `docs/port/SPIKE.md` before relying on them.

## Target architecture

```text
src/
  RhinoGuire.sln
  Directory.Build.props        TFM, nullable, version, pinned RhinoCommon
  RhinoGuire.Core/             pure C#, no Rhino, no WPF  -> unit-testable with dotnet test
  RhinoGuire.Engine/           RhinoCommon geometry + document logic, no WPF
  RhinoGuire.UI/               WPF library: skin loader, custom controls, charts, AAJ skin (embedded)
    Skins/AAJ/                 tokens.json (vendored), skin-map.json, Tokens.{Light,Dark,Shared}.xaml, Controls.xaml, Fonts/
  RhinoGuire.Plugin/           -> RhinoGuire.rhp: PlugIn, commands, panels, view models
  RhinoGuire.Core.Tests/       xUnit
skins/White/                   GITIGNORED. Tokens.*.xaml, Controls.xaml, Fonts/, skin.json
tests/fixtures/*.3dm           synthetic models built by agents
tests/golden/*.json            results dumped from the Python tools
tools/parity/*.py              golden-dump scripts run inside Rhino
tools/build_aaj_skin.py        tokens.json + skin-map.json -> AAJ Tokens.*.xaml
tools/sync_white_skin.ps1      copies _Skills/White…/wpf/*.xaml into skins/White/ and deploys the folder
docs/port/                     PLAN.md, STATUS.md, SPIKE.md, RECIPE.md, SKIN_CONTRACT.md
```

The `Core` / `Engine` / UI split continues the existing `_core` boundary (`TerrainTools/_core`, `RoadTool/_core`, `AreaMeasurer/_paths.py`): engines never touch UI.

- **Target framework:** `net7.0-windows` with `UseWPF` for UI and Plugin, `net7.0` for Core and Engine. The RhinoCommon 8 NuGet targets `net48` + `net7.0`; McNeel's guide recommends `net8.0` from Rhino 8.20. Scaffold from the current McNeel template (`dotnet new install Rhino.Templates`) and take its answer **(verify)**. No `net48` target. Pin RhinoCommon to a floor version (propose 8.20) so the Yak tag does not float to 8.37.
- **One tool = one command + one dockable panel.** Commands `RGLindero`, `RGArriero`, `RGChivito`, `RGBaquiano`, `RGPregonero`, `RGSebucan`, `RGPadGrader`, `RGWayGrader`, `RGCutFillReport`, `RGTrocha` open the tool's panel; `RGOptions` sets skin and light/dark. `Panels.RegisterPanel` on Windows accepts a `FrameworkElement` subclass with a `[Guid]` and a `uint documentSerialNumber` constructor **(verify, see McNeel `SampleCsWpf`)**.
- **MVVM** with CommunityToolkit.Mvvm. View models live in the Plugin project, hold no WPF types, and call Engine services.
- **Excel:** ClosedXML (MIT) replaces openpyxl **(verify it loads inside Rhino without dependency conflicts; fallback is the OpenXML SDK directly)**.
- **Charts:** custom-drawn WPF elements in `RhinoGuire.UI` (`OnRender`), one paint routine serving both screen and PNG export (`RenderTargetBitmap`). No chart library.
- **Settings:** `PlugIn.Settings` replaces `_prefs.json` folder memory (same keys: `lindero_export_xlsx`, `arriero_export`, `chivito_step2` …). `sc.sticky['terraintools_last_grade']` becomes an in-memory `LastGradeService` on the plugin.

### Skin system

A skin is a set of resource dictionaries that all define the same keys:

| File | Holds |
| --- | --- |
| `Tokens.Light.xaml`, `Tokens.Dark.xaml` | Semantic brushes and colors: `SurfacePage`, `SurfaceSubtle`, `TextPrimary`, `TextSecondary`, `BorderControl`, `FocusRing`, `UiDefault/Hover/Pressed/On`, `Activity*`, `Selection*`, `Status{Success,Warning,Error,Info}{Subtle,Strong}`, `Disabled*`, `Input*`, `Table*`, `ChartSeries1-5`, `ChartGrid/Axis/Label/Text`, `ChartDivergingLow/Mid/High` |
| `Tokens.Shared.xaml` | Metrics: `FontBody`, `FontHeading`, `FontMono`, font sizes and weights, `BorderWidth`, `FocusRingWidth`, `FocusRingOffset`, `ControlHeight`, `ControlHeightSmall`, `IconStrokeWidth`, `Radius` |
| `Controls.xaml` | Named styles and templates: `Tool.Button.Primary/Secondary/Text`, `Tool.Field`, `Tool.Select`, `Tool.CheckBox`, `Tool.Radio`, `Tool.Switch`, `Tool.Tabs`, `Tool.Table.*`, `Tool.SectionHeader`, `Tool.StatusLine` … |

The key list starts from White's semantic token names (they already exist in its generated XAML) and is frozen in `docs/port/SKIN_CONTRACT.md` by WP-1A. Custom controls (C# classes) are shared; their look comes from the skin.

The two systems differ in structure, so each skin has its own `Controls.xaml`:

| | AAJ (`_Skills/AAJ Design System/components/bundle.css`) | White (`_Skills/White…/components/bundle.css`) |
| --- | --- | --- |
| Ground | Hueso `bg-page`, fields and cards a step lighter on `bg-surface` | White `surface-page`, panels on `surface-subtle` |
| Borders | 2px Carbón on every control | 1px rule under inputs, 1px secondary-button outline |
| Primary button | Carbón fill, Gully 700, 40px, optional 4px hard shadow that sinks on press | Dark Denim `ui-default` fill, 44px (32px small), no shadow |
| Field | Boxed, mono uppercase 11px label | Underline rule, heavy 12px label, status states with icon and message |
| Tabs | Fused to the panel, active tab inverted | Text tabs, 5px line under the active one |
| Table header | Carbón, Fliege Mono uppercase; figures Fliege Mono 300 | Denim header; total, subtotal, alternate, changed-cell rows defined |
| Activity / selection | Mar Caribe for both; Buganvilia at most once per screen | Mint for activity, Pink for selection, never swapped |
| Focus | 2px Mar Caribe ring, 2px offset | 4px Dark Denim halo, 2px offset |
| Type | Gully (headings), Moderustic 14.5 (body), Fliege Mono (figures) | Mulish 14 / 12, two weights |

Shared by both: square corners, no gradients, no blur, 120 ms transitions, status always as icon plus word, eyebrow or mono labels typed in capitals (WPF has no text-transform or letter-spacing), numbers right-aligned with tabular figures. Icons are Lucide (ISC) geometry in `RhinoGuire.UI`, stroke width from `IconStrokeWidth`.

- **AAJ skin (public):** compiled into `RhinoGuire.UI` as the default. `tools/build_aaj_skin.py` reads the vendored `tokens.json` and `skin-map.json` (semantic key → AAJ token, light and dark) and writes the token dictionaries. AAJ has no tool-level token layer, so the map must propose values for what AAJ does not define: status subtle/strong pairs (its `color-success` fails contrast with Hueso text), table total, subtotal and alternate rows, changed-cell fill, input hover and focus fills. Mark each as `proposed` in the map for Aksel to approve at Gate 1A. The current `ui/theme.py` palette is already AAJ, so this skin continues today's look.
- **White skin (internal):** loose XAML in `skins/White/`, loaded with `XamlReader` from `<plugin folder>\Skins\<Name>\` or `%APPDATA%\RhinoGuire\Skins\<Name>\` **(verify loose-XAML skins load inside Rhino, including fonts by relative path and references to `RhinoGuire.UI` controls)**. Loose XAML cannot carry code-behind, so every behaviour lives in the shared controls. The token files are the generated `WhiteTokens.*.xaml`, never hand-edited. Mulish (OFL) ships in the skin folder.
- **Selection:** setting `Skin` (default AAJ) and `Theme` (Light, Dark, follow Rhino). When an external skin folder is present it appears in `RGOptions`; a `default.txt` in the Skins folder lets an internal deployment preselect it.

### Compatibility contracts (must survive the port unchanged)

| Tool | Contract | Python reference |
| --- | --- | --- |
| Lindero | Layer rules (parent → level → category, `_` prefix cascade); config JSON v2 and the v1 loader; config stored in the document; Excel sheet layouts; hole-aware union path | `AreaMeasurer/_paths.py`, `Lindero.py` `combined_area` (l.638), `_gather_config`/`_apply_config` (l.2807–2900), `_xl_s1.._xl_s4` (l.3622–3937) |
| Arriero | Excel layout keyed on object GUID; timestamped backup before import; empty-cell modes | `Arriero.py` `export_data_to_excel` (l.233), `import_data_from_excel` (l.328) |
| Chivito | Template and color-map workbook formats | `Chivito.py` l.79–288 |
| Pregonero | `{Key}` → `%<UserText("guid","Key")>%`; missing keys initialised to `TBD` | `Pregonero.py` l.75–300 |
| Trocha | `RG_ROAD*` user strings; layers `RoadTool::Roads` and `::Merged`; one undo record per action | `RoadTool/_core/state.py`, `config.py` |
| TerrainTools | Output layer `TerrainTools::Graded`; D5 daylight rule; grid-prism volumes; originals never modified | `TerrainTools/_core/*`, `DECISIONS.md` |

Models tagged by the Python tools must keep working with the plugin.

## Phase 0 — Bootstrap and spikes (one agent, sequential)

**WP-0** owns the repo skeleton. Nothing else starts until its gate passes.

1. Branch `feat/csharp-plugin`. Copy this plan to `docs/port/PLAN.md`; create `STATUS.md`.
2. `.gitignore`: add `_Skills/`, `skins/White/`, `bin/`, `obj/`, `*.rhp`, `*.yak`. The existing Python-template ignore file has `build/` and `lib/` patterns: do not name C# folders that.
3. Scaffold the solution above from the McNeel template; record the TFM decision.
4. Fonts: fetch TTF/OTF for Gully (github.com/magictype/gully), Moderustic (github.com/Tural/moderustic) and Fliege Mono (Pavel Laptev, locate the upstream release) with their OFL files into `src/RhinoGuire.UI/Skins/AAJ/Fonts/`. Protest Strike is a 36px+ display face: fetch it only if a panel uses it. If a family is not available in desktop format, stop and ask Aksel.
5. Spike in a spawned Rhino slot. Each item gets a pass/fail line in `SPIKE.md`:
   - Plugin loads (expected route: MCP `run_csharp` calling `Rhino.PlugIns.PlugIn.LoadPlugIn(path, out id)`), and the rebuild loop works: `close_slot` → `dotnet build` → `spawn_slot` → load. Rhino locks the `.rhp` while running.
   - A command opens a WPF panel registered as a `FrameworkElement`.
   - Typing in a panel `TextBox` reaches the textbox, not Rhino's command line.
   - Embedded dictionaries load by pack URI from `RhinoGuire.UI` in a host without `App.xaml`; an embedded font resolves; a **loose** dictionary loads from a folder and overrides the embedded skin; swapping skin and Light/Dark re-themes an open panel.
   - `RhinoGet.GetOneObject` started from a panel button click returns control correctly.
   - ClosedXML writes and re-reads a workbook inside Rhino.
   - A panel renders itself to PNG with `RenderTargetBitmap` (this becomes the agents' way to see their UI).

**Gate 0 (Aksel):** read `SPIKE.md`; confirm TFM and any fallback taken.

## Phase 1 — Foundations (three agents in parallel, then the White skin)

**WP-1A — Skin contract, shared controls, AAJ skin** (`src/RhinoGuire.UI/`, `tools/build_aaj_skin.py`)
- Inputs: `_Skills/AAJ Design System/README.md`, `tokens.json`, `components/bundle.css` (exact metrics and states), `components/*/README.md`; White's `README.md` sections *App UI tokens* and *Tool controls* for the semantic key names only.
- Write `SKIN_CONTRACT.md` (every key, its role, which controls consume it), then `SkinManager` (embedded default, external loose skins, Light/Dark, follow Rhino dark mode).
- Shared custom controls: `TreeDataGrid` (flattened rows with level and row kind: header, total, subtotal, data, alternate, hover, selected, changed-cell), `KeySearchComboBox` (port of `theme.bind_key_search`), `SlopeInput` (port of `TerrainTools/_widgets.py`), `StatusLine`, `BusyBar`, `BulletChart`, `ChartSurface` base with PNG export, Lucide icon set.
- AAJ `Controls.xaml` for: button (primary, secondary, text; default and small; icon), field with states, select, checkbox with mixed state, radio group, switch, tabs, scroll bars, tooltips, expander, list box, table styles.
- Gallery panel (`RGGallery`, Debug builds) showing every control in every state, plus `RGUiSnapshot <folder>` rendering any registered panel to PNG per skin and theme.
- Design points to put in front of Aksel at the gate: the `proposed` entries in `skin-map.json`; control density in a docked panel roughly 320–400px wide (AAJ 40px controls against a compact variant); whether the primary button carries the 4px hard shadow in panels.
- Acceptance: gallery PNGs in AAJ Light and Dark checked by the agent against `bundle.css`; a test that fails on any literal color, font or `StaticResource` token reference in tool or control XAML; a test that every contract key exists in every skin dictionary present.

**Gate 1A (Aksel):** approve the AAJ gallery, the proposed tokens and the density choice before any tool UI is built. Approved proposals should be fed back into the AAJ design system's own `tokens.json`.

**WP-1B — Core logic and tests** (`src/RhinoGuire.Core/`, `src/RhinoGuire.Core.Tests/`)
- Port: `TerrainTools/_core/slope.py`, `volumes.py` (grid-prism part), `AreaMeasurer/_paths.py`, `RoadTool/_core/config.py`, Baquiano `SearchCondition.matches_value` (8 match types), Pregonero template parsing (`parse_template_keys`, `_PLACEHOLDER_RE`), Lindero config model with v1→v2 mapping, number formatting helpers.
- Port every case from the three headless suites (`TerrainTools/_core/tests/test_headless.py`, `AreaMeasurer/tests/test_paths.py`, `RoadTool/_core/tests/test_headless.py`) to xUnit.
- Acceptance: `dotnet test` green; the original Python suites still pass unchanged.

**WP-1C — Plugin shell, services, parity harness** (`src/RhinoGuire.Plugin/` shell, `src/RhinoGuire.Engine/Services/`, `tools/parity/`, `tests/`)
- `RhinoGuirePlugIn`, a `ToolCommand<TPanel>` base, panel registration with icons, settings service, `LastGradeService`, `RGOptions`.
- Document services: user-text key scan (port the fast RhinoCommon scan in `Lindero.get_all_user_text_keys`, l.160), layer tree, `EnsureLayer` (from `_widgets.ensure_layer`), unit label, object picking wrappers, undo-record scope, Excel service over ClosedXML, viewport capture.
- Parity harness: build synthetic fixtures through the Rhino MCP and save them to `tests/fixtures/` (stacked boxes on a `Parent::Level::Category` layer tree with known overlaps and a `_TEXT` layer; a rolling terrain as mesh and as trimmed surface; pad curves; centerlines including a tight corner and a closed loop; objects with user text). `tools/parity/dump_<tool>.py` scripts import the Python calc functions and write `tests/golden/<tool>.json`; run them with MCP `run_command` `_-RunPythonScript "<path>"` (MCP `run_python` does not give a trustworthy `scriptcontext.doc`). `RGSelfTest <tool>` (Debug) runs the C# engine on the open fixture and reports each value against golden with its tolerance.
- Acceptance: empty-tool round trip works end to end (command → panel → pick object → status line); goldens exist for every tool.

**WP-1D — White skin** (`skins/White/`, `tools/sync_white_skin.ps1`; starts after Gate 1A, runs alongside Phase 2)
- Inputs: `_Skills/White Arkitekter Design System/README.md` (*App UI tokens*, *Tool controls*), `components/bundle.css`, `components/Tool*/README.md`, `wpf/WhiteTokens.*.xaml`.
- Write White's `Controls.xaml` against the frozen contract; add a small shared-metrics dictionary for contract keys the generated files lack. White's rules to enforce: semantic tokens only, never the `C*` brand keys; no shadows (strip WPF's default popup and tooltip shadows); Mint only for activity or changed values, Pink only for selection; one primary button per panel; 44px hit area, 32px small.
- Acceptance: the contract-coverage test passes for the White skin; gallery PNGs in White Light and Dark checked against White's `bundle.css`; every finished tool panel snapshots cleanly in White without any tool XAML change.
- The folder is not under version control. Tell Aksel when it is done so he can back it up somewhere private.

**Gate 1D (Aksel):** approve the White gallery.

## Phase 2 — Tool ports (parallel after Gate 1A; one agent per WP)

Every tool WP delivers: engine classes in `Engine/Tools/<Tool>/`, view model + panel + command in `Plugin/Tools/<Tool>/`, `RGSelfTest` parity green, panel PNGs in AAJ Light and Dark (and White once WP-1D exists), a manual smoke checklist in the tool's README, and updated `STATUS.md`. Follow `docs/port/RECIPE.md`.

| WP | Tool | Python source (lines) | Depends on | Notes |
| --- | --- | --- | --- | --- |
| 2A | **Baquiano** (pilot) | `SearchData/Baquiano.py` (525) | 1A, 1B, 1C | Dynamic include/exclude rows, key search. The pilot agent writes `RECIPE.md` (folder layout, view-model pattern, picking, busy/status, snapshot, self-test) for the others. |
| 2B | **Terrain engine + Sebucan** | `TerrainTools/_core/terrain.py` (219), `Sebucan.py` (603) | 1C | `TerrainModel.ProjectZ` and mesh coercion are shared by 2F and 2G: land the engine first. |
| 2C | **Pregonero** | `Tagging/Pregonero.py` (607) | 2A | Live text fields, dim-style list, interactive placement loop. |
| 2D | **Arriero** | `DataExporterImporter/Arriero.py` (597) | 2A | Becomes a modeless panel (today a modal `Dialog[bool]`); replace the nine message boxes with inline status where the action is not destructive. |
| 2E | **Chivito** | `DataVisualization/Chivito.py` (1244) | 2A | Three-step flow, color manager, template builder, legend PNG (`on_export_legend_clicked`, l.613), color scheme generator. Object colors are model data, not skin tokens. |
| 2F | **PadGrader, WayGrader, CutFillReport** | `TerrainTools/_core/{grading,meshbuild,volumes,report}.py`, three tool files (1245) | 2B | `paint_report` (CutFillReport l.68) becomes a `ChartSurface`. Cut and fill colors come from `ChartDivergingLow/High` and are used for mesh tint, legend and chart alike. |
| 2G | **Trocha** | `RoadTool/_core/*`, `Trocha.py` (673) | 2B | Keep the both-direction offset test and the half-thickness retry (`road_tool_plan.md` §4a, §8). |
| 2H-1 | **Lindero engine + Excel** | `Lindero.py` l.160–1183 and l.3568–3960 | 1B, 1C | Footprints, Z banding, hole-aware union, `calc_s1..s4`, `calc_r`, Excel writers. Highest parity risk: compare every scenario total and warning against golden. |
| 2H-2 | **Lindero UI** | `Lindero.py` l.1313–3566 | 1A, 2H-1 | Seven tabs, preview tree, target table, bullet charts, write-area panel, config save/load, copy-window-to-clipboard (l.2672, keep the DWM physical-pixel approach). |

**Gate 2 (Aksel):** run each tool's smoke checklist on a real model; this also closes the pending Lindero v0.8 in-Rhino verification (`AreaMeasurer/PLAN.md` §8).

## Phase 3 — Finish development (after a tool passes Gate 2)

**WP-3A — TerrainTools reliability.** `road_tool_plan.md` §12 says these tools "don't work 100% in reality", but no failure is written down. Prerequisite from Aksel: a short list of what goes wrong and, where possible, a stripped model that shows it. The agent reproduces each on a fixture, adds a failing self-test, fixes it in `Engine` (shared `TerrainModel` fixes benefit Sebucan and Trocha too), then delivers the per-way list UI (DECISIONS D9.3). Longitudinal design grade stays backlog unless Aksel promotes it. Final tool names (DECISIONS D8) are Aksel's call at Gate 3.

**WP-3B — Trocha.** Corner-radius fillet exposed in the panel (`corner_radius` exists in config), with a self-intersection check on the offset rails; cross-section profiles (crown or camber, curbs); terrain-change detection that flags roads whose terrain moved; store width and thickness on the slab as well so re-linking an orphaned slab recovers its parameters; decide whether `RG_ROAD_MERGED` output is excluded from Update All (open item §12).

**WP-3C — Lindero.** Excel export for R1/R2; viewport highlight of out-of-tolerance objects; silhouette footprint for cantilevers (README Case B2) behind an option, validated against the bottom-face method on the existing fixtures; optional "measured but excluded from totals" category class (`PLAN.md` §10) only if Aksel confirms.

**Gate 3 (Aksel):** accept each feature; decide TerrainTools names and the plugin version number (currently 0.1.0 in `manifest.yml`).

## Phase 4 — Release and documentation

**WP-4A — Packaging.** `icon.png` (64×64) in the AAJ palette; toolbar `.rui` whose macros are plain commands (`_RGLindero`), which removes the absolute-path problem in `FOOD4RHINO.md` §5; `manifest.yml` for the plugin; `yak build --platform win`; install the `.yak` into a clean spawned Rhino and run every command; listing screenshots from `RGUiSnapshot` in the AAJ skin plus viewport captures. Release check: unzip the `.yak` and confirm it holds no `Skins/White`, no Mulish or AkzWhite, and that `git ls-files` shows nothing under `_Skills/` or `skins/White/`.

**WP-4B — Documentation.** Rewrite for the plugin: root `README.md` (install, commands, panels, Windows-only, skins), `CLAUDE.md` (build and test commands, project layout, WPF and skin-contract rules replacing the Eto rules), the ten tool READMEs, `ui/README.md`, `FOOD4RHINO.md`, `TerrainTools/PLAN.md` and `DECISIONS.md` (add the C# port decision), `AreaMeasurer/PLAN.md` status, `RoadTool/road_tool_plan.md`; add `docs/ARCHITECTURE.md`, `docs/SKINS.md` (how to write an external skin), `docs/MIGRATION.md` (launcher macros → commands) and `CHANGELOG.md`; refresh the Claude auto-memory `MEMORY.md`. Public docs follow the AAJ voice: sentence case, no emoji (the current root README uses emoji headings), no buzzwords, exact names for materials and methods; English for international clarity, tool names stay Spanish. Stale items already found: `CLAUDE.md` says eight launcher shims (there are ten); root README tree lists `RoadTool/tests/` and `TerrainTools/tests/` (they are under `_core/tests/`); `RoadTool/README.md` and `road_tool_plan.md` §11 still say `RG_Trocha`; `FOOD4RHINO.md` status table predates `.yakignore` and the `authors` fix; `Lindero.py` header says v0.5 while `PLAN.md` says v0.8; `TerrainTools/README.md` still lists "Planned tools"; `MEMORY.md` describes Arriero as WinForms/IronPython.

**WP-4C — Retire Python (Gate 4, Aksel decides).** Either move the Python tools, launchers, `ui/theme.py` and `install.py` to `legacy/python/` (excluded from Yak) for one release, or delete them. Then merge `feat/csharp-plugin` to `main`.

## Verification

- **Every commit:** `dotnet build src/RhinoGuire.sln -c Debug` with no warnings from our code; `dotnet test src/RhinoGuire.Core.Tests`; the three Python headless suites still pass; nothing from White staged.
- **Engine parity:** in a spawned slot, open the fixture, `RGSelfTest <tool>`; every value within tolerance of `tests/golden/<tool>.json` (areas and volumes 1e-6 relative; mesh vertex Z at model tolerance).
- **UI:** `RGUiSnapshot` → the agent reads the PNGs for each skin and theme and checks them against that skin's gallery and `bundle.css`; the XAML lint and contract-coverage tests pass.
- **Interactive flows** that need picking in the viewport cannot be driven by agents end to end: each tool README carries a manual checklist that Aksel runs at Gates 2 and 3. The existing ones (`RoadTool/README.md` smoke test, `TerrainTools/PLAN.md` §10, `AreaMeasurer/PLAN.md` §8) are the starting point.
- **Release:** `.yak` installs into a clean Rhino 8, all ten commands open their panels in the AAJ skin, the toolbar loads, dropping `skins/White/` into the Skins folder switches the look, and a model tagged by the Python tools (Trocha roads, Pregonero tags, Lindero doc config) still updates correctly.

## Risks

- **Parity drift in Lindero's union logic.** Mitigated by golden files per scenario and by porting before changing anything.
- **Scope.** This is a rewrite of the whole suite plus two skins. Phases are ordered so that the integration branch always holds a loadable plugin with a growing set of finished tools, and `main` keeps the working Python tools throughout.
- **The White skin has no version history.** It lives only in a gitignored folder on this machine until Aksel backs it up.
- **Skin drift.** A tool that works in one skin may clip or misalign in the other because control sizes differ (40px against 44px, boxed against underlined fields). Every tool WP snapshots in all skins present.
- **Rhino running on .NET Framework** (user changed `SetDotNetRuntime`) cannot load the plugin: document it.
- **Unknown TerrainTools failures.** WP-3A cannot start without Aksel's repro list.

## References

- McNeel, Moving to .NET Core: https://developer.rhino3d.com/guides/rhinocommon/moving-to-dotnet-core/
- `Panels.RegisterPanel`: https://developer.rhino3d.com/api/RhinoCommon/html/M_Rhino_UI_Panels_RegisterPanel.htm
- `SampleCsWpf`: https://github.com/mcneel/rhino-developer-samples/tree/6/rhinocommon/cs/SampleCsWpf
- Yak multi-targeted packages: https://developer.rhino3d.com/guides/yak/creating-a-multi-targeted-rhino-plugin-package/
- Rhino.Testing (optional, for NUnit-in-Rhino later): https://github.com/mcneel/Rhino.Testing
