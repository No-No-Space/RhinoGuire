# Port Status

Tracking document for the RhinoGuire C# / WPF port. Updated at the start and end of each session.

## Work packages

| WP | Description | Owner | State | Branch | Open questions / Notes |
| --- | --- | --- | --- | --- | --- |
| **WP-0** | Bootstrap and spikes | Antigravity, finished by Claude | Done | `feat/csharp-plugin` | Gate 0 passed. Results, decisions and dev loop in `SPIKE.md` |
| **WP-1A** | Skin contract, shared controls, AAJ skin | Claude (started by a background agent) | In review: PR #9, Gate 1A passed | `wp-1a-skin` | Ready to merge. `SlopeInput` follows once WP-1B is merged |
| **WP-1B** | Core logic and tests | Claude (started by a background agent) | In review: PR #8 | `wp-1b-core` | 900 tests; golden data regenerated from the Python tools |
| **WP-1C** | Plugin shell, services, parity harness | — | Ready | — | Give panels a sensible size on first open (Gate 0 note, `SPIKE.md`) |
| **WP-1D** | White skin | — | Pending | — | Blocked by Gate 1A |
| **WP-2A** | Baquiano (pilot) | — | Pending | — | Blocked by Gate 1A, 1B, 1C |
| **WP-2B** | Terrain engine + Sebucan | — | Pending | — | Blocked by 1C |
| **WP-2C** | Pregonero | — | Pending | — | Blocked by 2A |
| **WP-2D** | Arriero | — | Pending | — | Blocked by 2A |
| **WP-2E** | Chivito | — | Pending | — | Blocked by 2A |
| **WP-2F** | PadGrader, WayGrader, CutFillReport | — | Pending | — | Blocked by 2B |
| **WP-2G** | Trocha | — | Pending | — | Blocked by 2B |
| **WP-2H-1** | Lindero engine + Excel | — | Pending | — | Blocked by 1B, 1C |
| **WP-2H-2** | Lindero UI | — | Pending | — | Blocked by 1A, 2H-1 |
| **WP-3A** | TerrainTools reliability | — | Pending | — | Blocked by Gate 2 & Aksel's repro list |
| **WP-3B** | Trocha finish | — | Pending | — | Blocked by Gate 2 |
| **WP-3C** | Lindero finish | — | Pending | — | Blocked by Gate 2 |
| **WP-4A** | Packaging | — | Pending | — | Blocked by Gate 3 |
| **WP-4B** | Documentation | — | Pending | — | Blocked by Gate 3 |
| **WP-4C** | Retire Python | — | Pending | — | Blocked by Gate 4 |

## Gates

- **Gate 0:** Passed 2026-11-25. Manual checks all succeeded; `net7.0-windows`, RhinoCommon floor 8.21 and on-demand loading accepted. Note: panels open small (follow-up in WP-1C and the tool WPs)
- **Gate 1A:** Passed 2026-11-27. 40 px controls, hard shadow on the primary button, all 31 proposed tokens accepted (`docs/port/gate-1a/README.md` on PR #9). WP-1D (internal skin) and the tool ports may start once PR #9 is merged
- **Gate 1D:** Pending (White gallery)
- **Gate 2:** Pending (Tool smoke checklists on live models)
- **Gate 3:** Pending (Finished features & naming)
- **Gate 4:** Pending (Retire Python & merge to main)
