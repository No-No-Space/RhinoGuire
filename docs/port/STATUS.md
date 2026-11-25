# Port Status

Tracking document for the RhinoGuire C# / WPF port. Updated at the start and end of each session.

## Work packages

| WP | Description | Owner | State | Branch | Open questions / Notes |
| --- | --- | --- | --- | --- | --- |
| **WP-0** | Bootstrap and spikes | Antigravity, finished by Claude | Done, awaiting Gate 0 | `feat/csharp-plugin` | All automated spike checks pass; 3 manual checks for Aksel. See `SPIKE.md` |
| **WP-1A** | Skin contract, shared controls, AAJ skin | — | Pending | — | Blocked by Gate 0 |
| **WP-1B** | Core logic and tests | — | Pending | — | Blocked by Gate 0 |
| **WP-1C** | Plugin shell, services, parity harness | — | Pending | — | Blocked by Gate 0 |
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

- **Gate 0:** Ready for review: `SPIKE.md` (results, target framework, RhinoCommon floor) and its three manual checks
- **Gate 1A:** Pending (AAJ gallery, proposed tokens, density)
- **Gate 1D:** Pending (White gallery)
- **Gate 2:** Pending (Tool smoke checklists on live models)
- **Gate 3:** Pending (Finished features & naming)
- **Gate 4:** Pending (Retire Python & merge to main)
