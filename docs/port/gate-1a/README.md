# Gate 1A: AAJ skin review

> **Passed on 2026-11-27.** Aksel accepted the three recommendations: **40 px controls** (AAJ's default; the small styles for secondary actions in dense panels), the **hard shadow on the primary button**, and **all 31 proposed tokens** as listed below. These are the skin's defaults; `Compact` and `NoShadow` stay available as optional variants. In `skin-map.json` the entries keep `"proposed": true`, which now means "RhinoGuire value, not in AAJ, approved at Gate 1A".

Three decisions for Aksel before any tool panel is built. The images in this folder show the AAJ skin as it stands; they are rendered by `GalleryRenderTests` (or the `RGUiSnapshot` command in Rhino, which produces pixel-identical images) at 1.5× scale.

| Image | Shows |
| --- | --- |
| `gallery-light.png`, `gallery-dark.png` | Every control of the skin contract in every state: rest, hover, pressed, keyboard focus, disabled, field states, table row kinds, status kinds, chart, icons. |
| `panel-{light,dark}-40px-shadow.png` | A 360 px Lindero-style panel with AAJ's 40 px controls and the 4 px hard shadow on the primary button. |
| `panel-{light,dark}-40px-noshadow.png` | The same without the shadow. |
| `panel-{light,dark}-32px-shadow.png` | The compact variant: 32 px controls, tighter padding. |
| `panel-{light,dark}-32px-noshadow.png` | Compact without the shadow. |

Not visible in static images: open drop-downs and tooltips (they are popups). Both use the same surface, 2 px border and no shadow as the rest of the skin.

## Decision 1: control density

AAJ's controls are 40 px high (`bundle.css`: buttons and fields `min-height: 40px`). The compact variant uses 32 px controls, 28 px table rows and tighter padding.

In the sample panel the compact variant saves about 7 % of the height (851 against 920 px). The table shows one more row; the fields leave little air around their 15 px text.

**Recommendation: keep 40 px** as AAJ specifies, and use the small styles (`Tool.Button.*.Small`, 32 px) for secondary actions where a panel is dense. Switching later is one setting: the compact values are a variant file, not separate styles.

## Decision 2: hard shadow on the primary button

AAJ allows an optional 4 px solid shadow on buttons, which sinks into place on press. In the skin it is used on `Tool.Button.Primary` only; secondary, text and icon buttons never carry it.

**Recommendation: keep it.** It marks the one main action of a panel without any extra colour, and it is AAJ's signature. Removing it is again one variant (`NoShadow`).

## Decision 3: proposed tokens

AAJ is a brand system with no tool-level tokens, so some contract keys needed values it does not define. Each is marked `"proposed": true` in `src/RhinoGuire.UI/Skins/AAJ/skin-map.json`, with the contrast computed by `tools/build_aaj_skin.py` (`--report` prints this table). Approved values can go back into the AAJ design system's own `tokens.json`.

Most worth a look:

- **Status colours.** AAJ's semantic colours are too light for small text on its surfaces (`color-success` 2.8:1, `color-warning` 2.1:1 on `bg-surface`). The strong values are deeper or lifted versions of the same hues, all above 4.5:1.
- **Alternate table rows** (`TableRowAlt`) differ from plain rows by only 1.05:1 in light. Subtle by intent; say if you want them stronger.
- **Activity and selection share Mar Caribe**, as AAJ says. `SelectionSubtle` therefore equals `ActivitySubtle`.
- **Text button.** AAJ has none. Proposed: no fill or border, with a 2 px Mar Caribe rule under the label on hover and keyboard focus.
- **Diverging chart colours in light**: `ChartDivergingLow` (Salmón) is 2.95:1 on the chart surface, just under 3:1 for graphics. It is used for the bullet chart's tolerance zone and cut on terrain maps, always next to labels.

| Key | Light | Dark | Contrast (light / dark) | Note |
| --- | --- | --- | --- | --- |
| `ActivitySubtle` | `#C3DDE2` (seq-mar-1) | `#3B4A4A` | TextPrimary 10.43 / 7.96; SurfaceSubtle 1.35 / 1.37 | Tint that marks a changed value. Dark: Mar Caribe on-dark mixed 20 % into bg-surface-dark (new value). |
| `SelectionSubtle` | `#C3DDE2` (seq-mar-1) | `#3B4A4A` | TextPrimary 10.43 / 7.96 | Equals ActivitySubtle: AAJ uses Mar Caribe for both. |
| `StatusSuccessSubtle` | `#E8EFE6` | `#3D493D` | TextPrimary 12.66 / 8.13 | color-success mixed into the surface: 12 % light, 20 % dark. |
| `StatusSuccessStrong` | `#2F6E4A` | `#7CC49A` | SurfaceSubtle 5.79 / 6.18; SurfacePage 5.23 / 8.62; StatusSuccessSubtle 5.19 / 4.61 | color-success reaches only 2.8:1 on bg-surface; deeper (light) or lifted (dark) green of the same hue. |
| `StatusWarningSubtle` | `#F8EFDF` | `#594831` | TextPrimary 13.00 / 7.53 | color-warning mixed into the surface: 12 % light, 20 % dark. |
| `StatusWarningStrong` | `#8A5A12` | `#EAB866` | SurfaceSubtle 5.62 / 6.99; SurfacePage 5.08 / 9.73; StatusWarningSubtle 5.18 / 4.82 | color-warning reaches 2.1:1 on bg-surface; deeper amber in light, lifted in dark. |
| `StatusErrorSubtle` | `#F4E4DF` | `#4F352F` | TextPrimary 12.02 / 9.58 | color-error mixed into the surface: 12 % light, 18 % dark. |
| `StatusErrorStrong` | `#A83A30` | `#EE8E78` (color-salmon-on-dark) | SurfaceSubtle 6.03 / 5.31; SurfacePage 5.44 / 7.40; StatusErrorSubtle 5.13 / 4.66 | color-error drops under 4.5:1 on the subtle fill; one step deeper. Dark reuses Salmón on-dark. |
| `StatusInfoSubtle` | `#E6EAEE` | `#3A414A` | TextPrimary 12.27 / 8.86 | color-info mixed into the surface: 12 % light, 20 % dark. |
| `StatusInfoStrong` | `#3A689E` | `#8DB0DA` | SurfaceSubtle 5.47 / 5.66; SurfacePage 4.94 / 7.89; StatusInfoSubtle 4.75 / 4.60 | Cielo Llanero one step deeper (light) or lighter (dark). |
| `InputBgHover` | `#FBF9F5` (bg-surface) | `#36322E` (bg-surface-dark) | TextPrimary 14.11 / 10.91 | AAJ fields have no hover fill; the 2 px border is the affordance. |
| `InputBgFocus` | `#FBF9F5` (bg-surface) | `#36322E` (bg-surface-dark) | TextPrimary 14.11 / 10.91 | AAJ marks focus with the 2 px Mar Caribe ring only. |
| `ButtonTextHover` | `#2A2725` (text-primary) | `#F2EDE6` (text-primary-dark) | ButtonSecondaryBgHover 11.42 / 8.80 | Text button label stays Carbón (Hueso). |
| `ButtonTextPressed` | `#2A2725` (text-primary) | `#F2EDE6` (text-primary-dark) | ButtonSecondaryBgPressed 9.76 / 7.38 | |
| `ButtonTextUnderline` | `#2A8B9C` (color-mar-caribe) | `#4DA8B8` (color-mar-caribe-on-dark) | SurfacePage 3.42 / 6.41 | The rule under a text button on hover and focus. |
| `TableRowAlt` | `#F6F3EE` | `#3C3833` | TextPrimary 13.41 / 9.99; TableRowBg 1.05 / 1.09 | Half a step between the row fill and the next neutral. |
| `TableCellChanged` | `#C3DDE2` (seq-mar-1) | `#3B4A4A` | TextPrimary 10.43 / 7.96 | Equals ActivitySubtle: a changed value is activity. |
| `TableSubtotalBg` | `#EDE7DE` | `#45403B` (state-secondary-hover-dark) | TextPrimary 12.07 / 8.80; TableRowBg 1.17 / 1.24 | Light: between bg-page and row-hover. Dark: one step lighter than the row fill. |
| `TableSubtotalRule` | `#B0ABA5` (seq-neutro-2) | `#A8A39E` (text-secondary-dark) | TableSubtotalBg 1.85 / 4.10 | 1 px rule above a subtotal row. |
| `TableTotalBg` | `#D5D1CB` (seq-neutro-1) | `#514B45` (state-secondary-pressed-dark) | TextPrimary 9.76 / 7.38; TableRowBg 1.45 / 1.48 | Two steps from the row fill, apart from subtotal and hover. |
| `TableTotalRule` | `#2A2725` (color-carbon) | `#F2EDE6` (color-hueso) | - | 2 px rule above the total row, in the border colour. |
| `ChartGrid` | `#D5D1CB` (chrome-grid) | `#57534F` (seq-neutro-4) | ChartSurface 1.45 / 1.67 | AAJ gives chrome-grid for light only. |
| `ChartDivergingLow` | `#E0735C` (div-0) | `#EE8E78` (color-salmon-on-dark) | ChartSurface 2.95 / 5.31 | AAJ's diverging scale is light-only; dark uses the lifted accents. |
| `ChartDivergingMid` | `#F2EDE6` (div-2) | `#57534F` (seq-neutro-4) | ChartSurface 1.11 / 1.67 | Hueso would vanish on a dark chart. |
| `ChartDivergingHigh` | `#4A7DB8` (div-4) | `#6E97C9` (color-cielo-llanero-on-dark) | ChartSurface 4.06 / 4.20 | |
| `AajScrollThumb` | `#B0ABA5` (seq-neutro-2) | `#57534F` (seq-neutro-4) | SurfacePage 1.96 / 2.32 | AAJ has no scroll bar: square thumb, no track, no arrows. |
| `AajScrollThumbHover` | `#8A8580` (color-tierra) | `#A8A39E` (text-secondary-dark) | SurfacePage 3.14 / 7.08 | |
| `ControlHeightSmall` | 32 | same | - | AAJ has no small control; for secondary actions in dense panels. |
| `TableIndent` | 16 | same | - | Indent per tree level in tables. |
| `AajFontSizeButtonSmall` | 13 | same | - | |
| `AajButtonPaddingSmall` | 12,5 | same | - | |

## How to answer

Reply with the three choices (density, shadow, and any token you want changed). Changing a token means editing its entry in `skin-map.json` and running `python tools/build_aaj_skin.py`; the tests then check the result.
