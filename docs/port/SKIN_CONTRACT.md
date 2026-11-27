# RhinoGuire skin contract

RhinoGuire's panels are drawn by a **skin**: a set of WPF resource dictionaries that define colours, sizes and control styles. The plugin ships one skin, **AAJ**, compiled in. Anyone can write another one as a folder of loose XAML files and drop it next to the plugin.

This document is the specification. It lists every resource key a skin provides, what each key is for, and the rules panels and skins follow. The same list lives in code as `RhinoGuire.UI.SkinContract`; a unit test keeps the two in step.

## How a skin is loaded

A skin has up to four files, all optional for an external skin:

| File | Holds |
| --- | --- |
| `Tokens.Shared.xaml` | Theme-independent metrics: fonts, sizes, borders, spacing. |
| `Tokens.Light.xaml`, `Tokens.Dark.xaml` | Colours, one file per theme, with the same keys in both. |
| `Tokens.<Variant>.xaml` | Optional overrides selected by name (AAJ ships `Compact` and `NoShadow`). |
| `Controls.xaml` | Control styles and templates: the named `Tool.*` styles and the implicit (type) styles. |

Layering, later wins:

1. The embedded AAJ skin always loads first: `Tokens.Shared`, the theme file, any selected variants, `Controls`. Every key therefore resolves even if an external skin leaves it out.
2. An external skin's files load on top, in the same order. Redefine only what you change: a skin that only recolours ships just `Tokens.Light.xaml` and `Tokens.Dark.xaml`.

External skins are found in sub-folders of, in priority order: folders the plugin registers, `<RhinoGuire.UI folder>\Skins`, and `%APPDATA%\RhinoGuire\Skins`. The folder name is the skin name. A `default.txt` in a Skins folder whose first line names a skin preselects it for that installation.

Loose XAML is loaded with the skin folder as base URI, so a skin can ship its own fonts and reference them by relative path:

```xml
<FontFamily x:Key="FontBody">./Fonts/#Mulish</FontFamily>
```

Loose XAML cannot have code-behind (`x:Class`). Every behaviour lives in RhinoGuire.UI's controls; a skin only supplies looks. Loose files may reference those controls:

```xml
xmlns:ctl="clr-namespace:RhinoGuire.UI.Controls;assembly=RhinoGuire.UI"
```

Check a skin with `SkinContract.FindMissing(SkinManager.BuildSkinDictionary(dark, folder))`: it returns every key that is missing or has the wrong type.

## Rules

**For panels (tool XAML).**

- Reference skin keys with `{DynamicResource …}`, never `StaticResource`: theme and skin changes swap the dictionaries under an open panel.
- No literal colours, font families or font sizes. Use the colour, metric and `Tool.Text.*` keys below. A unit test (`XamlLintTests`) enforces this.
- Plain controls (`Button`, `TextBox`, `ComboBox` …) take the skin's implicit styles. Opt into a named style only for a variant: `Tool.Button.Primary` for the panel's one main action, `Tool.Field.Number` for figures, `Tool.Switch` for settings that apply at once.
- Type labels that a skin shows in capitals (column headers, eyebrows) in capitals: WPF has no text-transform.

**For skins (Controls.xaml).**

- Give every control in the implicit-style list its own `ControlTemplate`. WPF's default chrome (rounded blue hovers, grey scroll bars) must never show in a panel.
- Draw the keyboard focus ring inside the template, triggered by `ctl:FocusVisible.IsVisible`. A Rhino panel has no adorner layer, so `FocusVisualStyle` is not drawn; set it to `{x:Null}`.
- Trigger hover and pressed looks on `ctl:StatePreview.State` (`Hover`, `Pressed`) as well as `IsMouseOver` / `IsPressed`, so the gallery can show every state in a static snapshot.
- Status is never colour alone: pair it with an icon and a word (`StatusLine` and `Field` do this).
- Several button looks can share one template: set `ctl:ControlColors.HoverBackground`, `PressedBackground`, `PressedForeground`, `CheckedBackground` and `CheckedForeground` in each style and bind to them from the template's triggers.

## Colours

Defined in `Tokens.Light.xaml` and `Tokens.Dark.xaml`. Each name `X` is defined twice: `XColor` (a `Color`) and `XBrush` (a `SolidColorBrush`). Panels use the brushes.

### Surfaces, text, borders, focus

| Key | Role |
| --- | --- |
| `SurfacePage` | Panel ground. |
| `SurfaceSubtle` | Raised surfaces on the ground: cards, tab panels, drop-downs. |
| `TextPrimary` | Body text. |
| `TextSecondary` | Labels, help text, captions. |
| `TextDisabled` | Text of disabled content that is not inside a control. |
| `TextInverse` | Text on a `TextPrimary`-coloured fill. |
| `BorderSubtle` | Hairlines and tracks that should recede. |
| `BorderControl` | Control outlines. |
| `BorderStrong` | Emphasised outlines and rules. |
| `FocusRing` | Keyboard focus ring. |

### Interface colour

Primary controls: the panel's main button, checked boxes, the active tab.

| Key | Role |
| --- | --- |
| `UiDefault` | Fill at rest. |
| `UiHover` | Fill on hover. |
| `UiPressed` | Fill while pressed. |
| `UiSubtle` | Light hover fill for neutral controls (list items, unchecked boxes). |
| `UiOn` | Text and marks on `UiDefault`. |
| `UiOnPressed` | Text on `UiPressed`. |

### Activity and selection

Activity means something runs or changed (a busy bar, an edited value). Selection means chosen rows and items. Skins may give them the same hue; panels must not use either as decoration.

| Key | Role |
| --- | --- |
| `ActivityDefault` | Running work: busy bar, switch on, text selection. |
| `ActivitySubtle` | Fill of a changed, unsaved value. |
| `ActivityStrong` | Text, icon and rule of a changed value; busy status word. |
| `SelectionDefault` | Fill of a selected item. |
| `SelectionSubtle` | Light selection fill (secondary selection). |
| `SelectionStrong` | Selection outline or text on a light ground. |
| `SelectionOn` | Text on `SelectionDefault`. |

### Status

`…Subtle` fills a field or message; `…Strong` carries its text, icon and rule.

| Key | Role |
| --- | --- |
| `StatusSuccessSubtle` | Success fill. |
| `StatusSuccessStrong` | Success text, icon, rule. |
| `StatusWarningSubtle` | Warning fill. |
| `StatusWarningStrong` | Warning text, icon, rule. |
| `StatusErrorSubtle` | Error fill. |
| `StatusErrorStrong` | Error text, icon, rule. |
| `StatusInfoSubtle` | Information fill. |
| `StatusInfoStrong` | Information text and icon. |

### Disabled, inputs, buttons

| Key | Role |
| --- | --- |
| `DisabledBg` | Fill of a disabled control. |
| `DisabledFg` | Text of a disabled control. |
| `DisabledBorder` | Outline of a disabled control. |
| `InputBg` | Field, select and checkbox fill at rest. |
| `InputBgHover` | Field fill on hover. |
| `InputBgFocus` | Field fill while focused. |
| `ButtonSecondaryBg` | Secondary button fill. |
| `ButtonSecondaryBorder` | Secondary button outline. |
| `ButtonSecondaryBgHover` | Secondary button fill on hover. |
| `ButtonSecondaryBgPressed` | Secondary button fill while pressed. |
| `ButtonDisabledBorder` | Outline of a disabled button. |
| `ButtonTextHover` | Text button label on hover. |
| `ButtonTextPressed` | Text button label while pressed. |
| `ButtonTextUnderline` | Rule under a text button on hover and focus. |

### Tables

Row kinds must read apart: header › total › subtotal › data rows.

| Key | Role |
| --- | --- |
| `TableHeaderBg` | Header row fill. |
| `TableHeaderFg` | Header row text. |
| `TableRowBg` | Data row fill. |
| `TableRowAlt` | Every second data row. |
| `TableRowHover` | Row under the pointer. |
| `TableRowSelected` | Selected row fill. |
| `TableRowSelectedFg` | Selected row text. |
| `TableRule` | Hairline between rows. |
| `TableCellChanged` | Fill of a changed cell. |
| `TableSubtotalBg` | Subtotal row fill. |
| `TableSubtotalRule` | Rule above a subtotal row. |
| `TableTotalBg` | Total row fill. |
| `TableTotalRule` | Rule above the total row. |

### Charts

Series keep their order in both themes; a sixth series folds into "other" or small multiples. Text never takes a series colour.

| Key | Role |
| --- | --- |
| `ChartSeries1` | First series (bullet chart: the measured bar). |
| `ChartSeries2` | Second series. |
| `ChartSeries3` | Third series. |
| `ChartSeries4` | Fourth series. |
| `ChartSeries5` | Fifth series. |
| `ChartGrid` | Grid lines and the empty range of a bar. |
| `ChartAxis` | Axes, baselines, goal marks. |
| `ChartLabel` | Tick labels and secondary chart text. |
| `ChartText` | Values and legends. |
| `ChartSurface` | Chart ground, also for PNG export. |
| `ChartTooltipBg` | Tooltip fill. |
| `ChartTooltipFg` | Tooltip text. |
| `ChartCrosshair` | Hover crosshair. |
| `ChartFocus` | The one highlighted series, at most one per view. |
| `ChartMuted` | Series that are not highlighted. |
| `ChartDivergingLow` | Below a baseline: cut on terrain maps, the below-goal tolerance zone of a bullet chart. |
| `ChartDivergingMid` | The baseline itself. |
| `ChartDivergingHigh` | Above a baseline: fill on terrain maps, the above-goal tolerance zone of a bullet chart. |

## Metrics

Defined in `Tokens.Shared.xaml` (or a variant file), with the value type a skin must use.

| Key | Type | Role |
| --- | --- | --- |
| `FontHeading` | `FontFamily` | Titles, section headers, buttons, tabs. |
| `FontBody` | `FontFamily` | Body and control text. |
| `FontMono` | `FontFamily` | Figures, labels, table headers, captions. |
| `FontSizeH1` | `Double` | Largest heading (rarely used in panels). |
| `FontSizeH2` | `Double` | Panel title (`Tool.Text.Title`). |
| `FontSizeH3` | `Double` | Section header. |
| `FontSizeBody` | `Double` | Body text. |
| `FontSizeBodySm` | `Double` | Help text, status messages. |
| `FontSizeMono` | `Double` | Figures in text and number fields. |
| `FontSizeTable` | `Double` | Table body. |
| `FontSizeEyebrow` | `Double` | Mono labels and table headers. |
| `FontSizeCaption` | `Double` | Captions. |
| `FontWeightHeading` | `FontWeight` | Headings and buttons. |
| `FontWeightBody` | `FontWeight` | Body text. |
| `FontWeightStrong` | `FontWeight` | Emphasis, subtotal and total rows, status words. |
| `FontWeightFigure` | `FontWeight` | Figures. |
| `FontWeightEyebrow` | `FontWeight` | Small mono labels. |
| `ControlBorderWidth` | `Double` | Control outline width. |
| `ControlBorderWidthThickness` | `Thickness` | The same, for `BorderThickness`. |
| `LineThickness` | `Double` | Rules (`Tool.Rule`, total row). |
| `LineThicknessThickness` | `Thickness` | The same as a `Thickness`. |
| `LineThicknessThin` | `Double` | Hairlines (subtotal rule). |
| `LineThicknessThinThickness` | `Thickness` | The same as a `Thickness`. |
| `FocusRingWidth` | `Double` | Focus ring width. |
| `FocusRingWidthThickness` | `Thickness` | The same as a `Thickness`. |
| `FocusRingOffset` | `Double` | Gap between control and focus ring. |
| `FocusRingOffsetThickness` | `Thickness` | The same as a `Thickness`. |
| `Radius` | `CornerRadius` | Corner radius of every container. |
| `ControlHeight` | `Double` | Minimum height of buttons, fields, selects. |
| `ControlHeightSmall` | `Double` | Small buttons, checkbox rows. |
| `TableRowHeight` | `Double` | Minimum table row height. |
| `TableHeaderHeight` | `Double` | Minimum table header height. |
| `TableIndent` | `Double` | Indent per tree level in tables. |
| `IconSize` | `Double` | Default icon size. |
| `IconSizeLarge` | `Double` | Large icons. |
| `IconStrokeWidth` | `Double` | Icon stroke width, independent of size. |
| `Space1` | `Double` | 4-step spacing scale, smallest. |
| `Space2` | `Double` | Spacing. |
| `Space3` | `Double` | Spacing. |
| `Space4` | `Double` | Spacing. |
| `Space6` | `Double` | Spacing. |
| `Space8` | `Double` | Spacing, largest. |
| `PanelPadding` | `Thickness` | Padding inside a docked panel and a card. |

## Named styles

Defined in `Controls.xaml`. Panels refer to them with `Style="{DynamicResource …}"`.

| Key | Target | Use |
| --- | --- | --- |
| `Tool.Text.Title` | `TextBlock` | Panel title. |
| `Tool.SectionHeader` | `TextBlock` | Section header inside a panel. |
| `Tool.Text.Body` | `TextBlock` | Body text, wraps. |
| `Tool.Text.Secondary` | `TextBlock` | Help and secondary text. |
| `Tool.Text.Eyebrow` | `TextBlock` | Small capitals label above a group. |
| `Tool.Text.Figure` | `TextBlock` | A number in text: mono, tabular. |
| `Tool.Text.Caption` | `TextBlock` | Caption, small mono. |
| `Tool.Rule` | `Separator` | Section rule. |
| `Tool.Card` | `Border` | A raised surface grouping content. |
| `Tool.Button.Primary` | `Button` | The panel's one main action. |
| `Tool.Button.Primary.Small` | `Button` | The same, small. |
| `Tool.Button.Secondary` | `Button` | Every other action (also the implicit `Button` style). |
| `Tool.Button.Secondary.Small` | `Button` | The same, small. |
| `Tool.Button.Text` | `Button` | Low-weight actions ("Export Excel", "Show all"). |
| `Tool.Button.Text.Small` | `Button` | The same, small. |
| `Tool.Button.Icon` | `Button` | Square button holding a `ctl:Icon`; give it a `ToolTip`. |
| `Tool.Button.Icon.Small` | `Button` | The same, small. |
| `Tool.ToggleButton` | `ToggleButton` | On/off choice shown as a button. |
| `Tool.Field` | `TextBox` | Text entry (also the implicit `TextBox` style). Placeholder: `ctl:Field.Placeholder`. |
| `Tool.Field.Number` | `TextBox` | Number entry: right-aligned, mono. Put the unit in the label. |
| `Tool.FieldFrame` | `Field` | Label, input, then a state message or help text (`ctl:Field`). |
| `Tool.Select` | `ComboBox` | One value from a list (also `KeySearchComboBox`). |
| `Tool.Select.Item` | `ComboBoxItem` | Item of a select's drop-down. |
| `Tool.CheckBox` | `CheckBox` | Independent option, including the mixed state. |
| `Tool.Radio` | `RadioButton` | One of a few visible options. |
| `Tool.Switch` | `CheckBox` | Setting that applies at once. |
| `Tool.Tabs` | `TabControl` | Views inside a panel. |
| `Tool.Tab` | `TabItem` | A tab. |
| `Tool.ListBox` | `ListBox` | A list to choose from. |
| `Tool.ListBox.Item` | `ListBoxItem` | Its items. |
| `Tool.Expander` | `Expander` | Disclosure for advanced settings. |
| `Tool.ScrollBar` | `ScrollBar` | Scroll bars. |
| `Tool.ScrollViewer` | `ScrollViewer` | Scrolling regions (keep the part name `PART_VerticalScrollBar`). |
| `Tool.ToolTip` | `ToolTip` | Tooltips. |
| `Tool.Table` | `TreeDataGrid` | Tables. Template parts: `PART_HeaderHost` (a `Decorator`) and `PART_ScrollViewer`. |
| `Tool.Table.Row` | `TreeDataGridRow` | Rows; style by `Kind` (`Data`, `Subtotal`, `Total`), `IsAlternate`, hover, selection. |
| `Tool.Table.Cell` | `TreeDataGridCell` | Cells; `IsNumeric`, `IsChanged`, `Kind`. |
| `Tool.Table.HeaderCell` | `TreeDataGridHeaderCell` | Column headers; `IsNumeric`. |
| `Tool.Table.Expander` | `ToggleButton` | Expand/collapse chevron of a tree row. |
| `Tool.StatusLine` | `StatusLine` | Icon, status word and message. |
| `Tool.BusyBar` | `BusyBar` | Progress. Template parts: `PART_Track` and `PART_Indicator`. |

## Implicit styles

`Controls.xaml` must also define a type-keyed style for each of these, so nothing in a panel falls back to WPF's default look:

`Button`, `ToggleButton`, `TextBox`, `ComboBox`, `ComboBoxItem`, `CheckBox`, `RadioButton`, `TabControl`, `TabItem`, `ListBox`, `ListBoxItem`, `Expander`, `Separator`, `ScrollBar`, `ScrollViewer`, `ToolTip`, `Field`, `StatusLine`, `BusyBar`, `KeySearchComboBox`, `TreeDataGrid`, `TreeDataGridRow`, `TreeDataGridCell`, `TreeDataGridHeaderCell`.

## Skin-private keys

A skin may define extra keys for its own `Controls.xaml`. AAJ's are prefixed `Aaj` (for example `AajShadowBrush`, `AajButtonPadding`, `AajButtonShadowVisibility`). Panels never reference skin-private keys.

## A minimal external skin

A folder `%APPDATA%\RhinoGuire\Skins\Forest\` holding only `Tokens.Light.xaml`:

```xml
<ResourceDictionary xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
                    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml">
    <Color x:Key="UiDefaultColor">#2F5D46</Color>
    <SolidColorBrush x:Key="UiDefaultBrush" Color="#2F5D46" />
    <Color x:Key="FocusRingColor">#2F5D46</Color>
    <SolidColorBrush x:Key="FocusRingBrush" Color="#2F5D46" />
</ResourceDictionary>
```

recolours AAJ's primary controls and focus ring in the light theme and keeps everything else.
