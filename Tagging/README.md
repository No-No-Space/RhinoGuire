# Pregonero — Object Tagger

Tag objects with leaders whose text comes from a reusable template. Placeholder
keys are inserted as **live Rhino text fields**, so each leader reflects the
tagged object's own user text and updates automatically when that text changes.
It works like Revit's *tag by category* tool.

## Workflow

1. **Pick template object** — select an object that carries the user-text keys
   you want to tag with. Its keys (and current values) are listed.
2. **Build the leader template** — type any text and insert keys as `{KeyName}`
   (use *Insert {key}* / *Insert all keys*, or type them). Line breaks are kept.
   Example: `Room: {RoomName}` on one line, `Level {Level}` on the next.
3. **Choose the leader style** — pick a **Dimension Style** (controls font and
   base size), optionally set a **text-height override**, and toggle
   **Show leader arrowhead** on/off. Tick **No leader line (text only)** to
   place a plain text label at the second click instead of a leader (no line,
   no arrow) — the text still carries the same live fields and style.
4. **Start Tagging** — for each object, **click on the object** (the arrow lands
   at that point), then **click where the text goes**. Repeat for as many
   objects as you like.
5. **Press Enter** (or Esc) in the viewport to finish and return to the window.

## How keys are handled

Each `{Key}` becomes the live field `%<UserText("<object-guid>","Key")>%`.

- If a tagged object **already has** the key, the leader shows its value.
- If a tagged object is **missing** the key, Pregonero creates it on the object
  with the value **`TBD`** (a live field over a non-existent key would otherwise
  display `####`). Replace `TBD` later — every tag referencing it updates.

The keys that get the `TBD` treatment are exactly those referenced in the
template, so an object only gains the keys you actually tag it with.

## Notes

- The window is **modeless** — Rhino stays accessible while it is open.
- A whole tagging session is a single **Undo** step.
- Leaders are placed on the active viewport's construction plane, oriented so
  the text is readable in that view (best used in plan / Top views or layouts).
- No external dependencies.

## Reference

Text-field syntax: <https://docs.mcneel.com/rhino/8/help/en-us/information/text_fields.htm>
(`AttributeUserText` → `%<UserText("ObjectID","Key")>%`).
