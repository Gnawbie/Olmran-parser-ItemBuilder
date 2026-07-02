# Gaming Log Parser v2.6

A tool for parsing Olmran/MUD-style game logs, extracting loot data, and finding optimal equipment builds.

For installation steps, see **INSTALL_INSTRUCTIONS.txt**. This file covers how to use the program once it's running.

## The Four Tabs

### Parse Tab
- Load CHAT logs and ACTION logs (auto-detects type from filename)
- **Search Logs** — find every time an item actually dropped across a large batch of logs, independent of Run Parse:
  - **Drops only** (default, on) — matches real drop events the same way the Loot parser does, so quality/material prefixes ("bright", "glowing", "shining", etc.) don't cause misses or false confidence. For each drop, shows a **snapshot**: every line from the last timestamp seen through the drop line itself, so you can see what led up to it (the fight, the kill, then the drop). Click a result row to view its snapshot below the list.
  - Uncheck **Drops only** to instead do a plain raw-text search of every loaded file's lines (any mention of the term, not just drops), showing filename, line number, and timestamp per match.
  - Case-sensitive matching is optional in both modes.
- Parse options (Chat / Combat / Loot) automatically enable or disable based on which file types are loaded
- Use the Snapshot buttons to preview parsed data before exporting
- Export everything to a formatted Excel workbook

### Fields Tab
- Customize which fields get extracted for Chat, Combat, and Loot
- Add, edit, remove, and reorder fields per data type
- "Reset to Defaults" restores the built-in field set

### Export Tab
- Create a new Master Loot Database or append to an existing one (fodder items excluded automatically)
- Export Combat and Loot as separate files if needed
- Optional summary sheet with run statistics
- Each chat log gets its own sheet on export

### Build Tab (the item/equipment builder)
This is where you search a master database for the best equipment for a character.

**Master Database File**
- **Browse...** — pick your own `.xlsx` database
- **Create New** — start a blank one
- **Use Community List** — instantly loads the bundled `Olmran_Community_Eq_and_Stats_List.xlsx`
- **Load** — reads the selected file into memory for searching

**Desired Spells**
Instead of typing spell names, pick them from category dropdowns, each paired with a tier dropdown (`(any)`, `i`, `ii`, `iii`):
- **Basic** — Agility, Dexterity, Constitution, Intelligence, Wisdom, Strength, Bless, Evade, Combat
- **Shields/Buffs** — Protect, Blur, Shield, Tough.skin, Vitalize, Regenerate (more being added over time)
- **Class Specific / Other1 / Other2** — reserved for spell lists that haven't been compiled yet

  Some spells restrict which tiers are selectable (e.g. Agility and Bless only go up to tier `ii`; the dropdown greys out `iii` automatically).

- **Manual** entry (temporary) — a plain text box for adding any spell not yet in a category dropdown above. This exists only until the Class Specific / Other1 / Other2 lists are filled in, at which point it will be removed.

Click **Add to List** next to whichever dropdown (or the Manual box) you used. Added spells show up as removable chips under **Wanted Spells**, flowing left-to-right and wrapping to new lines as needed — click the ✕ on a chip to remove it, or **Clear All** to start over.

Requesting the same spell at two different tiers (e.g. Dexterity i and Dexterity iii) is treated as one requirement, not two — the search picks whichever single item best satisfies it rather than trying to fill two equipment slots for the same stat.

**Level Filtering**
- **Min Level** / **Max Level** — restrict to a level range
- **Specific Level** — restrict to an exact level
- Fields grey out automatically depending on which option you're using (Specific vs. Min/Max are mutually exclusive)
- Leave all three blank for no level restriction

**Armor Constraints**
- The **All:** row sets an armor type for every slot at once
- Override individual slots (Head, Cloak, Body, Hands, Legs, Feet) with Cloth / Leather / Studded / Plate
- **Set as Default** / **Clear Default** / **Clear All** manage your saved preferences

**Weapon Constraints**
- **Weapon Style:** Melee, Direct (Caster), Parry Staff *(not yet implemented)*, or Any
- **Build Config:** Weapon, Shield, Two-Handed, 1 Claw, 2 Claw — check whichever your build actually uses (some classes can only use one claw, others can dual-wield two)
- **Dual-Wield** sub-option, for finding 1-handed weapons for both hands
- **Damage Type:** Slashing, Thrusting, Crushing

**Only Found In (Realm filter)**
Next to the spell dropdowns, check any combination of Evil, Chaos, Good, Kaid, Crafted, Glory Bea, or Event to restrict results to items found in those realms. Leave everything unchecked to search all realms.

A build can include **at most one Crafted-realm item** — this is enforced automatically during the search, not just a filter you have to remember.

**Search**
- **🎯 Find Optimal Build** — greedy search for the single best item per slot, covering as many wanted spells as possible under your constraints
- **📋 Show All Matches** — lists every item that matches your filters, without narrowing to one per slot
- **🎲 Generate multiple build options** (checkbox) — when checked, "Find Optimal Build" also generates up to 5 alternate full builds by swapping in equally-good "tied" items slot by slot, so you can compare a few options instead of just one

## Results Tab
- **Display: Best Per Slot / All Matches** — toggle between the two search modes above
- **Build Variant** dropdown — only active when "Generate multiple build options" produced more than one build; switches which variant is shown
- Results table columns: Slot, Item, Type, Spell, Level, Mob, Area, a thin divider, and **Alt Options** (other items that tied for that slot, if any)
- **Remove Area** — strip all items from a given area out of the current results (useful for excluding an event/expansion you don't want)
- **Export Results** — save the current results table to Excel (the divider column is left out automatically)

## Tips

**Build Tab quick start**
1. Click "Use Community List" (or load your own database)
2. Add the spells you want from the category dropdowns
3. Set level, armor, weapon, and realm constraints as needed
4. Click "Find Optimal Build"

**Example builds**
- **Tank:** Melee + Plate armor + Weapon + Shield + Slashing + Min Level 60
- **Two-Handed Warrior:** Melee + Plate armor + Two-Handed + Crushing + Level 50-70
- **Dual-Wield Rogue:** Melee + Leather armor + Weapon + Dual-Wield + Thrusting
- **Dual-Claw Fighter:** Melee + 2 Claw + Slashing
- **Caster with Staff:** Direct (Caster) + Cloth armor + Two-Handed + Crushing

## Known Limitations
- **Parry Staff** weapon style is not yet implemented
- **Class Specific**, **Other1**, and **Other2** spell dropdowns are empty until those lists are compiled — use the temporary Manual entry box in the meantime
- **Event** realm items work like any other realm filter, but there isn't yet dedicated Event-specific logic beyond the Realm column match

## Troubleshooting

**Program won't start** — see INSTALL_INSTRUCTIONS.txt.

**Files not loading** — check the file extension (.log, .txt, .xlsx), make sure it isn't open in another program, and try copying it somewhere else if the path has unusual characters.

**Parse results empty** — verify the log actually contains delve output ("You examine X closely") and that CHAT vs ACTION auto-detection picked the right type (check the Parse tab file list).

**Community List not found** — make sure `Olmran_Community_Eq_and_Stats_List.xlsx` is in the same folder as `gaming_log_parser.py`, or use "Browse..." to point at it manually.

## Version History

### v2.6 (Current)
- Added "Search Logs" to the Parse tab, with a "Drops only" mode: finds every real drop event of an item (using the same drop-detection and prefix-cleaning as the Loot parser) and shows a snapshot from the last timestamp through the drop line
- Raw-text search mode (any line, not just drops) still available as a fallback

### v2.5
- Item builder overhaul: category-based spell dropdowns (Basic, Shields/Buffs, Class Specific, Other1, Other2) with paired tier dropdowns, replacing free-text spell entry
- Per-spell tier restrictions (e.g. Agility/Bless capped below tier iii)
- Wanted Spells shown as removable, wrapping chips instead of a plain vertical list
- Realm filter ("Only Found In") for Evil, Chaos, Good, Kaid, Crafted, Glory Bea, and Event
- Build cap of 1 Crafted-realm item per build
- Duplicate-tier requests for the same spell no longer consume two equipment slots
- "Generate multiple build options" for comparing alternate builds, plus an Alt Options results column
- Build Config split into 1 Claw / 2 Claw
- Weapon Style "Direct (Stat)" renamed to "Parry Staff"
- Results table column widths tightened; window resized for the larger Build tab

### v2.4
- Added "Use Community List" button for instant database loading
- Bundled Olmran community equipment list with the release

### v2.3
- Added Min/Max/Specific level filtering with smart field disabling
- Spell abbreviation support (superseded in v2.5 by the category dropdowns)

### v2.2
- Redesigned weapon system: Melee vs Direct (Caster) vs Direct (Stat)
- Build Config checkboxes: Weapon, Shield, Two-Handed, Claw
- Dual-Wield sub-option
- Direct weapon parsing from delve text
- Smart parse options, "All:" armor row, per-file chat sheet exports

### v2.1
- UI consolidation (merged Files + Parse tabs), global Chat/Combat/Loot checkboxes, enhanced Build Creator

### v2.0
- Initial portable release with auto-installer

## Support
For issues, questions, or feature requests, contact the developer.

---
**Created by:** Claude & Developer Team
