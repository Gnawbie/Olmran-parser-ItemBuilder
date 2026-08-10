# Gaming Log Parser v5.2.0

**[⬇ Download the latest release](https://gnawbie.github.io/Olmran-parser-ItemBuilder/)**

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
- **Counters** sub-tab — XP Counter, Damage Counter, PvP Damage, and **Item Counter** (drop rate per mob/item, broken down by Zone and by Daily/Weekly/Monthly/Yearly period, with a Character filter and a persistent per-character "Consistent" tally — see Version History below for the full feature list)

### Settings (Parse sub-tab)
A mini-notebook holding Fields and Export.

**Fields**
- Customize which fields get extracted for Chat, Combat, and Loot
- Add, edit, remove, and reorder fields per data type
- "Reset to Defaults" restores the built-in field set

**Export**
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

**Required Items**
Want a specific piece of gear included no matter what? Type its name under **Require Item:** and click **Add to Build**. It's looked up in the loaded master database (case-insensitive exact match first); if nothing matches exactly, it falls back to a spelling-tolerant fuzzy search and asks you to confirm the closest match (or pick from a short list if there are several similar names). Once added, it shows up as a removable chip under **Required Items** and is forced into its slot on the next search — the rest of the build is then calculated around it, with the rest of the slots filled normally.

**Level Filtering**
- **Min Level** / **Max Level** — restrict to a level range
- **Specific Level** — restrict to an exact level
- Fields grey out automatically depending on which option you're using (Specific vs. Min/Max are mutually exclusive)
- Leave all three blank for no level restriction

**If no match at Specific Level** (only meaningful, and only enabled, once Specific Level has a value) — controls what Find Optimal Build does for a slot when nothing carries the wanted spell at exactly that level (and, for a spell requested at an explicit tier, at exactly that tier too):
- **Go down a tier** — keep the level fixed at Specific Level, but accept a lower tier of the wanted spell (iii → ii → i) if the exact tier isn't available there
- **Go down in level** — keep the exact requested tier, but search downward from Specific Level for the highest level that has it (never picks an arbitrary lower level - always the closest available)
- **Both** — relax level and tier together, still preferring the highest level / closest tier combination available
- **Don't populate slot** (default) — no fallback; if nothing matches exactly, the slot is left empty

**Armor Constraints**
- The **All:** row sets an armor type for every slot at once
- Override individual slots (Head, Cloak, Body, Hands, Legs, Feet) with Cloth / Leather / Studded / Plate
- Per-slot **Defense** (range) and **Sigil** dropdowns — soft preferences: an item matching them is favored, but a slot is never left empty just because nothing matches
- **Set as Default** / **Clear Default** / **Clear All** manage your saved preferences

**Weapon Constraints**
- **Weapon Types/Combo's** — check whichever your build actually uses: Dual-Wield 1h, 1h/Shield, 2h/Shield, Fired 1h/Shield, Two-Handed, 1 Claw, 2 Claw. Most have their own Style (Melee/Direct/Parry Staff/Fired, where applicable) and Damage Type (Slashing/Thrusting/Crushing) dropdowns
- Leave everything here unchecked and Weapon/Shield aren't populated at all - check a combo to have that slot (or slots) filled, even with a spell-less item if nothing wanted matches, rather than sitting empty. 1 Claw/2 Claw replace Weapon and Shield entirely (claws are their own one-handed weapon) rather than filling alongside them
- **Melee Weapon Constraints** — soft-preference Damage/Timer/Fumble/Accuracy/Sigil dropdowns, each with an optional Priority checkbox (capped at 3); apply to every weapon style (Melee, Direct, Parry Staff, Fired) and to Claw slots too
- **Shield Constraints** — Defense and Sigil dropdowns, working exactly like Armor Constraints' per-slot versions, plus Cloth/Leather/Studded/Plate checkboxes (one or more can be checked) - a hard filter, same as Armor Constraints' own per-slot armor type checkboxes. Shields are technically armor, but are built alongside a weapon, which is why they live here instead of in Armor Constraints

**Only Found In (Realm filter)**
Next to the spell dropdowns, check any combination of Evil, Chaos, Good, Glory Bea, Crafted, Event, or Kaid to restrict results to items found in those realms. Leave everything unchecked to search all realms except Crafted (see below). Kaid has its own column: **Kaid All** matches any Kaid sub-realm, or check one or more specific colors (Kaid White/Green/Red/Purple) instead - Kaid All and the colors are mutually exclusive.

**Crafted items never appear unless the Crafted checkbox is checked** - even if another checked realm would otherwise match part of the item's Realm text (e.g. a "Crafted - Evil" item with only "Evil" checked), and even when no realm boxes are checked at all. Once Crafted is checked, a build can still include **at most one Crafted-realm item** - enforced automatically during the search, not just a filter you have to remember.

**Search**
- **🎯 Find Optimal Build** — exact search across every slot at once for the combination that covers as many wanted spells as possible (and best satisfies Priority Tier targets) under your constraints, rather than committing slot-by-slot
- **📋 Show All Matches** — lists every item that matches your filters, without narrowing to one per slot
- **🎲 Generate multiple build options** (checkbox) — when checked, "Find Optimal Build" also generates up to 10 alternate full builds by swapping in equally-good "tied" items slot by slot, so you can compare a few options instead of just one

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
- **Event** realm items work like any other realm filter, but there isn't yet dedicated Event-specific logic beyond the Realm column match
- Some improved-tier (`iii`) **Protect** spells don't show up correctly in Find Optimal Build results yet — root cause not yet diagnosed
- **Direct** weapon spells are generally not captured in the bundled equipment list (Parry Staff spells are captured correctly) — this is a gap in the source data, not the search logic
- `OlmranItemBuilder.exe` isn't code-signed, so Windows SmartScreen shows an "Unrecognized app" warning on first run — click "More info" then "Run anyway" (see the [FAQ](https://github.com/Gnawbie/Olmran-parser-ItemBuilder/wiki/FAQ))
- The .exe is a PyInstaller "onefile" build, so it unpacks itself to a temp folder on every launch — expect a few seconds of startup delay each time, not just the first

## Troubleshooting

**"Windows protected your PC" / SmartScreen warning on launch** — expected for now, `OlmranItemBuilder.exe` isn't code-signed yet. Click "More info" then "Run anyway". See the [wiki FAQ](https://github.com/Gnawbie/Olmran-parser-ItemBuilder/wiki/FAQ) for details.

**The .exe takes a few seconds to open every time** — expected for the single-file `OlmranItemBuilder.exe` download. It's a PyInstaller "onefile" build that unpacks itself to a temp folder on every launch, not just the first. If this bothers you, or you've ever seen a "Failed to load Python DLL" error, use the **Folder** download instead ([download page](https://gnawbie.github.io/Olmran-parser-ItemBuilder/)) — unzip it once and it opens instantly every time after, since nothing needs to be unpacked on launch. Both are the same program; "Check for Update" correctly updates whichever one you're running.

**Program won't start otherwise** — see INSTALL_INSTRUCTIONS.txt.

**Files not loading** — check the file extension (.log, .txt, .xlsx), make sure it isn't open in another program, and try copying it somewhere else if the path has unusual characters.

**Parse results empty** — verify the log actually contains delve output ("You examine X closely") and that CHAT vs ACTION auto-detection picked the right type (check the Parse tab file list).

**Community List not found** — as of v5.0 the community list ships bundled inside `OlmranItemBuilder.exe` and loads automatically; if you're instead running from source (`gaming_log_parser.py`), make sure `Olmran_Community_Eq_and_Stats_List.xlsx` is in the same folder, or use "Browse..." to point at it manually.

## Version History

### v6.5.3 (Current)
- **Bank Build**: fixed the Gear Tag dropdown sometimes closing itself instantly, with nothing picked, right after opening - posting its list momentarily handed focus to the list's own popup, which was misread as the user clicking away. This was also the real cause behind the previous fix not fully sticking - the earlier "resets the list's sort order" report was this same instant-close, not an actual resort.

### v6.5.2
- **Bank Build**: fixed picking a Gear Tag resetting whatever column the list was currently sorted by back to its stored order - it now updates just that row's Tag cell in place instead of rebuilding the whole table, so the on-screen sort stays put (this also stopped a stray auto-close of the Tag dropdown, with no value actually picked, from doing the same reset).

### v6.5.1
- **Bank Build**: added a leading **Bank** icon column (📦, or a bold "L" for anything also available via a Locker) to Main, every character tab, every Locker tab, and the Trade tab - same convention already used on Results/Manual.
- **Bank Build**: Share and Trade flags now persist independently of each character's item list - Clear Saved List, Delete Character, or re-importing after a data wipe no longer loses them, so a big list of items never needs to be re-flagged from scratch.

### v6.5.0
- **Bank Build**: added an **Enchant** column (Character Items, Lockers, and Main alike, between Drop and Item) - shows whatever quality/material prefix (e.g. "Bright", "Shining Mithril") was stripped off the item's name when it was imported, so that text isn't lost even though the cleaned name remains the actual matching key everywhere else (Gear Tag, Share, Trade, and every search/build pool)
- **Build → Trade**: added a clickable "Olmran Trading Post" link next to the Trade heading, opening https://www.olmrantradingpost.com/ in your browser
- **Parse → PvP**: added a Kills/Deaths count (e.g. "10/4") next to the Search button for Your Deaths
- Fixed XP Gained and XP/Hour (and any other comma-formatted number column) sorting alphabetically instead of numerically when clicking the column header
- Fixed the app not scaling correctly on multi-monitor setups where displays use different Windows display-scaling percentages
- Fixed a wide-reaching layout bug where several tabs (Results, Saved Builds, Area Items, Bank Build, Manual, Trade, and every Counters sub-tab) forced extra scrolling to reach content that was already short enough to fit on screen, because the tab notebook sized itself to its single largest tab instead of whichever tab was actually showing
- Fixed the window's bottom edge landing behind the Windows taskbar after maximizing and then restoring it - a side effect of the multi-monitor scaling fix above, since Windows stops auto-correcting window placement once an app declares itself DPI-aware

### v6.4.5
- **Bank Build**: added a **Trade** column (Character Items and Lockers alike) - a checkbox-style "No"/"Yes" flag, defaulting to No, marking one item as available to trade. Purely bookkeeping - it doesn't affect Bank Build's own search pooling at all, unlike the Share column
- **Bank Build**: added a new **Trade** tab (next to Manual) - a flat, read-only chart of every item any character or Locker has flagged Trade = Yes, with a Character column showing who owns each one, so there's one place to see everything currently up for trade instead of checking every tab separately

### v6.4.4
- **Area Items**: added a navigation tree on the left, grouping every Area into its base zone with a child per sub-zone (e.g. "City of Zhak-Tor" → "Temple"/"Tower") - click a sub-zone to browse just that part, or the base zone itself if it's also its own standalone Area. Stays in sync with the search box either direction - picking one updates the other
- **Area Items**: added Kaid/Chaos/Good/Evil/GB (Glory Bea) checkboxes next to the search box, narrowing which areas the navigation tree shows (an area appears if at least one of its own items matches a checked realm) - the results table itself is unaffected, still scoped to whichever one area is picked

### v6.4.3
- **Bank Build**: each non-Locker character's own Character Items tab has a new **Share** column (a checkbox, after Gear Tag) - checking it flags that one item as shareable, folding it into every other character's Bank Build search automatically, regardless of that character's own "Search all characters" setting. This is the per-item equivalent of what a whole Locker already does for its entire list - useful for sharing just a few spare items without turning the character itself into overflow storage. Persists across restarts; not shown on Locker tabs, since a Locker already shares everything

### v6.4.2
- **Item Counter**: the Overall/Daily/Weekly/Monthly/Yearly Zone → Mob → Item trees are now sortable by clicking any column header, same as every other table in the app - this was an oversight in v6.4.0, they'd never actually been wired up. Sorting one level (e.g. Zone) also re-sorts everything nested under it recursively
- Fixed a build-up of stale widgets behind the scenes every time a Counter's results tab got replaced (re-clicking XP/Damage/PvP/Item Counter, or the Item Counter auto-restoring on startup) - the old tab was being hidden but never actually destroyed, which could show up visually as extra stacked border lines the more times it happened
- Removed the box border that used to sit around the Counters sub-tab's own notebook (XP/Damage/PvP/Item Counter) and each individual Counter's own results notebook, since it mostly just boxed in empty space - kept the one line under the tab strip
- **Manual tab**: several tables (its own results list, Bank Build's Import/Character Items/Lockers tables, and others under Build) now stretch to fill the rest of the window instead of stopping short and leaving blank space below them whenever the window is taller than the tab actually needs
- **Manual tab**: added a separator line above the "N matching items found" count, matching the line style used elsewhere

### v6.4.1
- **Item Counter**: the "Consistent" chart itself now persists across closing and reopening the program - previously the running tally saved correctly, but nothing re-displayed it until you reloaded logs and clicked the button again; now it rebuilds automatically on startup if Consistent is checked and a tally is on disk
- **Item Counter**: the Character dropdown's selection, and every character it lists, now persist across restarts too - previously it reset to "All Characters" and forgot any character not tied to a currently-loaded file
- **Item Counter**: added a date range (earliest–latest log date and the span in days) and a log count next to "Total Items Dropped", e.g. `1/15/2026 to 2/20/2026 (36 days)  •  2 logs loaded`
- **Bank Build**: a Locker's Kaid-realm items now fold into every other character's search pool like any other gear, instead of being silently excluded - Kaid became tradeable in-game, so the old "Lockers only share non-Kaid gear" restriction (from when Kaid items couldn't actually be handed to another character) no longer applies. This is a live in-game change still being tested and could be reverted later

### v6.4.0
- Added an **Item Counter** to the Counters sub-tab - drop rate of every item, per mob, across the loaded log(s), broken down by Zone and by Daily/Weekly/Monthly/Yearly period
  - **Overall** tab is a nested tree: **Zone → Mob → Item**, each level showing its own Drops/Kills/Drop Rate % (Zone and Mob rows sum up whatever's nested under them). Daily/Weekly/Monthly/Yearly are each set up the exact same way, with Period as one more level on top - a **Log Files** tab lists every loaded file's own extracted date and kill/drop counts
  - **Drop Rate %** = `(drops / kills * 100) / possible items` - possible items being how many distinct items the master/community database lists for that mob (1 if the mob isn't in the database), so an item competing against a bigger loot table reads as rarer at the same drops/kills ratio. The exact formula is spelled out as a note next to the tab's Save/Delete button
  - Right-click a **Zone** for **Open into new tab**, a **Mob** for **Open monster in new tab** (every zone it was killed in), or an **Item** for **Open item in new tab** (accumulates every item added this way into one shared "Item Comparison" tab, for building a side-by-side watchlist) - every level also gets **Add to (name)**, appending that selection into any already-open tab instead of opening a new one, and **Delete** (removes just that row from wherever it was clicked). Mini-tabs get a **✕** to close them (click it specifically - clicking elsewhere on the tab just selects it), and a blank line automatically separates rows from different zones in any table that spans more than one
  - **Items** / **Junk Loot** checkboxes - Items means "found in the loaded master/community database", Junk Loot is everything else (on/off by default respectively)
  - **Character** dropdown - defaults to "All Characters"; extracted from each log filename's own character-name segment (e.g. `ACTION-PDF-GNAWBIE-<timestamp>.log` → "GNAWBIE"), files that don't fit that pattern count as "Unknown"
  - **Consistent** checkbox (checked by default, confirms before letting you uncheck it) - keeps a running tally per character across every past click instead of starting over from just what's currently loaded; any log file already counted for that character is skipped, not counted again, so reloading the same logs across sessions never double-counts. "All Characters" shows every character's own tally combined
  - Daily/Weekly/Monthly/Yearly bucket by each log file's actual **OS creation date**, not anything parsed from the filename
  - A bracket like `[Sin      60]` (a class/summon status display, name padded with 3+ spaces before a trailing level number) is recognized as not a real zone - it's excluded from every Zone chart, though its kills/drops still count toward Overall totals
  - Same Save/Delete-a-named-report pattern as the other Counters, surviving a restart

### v6.3.0
- Export (Settings sub-tab) gained a **File Format** dropdown - alongside the existing Excel Workbook (.xlsx), you can now export to **XLS** (Excel 97-2003), **ODS** (OpenDocument Spreadsheet), **CSV**, or **TSV**, each with a one-line explanation of what it is shown just below the dropdown. XLSX/XLS/ODS keep every checked type as a separate sheet in one file; CSV/TSV (which can't hold more than one table per file) write one file per checked type instead - picking either of those two now shows a folder picker instead of a misleading single-file "Save As" dialog
- Added a **Sheets to Export** checklist - one entry per Chat file/Combat/Loot/XP Counter/Damage Counter/PvP Dealt/PvP Taken that currently has data, showing its row count, with Select All/Select None/Refresh buttons. This is what actually decides what gets exported, independent of Parse Options' own checkboxes (which only control what gets computed). It refreshes automatically after Run Parse, and now also after using any of the Counters sub-tab's own XP Counter/Damage Counter/PvP Damage Counter buttons directly - previously that data never made it into Export at all unless you separately ran Parse with the matching checkbox checked
- Export now has **two entirely separate destinations**, each with its own Output filename and Export button: **Export Chat / Combat / Loot** (carries the Master Database section and the separate-files option, unchanged) and **Export Counters** (a plain snapshot file for the 4 Counters, with no master-database/accumulation behavior - just that export's own data, easy to hand to someone without needing this program to view it)

### v6.2.0
- The Fields tab (Settings sub-tab) gained 4 new customizable field types - **XP Counter**, **Damage Counter**, **PvP Dealt**, and **PvP Taken** - alongside Chat/Combat/Loot, each with its own Col #/Header Label/Data Source rows (add/edit/remove/reorder/reset all work the same way)
- Parse Options (Files & Search) gained matching checkboxes for these 4 Counters - checking one and clicking **Run Parse** computes it (same engine as the Counters sub-tab's own buttons) and makes it available to **Export to Excel**, which now writes a sheet per checked Counter using whatever columns were customized in Fields. New Snapshot buttons and Parse Results tiles were added for all 4 as well
- The Counters sub-tab's **Load Log Files** table is now drag-resizable from its bottom edge - shrink it to give more room to the XP/Damage/PvP Damage notebook below, from 8 lines (its original size, now the maximum) down to a minimum of 3, starting at 4 (half) by default. The last size you drag it to persists across restarts. Files & Search's own copy of this table is unaffected and stays fixed

### v6.1.3
- Saved Items now default an untagged **Event**-realm item's Gear Tag to **Good Gear** instead of Blank, everywhere the Gear Tag column appears (Main, every character, every Locker) - also feeds into the Good Gear only/Invasion Gear only search filters. An explicit tag (including explicitly choosing Blank) always overrides this default

### v6.1.2
- Added the same **Bank** (📦) and **Locker** icon columns from the Results tab to the Manual tab's results table - a plain 📦 marks any item you've already saved anywhere, or a bold "L" plus the first 4 letters of the Locker's name for a Locker-sourced item, purely informational since Manual has no build/scoring concept of its own

### v6.1.1
- Moved the **Fields** and **Export** tabs from the top-level tab strip into a new **Settings** sub-tab of Parse (alongside Files & Search and Counters) - Settings is a small mini-notebook holding both. Nothing about either tab's own content changed, just where they live

### v6.1.0
- Added a **Damage Counter** next to the XP Counter (both now grouped under a single "Counters" sub-tab, sharing one "Load Log Files" section) - totals damage dealt to every mob across the loaded log(s), with a **Per-Mob Summary** (Mob/Damage/Hits/Avg per Hit) and an **Every Hit** detail table (Timestamp/Mob/Damage/Weapon/File) so individual hits on a mob can be isolated even when other mobs were hit in between. Both tables sort by clicking any column header, and both support saving named reports that persist across restarts (with Delete)
- Every Hit's **Weapon** column shows the melee weapon used, the actual spell cast, or - for a "Direct" weapon's own innate elemental proc (no spell-cast line at all) - falls back to the weapon that triggered it, so that damage is never left unattributed
- Melee-hit detection now recognizes any attack verb ("smash", "attack", "maul", "backstab", "thrust", "slash", etc. - varies by class/shapeshift form) instead of a fixed list, and mob names are consistently de-articled ("a"/"an"/"the" stripped) whether the hit came from melee or a spell/AoE line, so the same mob no longer fragments into two separate summary rows
- Added a **PvP Damage Counter** as a third tab in Counters - splits damage dealt to and taken from other players (distinguished from mob combat by a player's name being a single bare word with no "a/an/the" article) into its own Damage Dealt / Damage Taken tables, each with the same Per-Player Summary + Every Hit structure, sorting, saving, and Open Log support as the mob Damage Counter. A hit landing on a player is now excluded from the regular Damage Counter's mob totals entirely
- XP Counter gained its own **Every Gain** detail tab (Timestamp/Area/XP Gained/File) alongside the existing Per-Area Summary, with the same click-to-sort headers
- Added **"Open Log at This Line"** (right-click any row) to the Every Hit and Every Gain tables, the same feature already used by Search Logs/Drop Search/PvP results - jumps straight to and highlights that exact line in the source log

### v6.0.0
- Added an **Edibles** mini-tab to Build Constraints - covers consumable items (Type='edible' in the master database), which a build can use any number of at once. Has its own Available Spells whitelist (built from a data-driven dropdown reflecting whatever edible spells/tiers actually exist in the loaded database), an "Add edibles to the build" checkbox (off by default), and a Max Edibles Used slider (1-13). Used edibles show up as extra "Stomach" rows in the Results tab, and count toward real Wanted Spell coverage just like armor/weapons
- Fixed Wanted Sigils awarding its scoring bonus to every slot carrying that sigil independently, which could stack the same sigil across several slots when only one was ever intended - now only the first use of a given sigil type counts, with reuse allowed only when it's the sole way to cover a genuine Wanted Spell
- Fixed a related bug where an item picked purely for a Wanted Sigil could still redundantly restate an already-covered Wanted Spell at a different tier (e.g. wisdom.ii showing up in a build where wisdom.iii was already covered elsewhere, including by an Edible) - no wanted spell can stack across two places in a build now, regardless of sigil
- Empty slots ("No suitable item found") now always show in the Results tab, even when every wanted spell happens to be covered elsewhere - previously they were silently hidden in that case, which no longer reflects reality now that the sigil/redundancy fixes above can legitimately leave a checked slot empty
- Manual tab's "Add to Results Tab" now inserts into whichever Results view is currently showing (Best Per Slot or All Matches), not just its own separate list - it fills a matching empty slot first if one's available, otherwise groups as an extra row at that slot's normal position. Manually-added gear now survives Remove/Rebuild/Search Missing Slots elsewhere in the build (previously it could be wiped out entirely by an unrelated action), and can be removed via right-click same as before
- Saved Constraints now also captures the Edibles tab's settings and specific Event checkboxes (previously only the general "Event" realm checkbox was saved, silently losing specific event selections on load)
- Moved the "Download Page"/"Check for Update" buttons to sit in the same row as the top-level tabs instead of their own separate row above them, and trimmed their background so it doesn't cover the tab strip's border

### v5.6.4
- Fixed Manual tab's Armor Type filters (material checkboxes, per-slot "only this piece" checkboxes, per-slot Defense, per-slot Sigil) not actually excluding weapons/jewels/shields from the results list - they only ever narrowed which armor items showed, silently doing nothing for non-armor slots unless the separate top-row "Armor" checkbox was also checked. Any of those Armor Type controls being actively set now implicitly restricts to armor items too

### v5.6.3
- Saved Builds is now a Notebook - each saved build is its own renamable, removable sub-tab instead of a stacked, scrollable row, so having many saves just adds more tabs rather than a taller page
- Added a "Load" button to each Saved Builds sub-tab - copies that build's rows straight into the Results tab and jumps there

### v5.6.2
- Class Items is now a typeable field with a live-filtering suggestion popup (same autocomplete style as the Area Items tab), instead of a plain dropdown you could only scroll through
- Moved Min/Max/Specific Level, the "If no match at Specific Level" fallback options, and the Find Optimal Build/Show All Matches/Generate multiple build options buttons into the Wanted Spells column (right below Clear All), filling space that used to sit empty there. These now only show while Basic Constraints is the selected sub-tab, rather than staying visible on Bank Build/Manual too

### v5.6.1
- Parse tab's PvP section: "Your Kills" and "Participated" each now show a Total RP (realm points) field and a count field (Kills/Participations) next to their Search button, filled in automatically whenever that search runs

### v5.6.0
- Consolidated Basic/Armor/Weapon Constraints - Armor Constraints and Weapon Constraints no longer have their own top-level sub-tabs; both now live as compact mini-tabs directly inside Basic Constraints' "Build Constraints" area, alongside a "Basic" tab holding the spell dropdowns and Priority box. Every control is the same underlying setting either way - nothing was duplicated or reset
- Only Found In (with its Events tab) moved into that same "Basic" mini-tab, next to the Priority box, instead of sitting in the row beside Build Constraints
- Merged "Priority Spell" and "Priority Tier" into a single "Priority" box - the tier dropdown now includes "(any)": picking it behaves like the old Priority Spell (searched for regardless of tier), picking a specific tier behaves like the old Priority Tier (targets that tier specifically)
- Shrunk the Wanted Spells chip box, then resized it to match Required Items for a more even look

### v5.5.1
- Manual tab: Wanted Spell chips are now optional, not required for the results list to appear at all - with none added, every item matching the other active constraints (Only Found In, Armor Type, Weapon, Melee/Shield, Level, Armor/Weapons/Jewel) shows up; adding one or more chips narrows that down further
- Manual tab's Armor Type sub-tab: added a checkbox before each of Head/Cloak/Body/Hands/Legs/Feet - checking one or more restricts the list to just those specific armor pieces, on top of whatever the broader Armor/Weapons/Jewel checkboxes already allow
- Added a "Jewel" sub-tab alongside Armor Type/Weapon/Melee-Shield in Manual's Gear Constraints area, with a Sigil dropdown for jewel items
- Fixed the Weapon sub-tab's "Two-Handed" Style dropdown (Melee/Direct/Parry/Fired) not refreshing Manual's results list when changed - it was only updating its own Damage Type field's enabled state

### v5.5.0
- Added an "Events" tab next to "Only Found In" in Basic Constraints - one checkbox per specific event (e.g. "Halloween 2020", "Christmas 2023"), automatically populated from whatever's in the loaded database. Works additively alongside the existing broad "Event" checkbox: checking a specific event lets its items through even when "Event" itself is unchecked
- Added a new "Manual" tab to the Build section (after Bank Build) - a standalone spell/tier database lookup, separate from the actual build search. Pick a spell category + exact tier and add it to a small list; every matching item in the loaded database shows up immediately below, right-click any result to add it into the Results tab (multiple items can go into the same slot here, e.g. 3 different body pieces, since it's just a flat list you're curating by hand, not a real build)
  - Has its own independent copy of Only Found In (Realms + Events), Armor Type (a simplified Cloth/Leather/Studded/Plate checklist plus per-slot Defense/Sigil), Weapon (Two-Handed/Dual-Wield/1h-Shield/2h-Shield/Claws/Fired), and Melee/Shield Constraints - none of it touches or is affected by the real Basic/Armor/Weapon Constraints tabs
  - Armor/Weapons/Jewel checkboxes restrict the list to just one or two of those slot categories; checking Weapons also drops the requirement to have any spell picked at all (most real weapons carry no spell) and adds Weight/Fumble/Damage/Timer/Accuracy columns to the results table
  - Has its own independent Level range filter, separate from Basic Constraints' Level fields
- Added click-to-sort column headers to Manual's results list, the Area Items tab, and every read-only search/snapshot dialog (log search, drop search, action/chat snapshot viewers) - click a header to sort by that column, click again to reverse. Deliberately left off the Results tab (row order there matters for Remove/Rebuild) and the Fields tab/Parse tab's file list (row order there is meaningful, curated data)
- Widened the app's default startup window size

### v5.4.38
- Fixed the v5.4.37 "reopen it yourself" message sometimes stopping the update from actually installing at all - it was shown right as the app was closing, but the swap step in the background only waits about a minute for the app to release its own file, and that clock started ticking the moment the message appeared rather than when it was dismissed. Taking longer than that to click OK could let the swap give up entirely before the app ever actually closed. The message now shows (and needs to be dismissed) before the swap starts, not after

### v5.4.37
- Auto-updater no longer tries to automatically reopen the app after installing an update. Every attempt at fixing an intermittent "Failed to load Python DLL" error on the automatic reopen step (adaptive delays, environment cleanup, better retry detection) still hit the same failure on every retry, while launching the exact same exe independently (a plain double-click) reliably worked - strong evidence the problem was specific to reopening it automatically from within the updater's own process chain. The update still installs the same as before; you'll now see a message when it's done letting you know to open the app again yourself

### v5.4.36
- Auto-updater: fixed the launch-retry check (developer-diagnostics builds only) always declaring success after the very first attempt, even when the freshly-relaunched exe had actually failed and was sitting on its own fatal error dialog - a process wedged on an error dialog still has a matching image name in the task list, so a bare "is a process with this name running" check was reporting that as success and never actually getting a second attempt. Now also checks the process's own window title for the word "Error" and, if found, closes it and retries instead of stopping

### v5.4.35
- Auto-updater: the relaunched exe now starts from a sanitized environment instead of inheriting the old (about-to-exit) process's own - PyInstaller's bootloader sets variables like `TCL_LIBRARY`/`TK_LIBRARY` pointing at its own extraction folder, and those were being inherited down through the update batch script into the freshly-relaunched exe, which could then try to load Tcl/Tk or the Python DLL from the OLD process's (now-deleted) extraction folder instead of its own. A plain manual double-click was never affected since Explorer's own environment was never touched by any PyInstaller process to begin with

### v5.4.34
- Auto-updater: removed the "clear stale extraction folders" step added in v5.4.31 (developer-diagnostics builds only) - direct evidence has since shown every "Failed to load Python DLL" report points at a genuinely fresh extraction folder, not a reused stale one, disproving the theory that step was based on. Worse, it could take 20+ seconds right before the critical relaunch moment on a machine with many old folders to clean up, which is exactly the kind of disk activity burst that could interfere with the new extraction about to happen

### v5.4.33
- Fixed "Rebuild (Saved Items First)" sometimes filling two different slots with two different items for the same stat (e.g. a constitution item in both Cloak and Body) - its full-database fallback pass filled each empty slot independently against the same "still needed" list without ever narrowing it down, so two empty slots needing the same spell could both grab their own covering item instead of just one
- Fixed that same Rebuild's first pass (Saved Items only) sometimes wrongly rejecting an owned item over its exact spell tier or its drop realm, pushing it into the full-database fallback for no reason - it's now consistent with Bank Build's own Hard Search: any tier is accepted and realm no longer matters for items you already own
- Fixed the full-database fallback used by Rebuild and "Search Full Database for This Slot" ignoring which spell tier was actually requested, picking whichever candidate had the highest level regardless of tier - it now prefers a match for the tier you asked for first

### v5.4.32
- Added a second download option: a "Folder" distribution (`OlmranItemBuilder_Folder.zip`) alongside the existing single-file `OlmranItemBuilder.exe`. Unzip it once and run the exe from inside - no per-launch self-extraction step, so it opens instantly every time, unlike the onefile build which unpacks itself to a fresh temp folder on every launch. Same program either way; "Check for Update" correctly detects and updates whichever one you're running, including swapping in a fresh copy of its files and relaunching automatically

### v5.4.31
- Fixed Bank Build's Hard Search silently excluding an owned Saved Item just because its drop realm wasn't checked under "Only Found In" - Hard Search only ever searches items you already have, so an item's original drop realm no longer matters there (this fixed a case where the only Saved Item carrying a required Wanted Sigil never got considered, and that sigil requirement went unmet with no indication why). Every other search mode still respects "Only Found In" as before
- Auto-updater: fixed the launch-retry check (added in v5.4.30, developer-diagnostics only) always reporting the freshly-updated exe as "not running" and retrying regardless of whether it actually launched, on machines where a Unix-style `find` command (e.g. from Git for Windows) shadows Windows' own `find` on PATH - now calls the real `find.exe` directly instead of trusting PATH. Also clears any stale leftover extraction folder before each relaunch attempt, and waits progressively longer between retries (4/6/8/10/12 seconds instead of a flat 4), since a freshly-extracted, unsigned exe can need more time for antivirus scanning to clear on some machines

### v5.4.30
- Fixed Results tab's right-click "Rebuild" (Saved Items First / Full Database, Prefer Owned) not respecting the originating character's Bank Build settings - Search all characters, Good Gear only/Invasion Gear only now carry over into a later Rebuild instead of silently falling back to a generic, every-character pool
- Fixed "Rebuild (Saved Items First)" not respecting Wanted Sigils circle requirements when every Wanted Spell was already covered by Saved Items alone - a required sigil the bank box couldn't provide can now still be found in the full database on Rebuild
- Auto-updater: the swap step now automatically retries relaunching if the freshly-updated exe doesn't seem to stick around a few seconds after launch, instead of leaving the "Failed to load Python DLL" error on screen (this retry only shows its own diagnostics on this developer's machine - everyone else's update stays silent and automatic either way)

### v5.4.29
- Wanted Sigils (Build → Armor Constraints): the two circles now stay their own color (red/blue) whether hollow or filled in, instead of turning gray when off
- Internal: added a developer-only diagnostic mode for the auto-updater's swap step, to help track down a recurring "Failed to load Python DLL" report - invisible to everyone else, who keep the normal silent, auto-closing update

### v5.4.28
- Added two circles to each Wanted Sigils chip (Build → Armor Constraints) - red requires one piece of gear carrying that sigil AND any Wanted Spell; blue requires that sigil AND a matching Protect spell (that element's own, or the generic Elemental Protect) instead, independent of the Wanted Spells list. Mutually exclusive per sigil, enforced as hard requirements even if it leaves other slots empty

### v5.4.27
- Fixed Locker Groups drag-and-drop not working - dragging a tab required clicking a tiny sliver right at the tab's edge to register at all, since the check relied on a Tk quirk where the actual clickable label text (where anyone would naturally click) reports differently than the tab's edge does. Dragging a Locker tab onto a group tab (or back out) now works from anywhere on the tab, not just that edge

### v5.4.26
- Fixed the auto-updater's "Failed to load Python DLL" error still occurring after a successful update on some machines - the v5.4.23 fix added a fixed ~2 second pause before relaunching to let antivirus finish scanning the freshly-placed exe, but scan time varies by machine and 2 seconds wasn't always enough. Replaced the fixed delay with an adaptive check that actively tests whether the file is still locked (the same technique already used for waiting on the old process to exit) and only launches once it's genuinely clear, waiting exactly as long as needed instead of guessing

### v5.4.25
- Added Locker Groups - Build → Bank Build → Lockers now has a leftmost "+" tab (no page of its own) that prompts for a name and creates a new group tab to organize Locker tabs into. Drag a Locker tab onto a group tab to move it in, or drag it back out onto the main tab strip to ungroup it. Group membership persists across restarts

### v5.4.24
- Fixed a bug where scrolling the mouse wheel over the Saved Builds tab's list would permanently break scrolling the main window for the rest of the session (the first time the mouse ever crossed it) - it used its own separate app-wide wheel binding that replaced, then fully removed, the one everything else relies on. Now uses a shared registry instead, so the mouse wheel scrolls the main window everywhere by default, defers to whichever list/table/canvas is actually under the cursor when it has its own scrolling, and keeps working correctly afterward no matter where you've scrolled

### v5.4.23
- Fixed the auto-updater's "Failed to load Python DLL...LoadLibrary: The specified module could not be found" error on the very first launch after an update - confirmed by a second, independent report that a plain manual relaunch immediately fixed it, meaning the freshly-swapped exe was never actually corrupted, just briefly locked by antivirus finishing its on-write scan. The swap script now waits a couple seconds after moving the new exe into place before launching it, giving that scan time to finish first

### v5.4.22
- Closed a real gap in the v5.4.20 auto-updater fix: the download size check only ran once, in memory, before writing the new exe to disk - if antivirus scanned and tampered with the freshly-downloaded, unsigned exe sitting in Temp afterward (which can happen, since the swap can only run once the old process fully exits, not necessarily instantly), that corruption slipped through untouched. The swap script now re-verifies the file's on-disk size immediately before every move attempt, retrying or safely relaunching the current exe untouched if it doesn't match, instead of ever installing a tampered file

### v5.4.21
- Fixed Build → Basic Constraints → Wanted Spells: Class Specific spells (Reverb, Aura, Backstab, Bash, Berzerk, etc.) were capped at tier ii - their tier dropdown now goes up to iii like other spell categories

### v5.4.20
- Fixed the auto-updater being able to corrupt the installed exe on a bad download - the sanity check only rejected downloads under 1MB, so a truncated/interrupted download (e.g. 8MB of the real ~15.7MB) could still sail through and get written over the working install, producing an exe that fails to launch at all ("Failed to load Python DLL"). Now validates the exact byte count against GitHub's own reported asset size, plus checks the file starts with a valid EXE header, before ever touching the installed exe - a bad download is now reported as a failure instead of corrupting anything

### v5.4.19
- Added an "XP Counter" - Parse tab now has its own inner sub-tabs, "Files & Search" (everything that used to be there) and a separate always-visible "XP Counter" sub-tab. Computes XP/hour from the loaded log(s) (parsing "You gain BASE (+BONUS) experience points." lines) plus a per-area breakdown by time actually spent in each area (tracked via the game's room-title lines). Results show in an ephemeral working tab that's replaced on every click; a Save button turns a result into its own persistent, named tab (with a Delete button), surviving a restart. The "Load Log Files" file list appears on both sub-tabs, kept in sync, so files can be loaded without switching away from XP Counter

### v5.4.18
- Fixed the auto-updater's "Download & Update" not actually installing the new version - it would download successfully, close the app, then never reopen it. Root cause: the batch script that swaps the exe used `tasklist | find` to wait for the app to fully exit, but that check silently fails (an encoding quirk of running with a hidden console) and falsely reports the process already gone. It also used `timeout` for retry delays, which doesn't actually pause without a real console attached, so any retry logic burned through instantly. Replaced both with a single retry loop that just keeps retrying the file move itself (a still-open exe naturally blocks it) using a `ping`-based delay instead, and falls back to relaunching the old exe if the move never succeeds, so the app can no longer just vanish after clicking update

### v5.4.17
- Added "Non-Kaid" and "Non-Event" checkboxes to Build → Bank Build → Build Constraints → Only Found In (below Crafted) - hard exclusions that apply regardless of other realm filters. Non-Kaid excludes every Kaid realm (All/White/Green/Red/Purple); Non-Event excludes both Event and Glory Bea items

### v5.4.16
- Added a "Check for Update" button (top of the window, visible on every tab) that checks GitHub's latest release against the running version. If you're current, it just says so; if a newer version exists, the button itself becomes "Download & Update" - click it to download the new exe and have the app automatically swap it in and restart, no manual reinstall needed
- Added a "Download Page" button next to it, for anyone who'd rather download and install a new version manually instead of using the built-in updater

### v5.4.15
- Added a "Credits" button to the Parse tab, in the Parse Results row, far right - opens a popup listing everyone who contributed to the project

### v5.4.14
- Fixed "Your Deaths" missing valid PvP deaths whose killing blow wasn't worded "...at you for N damage!" (e.g. "Aerion attacks you with his fiery hands for 145 damage!") - the match no longer requires any specific phrasing before "damage!", just that the line ends with it
- That same looser match could pick up PvE mob deaths ("A savage owlbear claws you for..."), mislabeled "A Killed you"/"An Killed you"/"You Killed you" - those three are now filtered out of the results

### v5.4.13
- "Your Kills" and "Your Deaths" (Parse tab → PvP) no longer need a player name typed in - both are now plain buttons, like "Participated" already was. Each finds every match across the loaded logs and pulls the player's name straight out of the matched line for the label

### v5.4.12
- Search Logs results (both "Drops only" and plain raw-text search) can now be right-clicked → "Open Log at This Line" - opens the original log file in a new window, scrolled to and highlighting that exact line, with its own Find box (▲/▼ or Enter/Shift+Enter) to search around the rest of the file
- New "PvP" section on the Parse tab, with three searches: **Your Kills** (finds every time you killed a named player), **Your Deaths** (finds every time a named player killed you), and **Participated** (a plain button - finds every realm-points award you got without landing the kill yourself). Right-click any result row here too for "Open Log at This Line"

### v5.4.11
- Every Character Items tab now also has the "Delete" button (bottom-right) that Lockers got in v5.4.10 - permanently removes that character's entire tab and its saved items, same placement and behavior as on a Locker tab

### v5.4.10
- Import's Character Name field now matches an existing character or Locker by spelling only, regardless of capitalization - typing "aria" or "ARIA" when "Aria" already exists updates that same character instead of creating a case-sensitive duplicate
- Each Locker tab has a new "Delete" button (bottom-right) that permanently removes that Locker's entire tab and its saved items - unlike "Clear Saved List", which only empties the list but keeps the tab around. Regular character tabs don't get this button

### v5.4.9
- Fixed the Gear Tag dropdown sometimes not opening on the first click - it previously only got focused, not popped open, so the same click that revealed it didn't expand the list
- Each Character Items tab (and Locker tab) has two new checkboxes: "Good Gear only" and "Invasion Gear only", which limit Find Best Bank Build to just items carrying that Gear Tag. Both can be checked together (either tag qualifies); neither checked runs as normal, with no restriction
- Gear Tag's "Blank" option is back (now alongside "Both") - every item starts out Blank by default. "Both" always qualifies under either of the new checkboxes; "Blank" qualifies under neither, so a user who never uses Gear Tag will find both checkboxes just filter down to nothing until they start tagging

### v5.4.8
- Bank Build's "Saved Items" tab is renamed to "Character Items"
- Added a new "Lockers" tab in Bank Build, alongside Import/Character Items (always the rightmost of the three) - every Locker character now gets its own sub-tab there (with its full existing controls: Only Found In, Prioritize/Hard Search, Find Best Bank Build, Clear, exclude-from-others checkbox) instead of being mixed in among regular character tabs. A character automatically moves to/from Lockers if its Locker checkbox is toggled on a later Import

### v5.4.7
- Weapon Constraints' Weapon Types/Combo's rows each now sit in their own thin bordered box, alternating between two shades of grey row by row, and Fired 1h/Shield moved to the bottom of the list (least commonly used combo)
- Shield Constraints no longer offers a Cloth checkbox - shields are never made of cloth in-game. Leather/Studded/Plate are unaffected

### v5.4.6
- Fixed a bug in "Rebuild (Full Database, Prefer Owned)" (v5.4.4's fallback-fill for empty armor/jewel slots) where a slot could get filled with an item whose spell duplicated a base another slot already covered (e.g. Evade.Enhance i on one slot and Evade.Enhance ii on another) - the two tiers don't stack and shouldn't both appear in the same build. That fallback now only fills a slot with an item carrying no wanted spell at all (or only a Wanted Sigil); it can no longer pick one whose spell is merely redundant with something else already in the build

### v5.4.5
- Basic Constraints has a new "Save Constraints" button (to the right of Only Found In, above Required Items) - names and saves a full snapshot of every current Basic, Armor, and Weapon Constraints selection at once, listed below it with Load/Rename/Delete buttons. Selecting an entry only highlights it; Load applies it explicitly so a stray click can't overwrite what you're still working on. Persists across closing and reopening the program

### v5.4.4
- Empty-slot right-click → "Rebuild (Full Database, Prefer Owned)" now actually fills gaps like Body when doing so needs a new item but doesn't gain any new Wanted Spell coverage - previously an armor/jewel slot could only ever be filled by an item that contributed genuinely new coverage, so a slot only fillable via a "redundant" pick stayed empty no matter how many new items the cap allowed. Also fixed a bug this exposed where the two Jewel slots could get assigned the exact same physical item
- When a cap (1/2/3 new items) has spare budget left over after coverage is already maxed out, it's no longer wasted - extra variants labeled "N new items (alt)" now show otherwise-equally-good unowned items swapped into owned slots one at a time, so there's still something new to compare even once nothing is objectively better to find

### v5.4.3
- The build search now prefers an item whose Sigil isn't already used elsewhere in the build, as a pure last-resort tie-break - it only ever decides between candidates already tied on everything else (coverage, Priority Spells/Tier, Sigil match, Melee stats, Defense, item Tier, and Level), so it can never override a genuinely better item. Applies to the main search, Max Lvl priority slots, and Claw slots. Also fixed a related bug where two otherwise-identical items differing only by Sigil could get silently collapsed into one candidate before the search even ran

### v5.4.2
- A Locker character's tab now has an "Exclude this Locker from other characters' Bank Build searches" checkbox - unchecked by default (matching existing behavior: its non-Kaid gear folds into every other character's search automatically). Checking it opts that one Locker out, for gear you want to keep as pure storage without it silently feeding everyone else's builds. Also respected by "Rebuild (Saved Items First)"

### v5.4.1
- Right-clicking an empty slot in Results now offers "Search Full Database for This Slot" (replacing the old "Rebuild (Full Database)" there) - finds a single best-fit item for just that slot, honoring every active constraint, without touching or recomputing any other slot. Correctly avoids duplicating a wanted spell that's already covered by another slot in the current build, targeting a genuinely uncovered one instead. "Rebuild (Saved Items First)" is unchanged
- Removing an item from the build (right-click → Remove) now always leaves that slot visible showing "No suitable item found" instead of sometimes disappearing entirely
- "Rebuild (Saved Items First)" now correctly excludes a Locker's Kaid items from its trusted pool, matching the same rule every per-character search already follows (a non-Locker character's own saved Kaid items still count fully)
- Also strengthened `OlmranItemBuilder_TEST.exe`'s testing workflow (no user-facing change): its config no longer resets between launches, and Basic Constraints selections now persist alongside Saved Items/characters, so repeat testing starts from a known state instead of empty

### v5.4.0
- Saved Items (Main and every character tab) columns are now sortable - click any header to sort by it, click again to reverse. Level sorts numerically; everything else sorts as text
- Added a "Gear Tag" column to Saved Items: click a row's Tag cell for a dropdown of "Good Gear" / "Invasion Gear" / "Blank". One tag per item name, shared everywhere that item appears (Main and every character), and persists across closing and reopening the program
- Added a "Locker" checkbox to the Import tab - a Locker isn't a character you play, it's extra bank space (a spare character made just to hold overflow gear). A Locker's non-Kaid gear is automatically folded into every other character's Bank Build search, regardless of that character's own "Search all characters" setting
- Results tab has a new "Locker" column between the Bank icon and Slot - shows the first 4 letters of the Locker's name whenever a row's item came from one. The Bank column shows a bold "L" instead of the usual 📦 icon for those rows

### v5.3.2
- The shared "Find Optimal Build" / "Show All Matches" buttons now hide while the Bank Build sub-tab is selected - they'd otherwise run against only Basic/Armor/Weapon Constraints with no bank context, which is redundant now that every character tab has its own "Find Best Bank Build" button. "Generate multiple build options" stays visible either way

### v5.3.1
- Fixed a layout bug in a Bank Build character's Saved Items tab: creating a brand-new character while viewing Bank Build could leave the Only Found In box and the Find Best Bank Build/Clear Saved List buttons clipped off below the visible area (the Build tab's sub-tab area wasn't re-measuring itself when a nested tab like this one appeared, only when you switched sub-tabs directly)
- Each character tab's Only Found In box now sits to the left of its Prioritize/Hard Search/Search all characters checkboxes and buttons instead of stacked above them, using noticeably less vertical space

### v5.3.0
- Bank Build reorganized around characters: a new "Import" tab pastes a bank/inventory listing plus a character name, then saves it to that character's own tab under Saved Items (created automatically the first time the name is used) - the old "Best Build" and "Search" tabs are gone, fully covered by Import + the character tabs now
- Saved Items is now a notebook itself: "Main" is a read-only aggregate of every character's items (plus anything saved before character tracking existed), and each character gets its own tab with its own Only Found In checkboxes, Prioritize/Hard Search, "Find Best Bank Build", and Clear button
- Each character tab has a new "Search all characters" checkbox - unchecked, it searches only that character's own items; checked, it pools every character's non-Kaid items together but still only uses Kaid items from that one tab (since Kaid gear doesn't drop when you die, it represents what that character is actually carrying)
- Everything persists across closing and reopening the program, same as before
- Fixed a layout bug where the Build tab's Basic/Armor/Weapon Constraints sub-tabs showed a large blank gap above the Min/Max/Specific Level controls once Bank Build's content grew taller than theirs - the sub-tab area now resizes to fit whichever sub-tab is actually selected
- The whole window can now be scrolled (mouse wheel, or a scrollbar on the right) if it's resized smaller than its content - nothing gets permanently cut off at the bottom anymore

### v5.2.0
- Results tab: right-click any Build 1 item to remove it from the build - it stays excluded from every future search until a fresh search runs. Two new buttons appear once something's been removed: "Rebuild (Full Database)" (searches everywhere) and "Rebuild (Saved Items First)" (builds as much as possible from Saved Items, then automatically fills whatever it can't cover from the full database - a complete set every time, e.g. after removing a good item you don't want to risk losing in PvP). Right-clicking an empty slot offers the same two Rebuild actions
- The 📦 Bank column icon now shows up on every search's results (Find Optimal Build, Show All Matches, everywhere) whenever an item matches your Saved Items list - not just Bank Build/Saved Items flows
- Bank Build's Best Build and Search tabs each got a "💾 Save to Saved Items" button - parses the paste and updates Saved Items without running the full search/build
- Saved Items now reconciles by content type (bank listing vs. inventory/equipped listing) instead of which paste box was used - updating just your bank paste no longer wipes out inventory-sourced items, and vice versa; a paste with both kinds updates both
- An Inventory listing now also recognizes an unmarked (or stack-count) line if its name matches a real item in the master database, on top of the always-counted "(w)"/"(h)" lines - so unworn gear sitting in inventory gets picked up too, while crafting mats/consumables still don't

### v5.1.11
- Items whose Area is "Class" are now excluded from every search (Find Optimal Build, Show All Matches, Bank Build, Saved Items, Search Missing Slots) - they're just never candidates anywhere
- Added a "Class Items" dropdown next to Required Items (Basic Constraints), listing every excluded item as "Item-MOB-Spell" so one can still be deliberately forced into a build via Required Items - the one way back in for an otherwise-excluded item. The list is sorted by Mob, with Mob shown in bold caps; the closed dropdown stays a fixed width but its popup list widens to fit the longest entry

### v5.1.10
- Saved Items' "Hard Search" now holds each wanted spell's exact tier - it never quietly substitutes a lower/different tier the way every other search does. A slot with nothing at the exact tier shows "No available item" instead
- Results tab: a new "Search Missing Slots (Full Database)" button appears whenever a Hard Search leaves gaps. It fills in just the missing slot(s) in place - every other slot's item, Alt Options, and Bank icon stay exactly as they were - by searching the full database for the same exact tier first; if nothing's found, it asks whether to accept a lower tier instead (yes searches and fills it in, no leaves it alone for another try later). A slot that's searched again and still comes up empty reads "No Items found after re-search"
- Saved Items now has its own "Prioritize Saved Items" / "Hard Search" checkboxes and a "Find Best Bank Build" button, working the same way as Best Build's but sourced straight from the persisted Saved Items list instead of a fresh paste
- Saved Items has a new "Drop" column (between Slot and Item): "No Drop" for any Kaid-realm item, "Drop" for everything else

### v5.1.9
- Fixed a mistake in the Inventory paste format: "(h)" (held) lines are now recognized alongside "(w)" (worn) lines, not ignored

### v5.1.8
- Bank Build's Saved Items tab now searches too: "Prioritize Saved Items" (search everything, favoring them when tied) and "Hard Search" (build only from Saved Items) work the same way Best Build's checkboxes do, just sourced from the persisted list instead of a fresh paste. Hard Search also reports any wanted spell/tier with nothing available as an explicit "No available item" row
- When a Hard Search leaves slots empty, the Results tab shows a new "Search Missing Slots (Full Database)" button - click it to re-search just those gaps against the whole master database, so you can see what to go acquire. It only appears after a Hard Search actually leaves something missing

### v5.1.7
- Bank Build has a third inner tab, "Saved Items": a running, persisted list of every item ever recognized from a Best Build or Search paste. Each new paste updates it automatically - adding items it newly finds, removing ones it no longer sees - and an item that moves between the two tabs' pastes (e.g. it was in the bank, now it's worn) stays on the list instead of disappearing. A second (or third...) copy of the same item is listed at the bottom, prefixed "::extra::"

### v5.1.6
- Bank Build (both "Best Build" and "Search") now recognizes three more paste formats, on top of the original numbered Strongbox listing: an unnumbered Strongbox listing (same `name [Level|Slot|...]` shape, just without the "N.)"), an Inventory listing (only lines marked `(w)` count - anything else, like stack counts or unmarked items, is ignored), and an Items in use listing (`On Head:  item name`, `Held Left:  nothing`). All four can even be mixed in the same paste

### v5.1.5
- Fixed a rendering glitch where the Build tab's bottom controls (Min/Max/Specific Level, the "If no match at Specific Level" radio buttons, and the search buttons) could show garbled/overlapping text after maximizing or restoring the window - a known Windows + ttk 'clam' theme redraw bug, now worked around with a forced repaint on maximize/restore

### v5.1.4
- Only Found In now has an "All" checkbox - an explicit way to say "no realm restriction" instead of relying on every box being left unchecked to mean the same thing. Mutually exclusive with every individual box (checking one greys out the other side)

### v5.1.3
- Bank Build now has two inner tabs: "Best Build" (the existing feature, unchanged) and a new "Search" tab - paste a bank/inventory listing and click Search to list every recognized item as-is (no combo-building), showing which Area each one actually drops in. Includes the same Only Found In checkboxes as Basic Constraints (the same settings, not a separate copy) to narrow it down by realm

### v5.1.2
- Removed the white/gray card background from the program's icon, so its colors show clearly instead of sitting on a white square

### v5.1.1
- Class Specific: added Reverb.enhance
- Class Specific spells no longer offer tier iii in the tier dropdown (none of them go that high) - just (any)/i/ii now
- Added a Bank column (farthest left) to Build Search Results: shows a 📦 icon on any row whose item came from a Bank Build paste - most useful with "Prioritize items I own" checked, where it's the only way to tell which recommended items you already have versus which you'd still need to get
- Bank Build now has an explicit "Only Items I own" checkbox (checked by default) alongside "Prioritize items I own" - the two are mutually exclusive, checking one disables the other
- The program now has a proper icon (window/taskbar and the .exe file itself) instead of the generic default

### v5.1.0
- Added a Bank Build tab (Build > Bank Build): paste a bank/inventory listing (the same format the game's own "Items in Strongbox" listing uses) and click "Find Best Bank Build" to get the best gear combo using it. "Prioritize items I own" unchecked restricts the search to only what you pasted; checked instead searches the full database like a normal build, just favoring owned items over non-owned ones when otherwise close. Every other setting (Wanted Spells, Armor/Weapon Constraints, etc.) still applies exactly as normal on top of this

### v5.0.20
- Added a Weight range to Melee Weapon Constraints (applies to any weapon, including claws): soft preference by default (favors weapons in the range, but won't leave a slot empty over it), or check "Hard Filter" to exclude out-of-range weapons outright

### v5.0.19
- Weapon Constraints' 1 Claw/2 Claw now has two Sigil dropdowns: the 1st applies to the first claw slot (used whether 1 Claw or 2 Claw is checked), the 2nd applies to the second claw slot (only meaningful when 2 Claw is checked) - same soft preference as every other Sigil dropdown
- Find Optimal Build no longer requires a wanted spell/priority spell/wanted sigil/required item to run - a search with none of those set still works (e.g. just checking a Weapon Type/Combo to hunt for weapons alone), instead of being blocked with a "please add at least one" warning

### v5.0.18
- Fixed the .exe being flagged as a Trojan by some antivirus software - it was being compressed with UPX during packaging, which is a common trigger for antivirus false positives (packed executables resemble how real malware hides itself). Built without UPX from here on. No functional changes - same program, just packaged differently

### v5.0.17
- Area Items' Area field now shows live suggestions in a popup as you type (narrows to matching Areas on every keystroke, arrow down into the list, Enter or double-click to pick) instead of only working through a fixed dropdown list
- Area Items results are now sorted: armor slots first (Plate, then Studded, then Leather, then Cloth), then Jewel, then Shield, then Weapon - instead of whatever order they happened to load in
- Clicking/tabbing into the Area field now shows the full alphabetical list of every Area right away (scrollable), same as opening an ordinary dropdown, instead of only appearing once you start typing

### v5.0.16
- Added an Area Items tab (after Saved Builds): pick an Area and browse every item droppable there, straight from the loaded master database
- The Area dropdown is now typeable - type a few letters and it narrows to matching Areas as you go, instead of only scrolling a fixed list
- Fixed a bug from the last update: Results tab's "Remove Area" was reading the wrong column since the Sigil column was added, so it silently stopped removing anything
- Removed the "(more added soon)" note from the Shields/Buffs dropdown

### v5.0.15
- Class Specific: added Aura.enhance
- General Skills: added All.weapons.enhance, Chaos.crush, Leathers.enhance, Platemail.enhance, Slash.enhance, Thrust.enhance
- Moved Weapons.enhance from Class Specific to General Skills

### v5.0.14
- Added a Sigil column to Build Search Results (Find Optimal Build and Show All Matches), showing the chosen item's Sigil whenever it has one
- Reworked how Wanted Sigils (added last update) factors into the search: it's now a secondary consideration behind Wanted Spells rather than an equal-priority requirement - a Wanted Spell always wins a slot over a Wanted Sigil when both are available, and unlike spells, sigils don't need to be "the only one" - multiple slots can each independently carry a wanted Sigil rather than the search treating extra copies as redundant

### v5.0.13
- Added Wanted Sigils to Armor Constraints: pick one or more Sigil types (Cold/Earth/Fire/Lightning/Pain/Shock/Water) and Find Optimal Build actively searches for them, same as Wanted Spells - useful since many armor pieces carry a Sigil but no Spell at all, which previously made them unreachable no matter how good they were
- Exported builds (Excel/HTML/image/text, from both the Results tab and Saved Builds) now show the build's name as a title at the top of the file itself, not just in the filename

### v5.0.12
- Fixed Protect spell tiers (minor/normal/improved, all 8 elements - Cold/Earth/Elemental/Fire/Lightning/Mental/Shock/Water): the real item data stores a Protect's tier differently than every other spell (minor./improved. as a prefix, and no prefix at all for normal), which the search never accounted for - so a Protect item's tier was invisible to it entirely, regardless of which tier was actually picked in the UI. This could surface as the wrong tier (or even a same-slot item for a different element entirely) getting chosen over the one actually requested. Tier and Priority Tier targeting for Protects now works the same as for every other spell

### v5.0.11
- Find Optimal Build now honors a Wanted Spell's own tier for every spell, not just one - adding any spell chip with a specific tier (e.g. "Dexterity ii", "Wisdom iii", etc.) searches for that tier instead of always upgrading to the highest tier available for that spell. Chips added as "(any)" keep the old highest-tier behavior. If a Priority Tier is already set for that spell, the Priority Tier's target still takes precedence over the chip's own tier

### v5.0.10
- Shield Constraints layout refined: Cloth/Leather now share the Sigil row and Studded/Plate share the Defense row, packed tighter together, with Leather and Plate lined up vertically

### v5.0.9
- Shield Constraints now has Cloth/Leather/Studded/Plate checkboxes (one or more can be checked) - a hard filter on the shield's armor type, matching Armor Constraints' own per-slot checkboxes

### v5.0.8
- Weapon and Shield are no longer populated at all if no Weapon Types/Combo's checkbox that implies them is checked - previously they were always filled with the best available item regardless. Two-Handed still fills just the weapon; 1h/Shield, 2h/Shield, and Fired 1h/Shield still fill both

### v5.0.7
- Checking 1 Claw or 2 Claw now excludes Weapon and Shield from the build entirely - claws are their own one-handed weapon, so a claw build no longer also tries to equip a separate physical weapon/shield
- Saved Builds now persist across closing and reopening the program (previously lost every time the app closed)

### v5.0.6
- Find Optimal Build now tries to fit a wanted spell into a Jewel slot before spreading it onto an armor slot - a Jewel's own level doesn't otherwise matter (no armor type, no Defense/Sigil options yet), so it's no longer at a disadvantage against an armor-slot item just because that item happens to have a slightly higher level. A genuinely higher spell tier elsewhere still wins, though - this only breaks ties, not tier differences

### v5.0.5
- Replaced the "No Items For Some Spells" popup - a slot the search couldn't populate at all (given the current constraints, only relevant when at least one wanted spell couldn't be covered anywhere) now shows up in the Results table as "No suitable item found" instead, with the uncovered spell(s) also noted in the status line
- "Generate multiple build options" now produces up to 10 alternate builds (was 5)
- "Save Build" now saves only the top (best) build, not every stacked alternate variant
- Confirmed (not a bug, no change needed): Find Optimal Build already explores every combination of wanted spells regardless of the order they were added - verified this explicitly with a 12-spell/8-slot stress test in three different orderings, all producing identical results

### v5.0.4
- Added an "If no match at Specific Level" fallback policy (Go down a tier / Go down in level / Both / Don't populate slot) for Find Optimal Build - when nothing carries a wanted spell at exactly the level (and tier, if one was requested) you specified, this controls whether the search relaxes level, tier, both, or leaves the slot empty. Always picks the highest available level/closest tier, not just the first match found

### v5.0.3
- Fixed: Crafted-realm items (e.g. Realm = "Crafted - Evil") could show up in results just because another checked realm (like "Evil") happened to be a substring match, even with the Crafted checkbox unchecked, and even when no realm boxes were checked at all - Crafted items now never appear unless the Crafted checkbox is explicitly checked

### v5.0.2
- Alt Options in the Results tab now sorts highest level (leftmost) to lowest (rightmost), instead of whatever incidental order the source data happened to be in
- Added a CodeQL code scanning workflow to the repo (GitHub Security tab)

### v5.0.1
- Fixed the Results tab getting stuck showing "All Matches" after clicking "Find Optimal Build" if a validation warning fired while "All Matches" was still the active view from an earlier search
- Trimmed unused Pillow plugins (AVIF/WebP/CMS/Math/Tk) from the bundled exe - not used anywhere in the app, shrinks it from ~15.5MB to ~13.3MB
- Added a "Max Lvl" priority checkbox to each Armor Constraints slot (not the All: row) - up to 3 at once, greedily locks that slot to the highest-level item that still carries a wanted/priority spell before the normal optimal-build search runs for everything else
- Fixed Armor Constraints' Cloth/Leather/Studded/Plate (and Defense/Sigil) columns not lining up between the All: row and the per-slot rows
- Only Found In: Kaid moved into its own column, renamed "Kaid All" (same behavior as before) with 4 new sub-realm checkboxes below it (Kaid White/Green/Red/Purple), mutually exclusive with Kaid All
- Only Found In: rearranged the left two columns to Evil/Glory Bea, Good/Event, Chaos/Crafted

### v5.0
- Packaged as a standalone `OlmranItemBuilder.exe` via PyInstaller - no Python install, no `pip install` step, no `.bat` launcher files. Just download and double-click; the bundled community equipment list ships inside the .exe itself
- Fixed: Find Optimal Build could pick a weapon/shield/Parry Staff item purely because its own (otherwise irrelevant) spell had a high tier, making the same spell/tier appear to be "duplicated" into a second slot for no real benefit - fallback items no longer get tier/priority credit for a spell that isn't actually contributing new coverage
- Fixed: 1 Claw / 2 Claw did not work at all against the real bundled data (claws are stored as `Slot=weapon`/`Type=claw`, not a distinct `Slot=claw` value, so every claw-handling check was silently matching nothing) - claws now correctly fill, respect Melee Weapon Constraints, and support Required Items
- Moved Bless next to Agility, and added Direct.enhance to Class Specific, in the Basic Constraints spell dropdowns
- Min Level / Max Level / Specific Level fields are now centered in their shared row

### v4.9999999999
- Weapon Constraints redesigned: granular per-combo "Weapon Types/Combo's" (Dual-Wield 1h, 1h/Shield, 2h/Shield, Fired 1h/Shield, Two-Handed, 1 Claw/2 Claw), each with its own Style/Damage Type dropdowns where applicable, replacing the old global Weapon Style radios and Damage Type checkboxes
- Added Melee Weapon Constraints: soft-preference Damage/Timer/Fumble/Accuracy/Sigil dropdowns (each with an optional Priority checkbox, capped at 3) that apply to every weapon style - Melee, Direct, Parry Staff, and Fired alike - and to Claw slots too
- Added Shield Constraints (Defense + Sigil), reusing the same scoring as Armor Constraints' per-slot Defense/Sigil
- Armor Constraints: added per-slot Defense and Sigil soft preferences (never a hard filter - a slot is always still filled)
- Most weapons/claws carry no Spell at all in the source data - weapon/shield/claw slots now always fill with the best available match instead of being left empty when nothing carries a wanted spell
- Min Level / Max Level / Specific Level and the Find Optimal Build / Show All Matches / Generate multiple build options controls moved out of the "Basic Constraints" sub-tab into a shared area visible (and centered) no matter which Build sub-tab is active
- Renamed the Build tab's "Search" sub-tab to "Basic Constraints"
- Items struck through in the source equipment spreadsheet are now automatically skipped on load, so removed/invalid entries never show up in search results
- Default window size trimmed vertically so it opens just below the search buttons instead of with a large empty gap

### v4.0
- Major change to Find Optimal Build: replaced the greedy per-slot search with an exact search that considers every slot at once, so a spell no longer gets stuck at a lower tier in one slot just because a slot processed earlier grabbed the only decent item first while a better-tier item for it sat unclaimed in a different, swappable slot
- Fixed: Weapon Style "Any" could still pick Parry Staff-type weapons even when Parry Staff wasn't selected; staves are now only used when Parry Staff is explicitly chosen
- Updated the bundled `Olmran_Community_Eq_and_Stats_List.xlsx` default equipment list, trimmed to just the Equipment sheet

### v3.1
- Results tab: "Export As..." now supports four formats - Excel Spreadsheet, HTML Page, Image (PNG), and a fixed-width aligned Text Document - not just Excel
- Fixed: exporting a build with multiple stacked build variants used to include the black divider rows as garbage rows of block characters; they're now correctly excluded from every export format
- Added a permanent "Saved Builds" tab: click "Save Build" from the Results tab to add the current results as a panel there (rather than spawning a new notebook tab per save). Each panel has its own renamable name, its own "Export As..." (all four formats), and a Remove button

### v3.0
- Parry Staff is now implemented: it matches staff-type weapons and works like any other gear slot (picked purely by wanted spells), exempt from the Weapon/Shield/Two-Handed/Claw build-config checkboxes and Damage Type constraints
- New spell category: Class Specific populated with 22 combat-enhance skills (backstab, bash, berzerk, crush, etc.)
- Other1 renamed to General Skills, populated with climb/hide/jump/swim/percept/sneak.enhance
- Other2 renamed to Protects, populated with the 8 elemental/mental protect spells; tier dropdown shows minor/normal/improved, mapping directly onto the ordinary i/ii/iii suffix used everywhere else
- Shields/Buffs expanded with Bleed.resist, Disease.resist, Poison.resist; disease.resist specifically supports tier iv
- Fixed: Priority Tier's own tier dropdown wasn't narrowing per selected spell (always showed i/ii/iii) - now matches the category dropdowns' per-spell tier restrictions
- Known issue: some improved-tier (tier iii) Protects don't show up correctly in search results yet - root cause not diagnosed, fix deferred

### v2.9
- Spell matching now qualifies on base spell regardless of tier, so a slot no longer sits empty just because nothing hits the exact tier you requested (e.g. wanting combat.iii but only combat.ii exists) - the "no duplicate spell across slots" rule still applies
- Added Min Tier / Max Tier, next to the level filters, to bound how far that tier fallback is allowed to reach
- Tier now takes priority over level when choosing between candidates - the search always prefers the highest available tier first, only comparing level as a secondary tie-break
- New warning if a wanted spell has no matching item at any tier under your current constraints
- Added Priority Tier: pair a specific spell with a specific tier (e.g. wisdom + ii) so the search targets that tier for that spell specifically, even over a higher tier that's available - other spells are unaffected
- New warning when a Priority Tier can't actually be honored (e.g. "wisdom (ii) cannot be used - tier iii used instead")
- Alt Options column header is now left-aligned instead of centered

### v2.8
- Fixed: the Results tab could appear empty until a different Build Variant was selected - the tab is now switched to before results are inserted
- Fixed: Melee and Direct (Caster) weapon styles could still pick staff-type weapons meant for the unimplemented Parry Staff style
- Fixed: "Generate multiple build options" could produce a variant with the same spell duplicated across two slots at different tiers - tied alternates now must cover the exact same wanted spells, not just the same score, to be considered interchangeable
- Alt Options redesigned: lists any other item in that slot providing the same spell (any tier) within your current level/armor/weapon/realm constraints, shown as "Level - Item Name", with the spell only listed when its tier differs from the item actually picked
- Priority Spell is now Priority Spells (plural): add multiple spells to prioritize, viewable and removable in a chip list below the dropdown; items providing a priority spell are now always searched for and included when possible, not just preferred among items that already matched a wanted spell
- Fixed: results table auto-sizing could make every column the same (oversized) width - divider rows between stacked build variants are now excluded from the size calculation, and columns no longer get squeezed by Treeview's default column-stretch behavior

### v2.7
- Find Optimal Build now prefers the highest level available (targeting Max Level and falling back to progressively lower levels) as a tie-break when items are otherwise equally good
- Alt Options in the results table now shows each alternate's spell alongside its item name
- Results table columns auto-resize to the minimum width needed for their header and current contents
- Added Required Items: force a specific piece of gear into the build by name, with spelling-tolerant fuzzy matching if there's no exact match; the rest of the build is calculated around it
- Wanted Spells and Required Items now sit side by side (Wanted Spells wider, both the same height) instead of stacked full-width
- The Build tab is now scrollable, so armor/weapon/realm constraints and the search controls below Damage Type are reachable regardless of window size

### v2.6
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
