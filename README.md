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

**The .exe takes a few seconds to open every time** — expected. It's a PyInstaller "onefile" build that unpacks itself to a temp folder on every launch, not just the first.

**Program won't start otherwise** — see INSTALL_INSTRUCTIONS.txt.

**Files not loading** — check the file extension (.log, .txt, .xlsx), make sure it isn't open in another program, and try copying it somewhere else if the path has unusual characters.

**Parse results empty** — verify the log actually contains delve output ("You examine X closely") and that CHAT vs ACTION auto-detection picked the right type (check the Parse tab file list).

**Community List not found** — as of v5.0 the community list ships bundled inside `OlmranItemBuilder.exe` and loads automatically; if you're instead running from source (`gaming_log_parser.py`), make sure `Olmran_Community_Eq_and_Stats_List.xlsx` is in the same folder, or use "Browse..." to point at it manually.

## Version History

### v5.3.0 (Current)
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
