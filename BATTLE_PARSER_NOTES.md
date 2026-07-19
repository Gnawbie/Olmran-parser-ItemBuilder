# Battle Parser Notes

**Scope: PVP only.** This feature is specifically for parsing player-vs-player battles, not PVE/mob fights. Keep that in mind when adding patterns - anything that only makes sense in a PVP context (group disbands on death, one player's weapon/spell hitting another player, etc.) is exactly the kind of thing this should key off of. The example batch already collected (Zirak vs. Rhythm vs. Lokust) is PVP.

Working notes for a new feature: parsing **full battles** (start-to-finish combat encounters) out of ACTION/combat logs, not just individual drop/loot lines like the existing parser does.

This is a scratchpad, not a spec — fill it in incrementally as info comes in. Nothing here is implemented yet; this is collection only until a section is marked **Ready to implement**.

## How to use this file

- Add info under whichever class/section it belongs to, as you get it. Paste raw log line examples verbatim when you have them — exact wording/punctuation matters for building correct match patterns later.
- Each class section has three buckets: **Confirmed** (patterns we're sure about), **Needs more examples** (something's expected here but we don't have enough to pin down the pattern yet), and **Not needed** (explicitly ruled out — so we don't waste time hunting for it later).
- Update the Status table at the top as classes move between "not started" → "collecting" → "ready."
- When a class's Confirmed section looks complete and stable, flag it "Ready to implement" and we'll wire it into the actual parser code.

## Status

| Class | Status | Notes |
|---|---|---|
| (none yet) | not started | The *act* of casting is tracked generically (see General section) - the specific spell being cast isn't important, so per-class spell lists likely aren't needed at all for this feature |

## General / Class-Agnostic

Things that should apply no matter which class is fighting — battle start/end detection, shared messages, etc. **The act of casting a spell matters here (something happened); which specific spell it was does not** - so cast lines are tracked generically, not per-class/per-spell.

### Confirmed
- **Spell cast announcement**: `You cast a <Spell.Name> spell!` (the spell name/tier itself is not important - only that a cast happened)
  - Example: `You cast a Dark.Plasma.II spell!`
- **Power/mana cost line**: `You tap your inner reserves for <N> extra power points!`
  - Example: `You tap your inner reserves for 150 extra power points!`
  - Appears right after the cast announcement - likely means the base pool wasn't enough and it drew from a reserve; unclear if this line appears on every cast or only when reserves are actually tapped
- **Melee attack**: `<Attacker> attacks <Target> with <weapon name>!`
  - Example: `Zirak attacks Rhythm with a glowing battlestaff of Al'Tizor!`
  - Same attacker/weapon can appear back-to-back as separate swings (seen twice in a row against the same target in the example batch)
  - Attacker can switch targets mid-battle: `Zirak attacks Lokust with a glowing battlestaff of Al'Tizor!` (same attacker, different target than the lines just before it)
- **Weapon proc/effect**: `<Effect description> from the <weapon name>!`
  - Example: `Fire erupts from the glowing battlestaff of Al'Tizor!` (follows a melee attack line using that same weapon)
- **Damage line**: `<Target> is hit for <N> damage!`
  - Example: `Lokust is hit for 733 damage!`
- **Death**: `<Name> just died!  His corpse lies on the ground.` (note: double space after "died!"; "His"/"Her"/"Its" presumably varies by target gender/type - need more examples to confirm the variants)
  - Example: `Lokust just died!  His corpse lies on the ground.`
- **Death loot-drop announcement**: `<Name> dropped all of his items and money!` (trailing whitespace observed in the example - may just be log formatting noise, not meaningful)
  - Example: `Lokust dropped all of his items and money!  `
- **Group disband on death**: `<Name> disbands from <Leader>'s group.`
  - Example: `Lokust disbands from Rhythm's group.`
- **Buff/effect expiring**: `The <effect name> around <Name> goes away.`
  - Example: `The parrying force around Rhythm goes away.`
- **NPC/mob arrival** (not necessarily combat-specific, but shows up around battles): `<Name> arrives from the <direction>.`
  - Example: `Swampers arrives from the west.`
- **Your own death (from the victim's perspective)** - killing blow line ending in `damage!` (no fixed wording required before it - each class has its own ability/weapon description, e.g. "fiery hands", "lustrous wingclip shortbow", none of which is part of the match), then a blank line, then a fixed death message block (the `N minutes` at the end is unique per death and not part of the match). **Already implemented** as its own small feature (Parse tab → PvP → "Your Deaths") - not waiting on the bigger full-battle parser.
  - Example: `Rhythm fires a lustrous wingclip shortbow at you for 105 damage!`
  - Second wording seen: `Aerion attacks you with his fiery hands for 145 damage!` (confirms the match can't rely on fixed phrasing like "at you for" - only that the line ends in "damage!")
  - Example death block: `You were just killed!  You now float as a ghost / above your dead corpse.  Type RELEASE to complete the / death process to start again in the temple, or try to / find a Priest to resurrect you in the next 12 minutes!`
- **A kill you got (from the killer's perspective)** - `You just killed <Name>!` immediately followed (no blank line, per the example) by `You earn <N> realm points!` (N is 1-3 digits and not part of the match). **Already implemented** alongside PvP Death (Parse tab → PvP → "PvP Kill").
  - Example: `You just killed Hippew!` / `You earn 62 realm points!`
- **Participating in someone else's kill** - any `You earn <N> realm points!` line that ISN'T immediately preceded by `You just killed <name>!` (that combo is a direct kill, already covered by "Your Kills" above) is credit for participating without landing the final blow yourself. No player name needed - it's a plain button. **Already implemented** (Parse tab → PvP → "Participated").

### Needs more examples
- What marks the *start* of a battle, precisely? (First attack line? First damage line? Does an arrival line like "Swampers arrives from the west" ever mark the start of an encounter, or is that unrelated/incidental?)
- What marks the *end* of a battle? Is the death line always the end, or can a battle end other ways (target flees, caster disengages, no-damage timeout)?
- Do multiple simultaneous battles ever interleave in one log, and if so, how do we tell them apart (mob name? a battle/session ID?)
- Confirm the death line's pronoun variants ("His"/"Her"/"Its"/"Their"?) and whether "corpse lies on the ground" text is always identical
- Since this is PVP-only: is there a reliable way to tell a PVP battle apart from a PVE (mob) fight just from the log text itself (e.g. "disbands from group" only makes sense for a player target), or does the parser need an external list of player names to know who's a player vs. a mob?
- Is "goes away" the only wording for a buff/effect expiring, or are there others (e.g. an effect *starting* rather than ending)?

### Not needed
- The specific spell being cast (name/school/tier, e.g. "Dark.Plasma.II") - only that a cast happened matters
- Spell-hit flavor text (the descriptive line unique to a given spell landing on the target, e.g. "You slowly rise into the air as flames of living darkness brand a profane symbol into Lokust's flesh!") - not needed since the specific spell isn't tracked
- Which class cast the spell - not needed for the same reason

## Per-Class Sections

**Likely not needed at all**, given the act-of-casting-not-the-spell clarification above - per-class spell lists were the original plan, but if only the generic action (cast/attack/damage/death) matters, there may be nothing class-specific left to collect. Keeping the template here in case something genuinely class-specific does turn out to matter later (e.g. a class-unique mechanic that isn't just "a cast happened").

### Template
```
### <Class Name>

#### Confirmed
- (spell/ability name) — example log line: `...`
- (spell/ability name) — example log line: `...`

#### Needs more examples
- (ability we know exists but don't have a log example for yet)

#### Not needed
- (explicitly ruled out - e.g. an ability that never appears in combat logs, or isn't relevant to battle tracking)
```

---

## Open Questions

- Should the battle parser output one row per battle (aggregated) or keep every individual combat line too (detailed)?
- Does every class need its own detection logic, or is there a shared "attack/damage" line format that works across classes, with only *special* abilities being class-specific?
- Any existing Combat-tab field/export this should reuse or extend, vs. being a fully separate parse mode?
