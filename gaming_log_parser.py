#!/usr/bin/env python3
"""
Gaming Log Parser - Built for Olmran/MUD-style text game logs
Parses Chat, Combat, and Loot from action and chat .log files
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter import font as tkfont
import re, os, json, difflib
from pathlib import Path
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ─────────────────────────────────────────────────────────────
#  AREA TO REALM MAPPING (from Olmran_Realm_Leveling.xlsx)
# ─────────────────────────────────────────────────────────────
AREA_TO_REALM = {
    'Abandoned Karimere Castle': 'Good', 'Abrishamkar Coast': 'Chaos', 'Abrishamkar Mountains': 'Chaos',
    'Acheronian Keep': 'Good', 'Allorien': 'Good', 'Ancient Crystalline Mines of Fala': 'Good',
    'Arachenlair Forest': 'Chaos', 'Arctic Glaciers': 'Evil', 'Arctic Glaciers Castle': 'Evil',
    'Azure Rainforest': 'Good', 'Barrows of Zden': 'Evil', 'Black Keep': 'Evil', 'Blackwoods': 'Chaos',
    'Bluffs of Zden': 'Evil', 'Carendel': 'Good', 'Carendel Nobles Graveyard': 'Good',
    'Carendel Potter\'s Field': 'Good', 'Cavern of Roots': 'Chaos', 'Cavern of Shadows': 'Good',
    'Caverns of Gaggsith': 'Chaos', 'Cerulean Lakes': 'Evil', 'Chasmanic Swamp': 'Good',
    'Chiligulla Mountains': 'Good', 'Chui Rainforest': 'Good', 'Chui Rainforest Treetops': 'Good',
    'Chui Savanna': 'Good', 'Citadel of Gnovormir': 'Chaos', 'Crescent Moon Gorge': 'Good',
    'Crystal Towers of Fala': 'Good', 'Crystalline Mines': 'Chaos', 'Curdled Blood Marsh': 'Chaos',
    'Cursed Woods': 'Evil', 'Cverick\'s Tower': 'Good', 'Darikor': 'Chaos', 'Darikor Caverns': 'Chaos',
    'Darikor Coast': 'Chaos', 'Darikor Forest': 'Chaos', 'Dark Forest': 'Evil',
    'Dark Heart of Karimere': 'Good', 'Deep Mountain Caves': 'Good', 'Diseased Temple': 'Evil',
    'Doral Coast': 'Chaos', 'Doral Sea': 'Chaos', 'Doral Straits': 'Chaos', 'Dragonclaw Island': 'Chaos',
    'Drakurat\'s Cave': 'Chaos', 'Drakurat\'s Spine': 'Chaos', 'Dunes of Kad\'iril': 'Evil',
    'Eastern Desert': 'Evil', 'Enchants Woods of Acornak': 'Chaos', 'Evermist Knolls': 'Good',
    'Fields of Mo\'Serat': 'Chaos', 'Forests of Shalifi\'s Demise': 'Good',
    'Forests of Shalifi\'s Demise - Mine': 'Good', 'Fort Darkbane': 'Chaos', 'Frozen Wastelands': 'Evil',
    'Gateway to Madness': 'Good', 'Grasslands': 'Evil', 'Grazzt\'s Refuge': 'Good',
    'Grazzt\'s Underground': 'Good', 'Great Desert': 'Evil', 'Great Swamp': 'Evil',
    'Greater Rulan Jungle': 'Chaos', 'Greater Spider Caves': 'Evil', 'Greater Wailing Woods': 'Chaos',
    'Greenmist Forest': 'Evil', 'Greenmist Forest Gateway': 'Evil', 'Greenmist Palisades': 'Evil',
    'Greenmist Palisades - Stockade': 'Evil', 'Grey Mountains': 'Evil', 'Grey Mountains Castle': 'Evil',
    'Grey Mountains Treehouse': 'Evil', 'Grey Mountains Tunnels': 'Evil', 'Greyknife Peaks': 'Evil',
    'Island of Mingo': 'Good', 'Karimere Crypt': 'Good', 'Karimere Highlands': 'Good',
    'Karimere Timber Forest': 'Good', 'Lake Dragonclaw': 'Chaos', 'Lesser Rulan Jungle': 'Chaos',
    'Lord Garon\'s Keep': 'Evil', 'Marshlands': 'Evil', 'Mausoleum': 'Good', 'Medoran Forest': 'Good',
    'Millers Pointe': 'Good', 'Mingo Jungle': 'Good', 'Mingo Mountain': 'Good',
    'Misery\'s Crossing': 'Chaos', 'Motapa Hills': 'Chaos', 'Mount Decadare': 'Chaos',
    'Mount Desparare': 'Evil', 'Mountain Caves': 'Good', 'Myrobi Hills': 'Good', 'Ninja Caves': 'Good',
    'Nogrim, the Deep City': 'Good', 'Nomadic Caravan': 'Chaos', 'Pine Forest/Tundra': 'Evil',
    'Pirate\'s Cove': 'Chaos', 'Plains of Hevak': 'Chaos', 'Poisoned Temple': 'Good',
    'Ravenwood Marsh': 'Chaos', 'Remorse Mountains': 'Good', 'Rigan Coast': 'Evil',
    'Ruins of Gnovormir': 'Chaos', 'Sanctum of the Exile': 'Chaos', 'Saurian Wastes': 'Chaos',
    'Sea of Grass': 'Chaos', 'Sea of Grass Underground': 'Chaos', 'Serpent\'s Pass': 'Evil',
    'Serpentine Mountains': 'Evil', 'Shady Brush Hills': 'Good', 'Shores of the Thundermist': 'Chaos',
    'Shrouded Castle of Craebaen': 'Good', 'Shrouded City of Craebaen': 'Good',
    'Shrouded Forest of Craebaen': 'Good', 'Spider Caves': 'Evil', 'Spider Caves Lower': 'Evil',
    'Stables of Yazik': 'Evil', 'Stonefist Temple': 'Chaos', 'Straits of Mingo': 'Good',
    'Straits of Mingo Islands': 'Good', 'Tamia': 'Evil', 'Tamia Local': 'Evil',
    'Tamian Foothills': 'Evil', 'Teeth of Heaven': 'Good', 'Temple of the Divine Horn': 'Evil',
    'Tenebrous Hollow': 'Chaos', 'Terandhel': 'Good', 'The Bladegrass': 'Good', 'The Catacombs': 'Evil',
    'The Deadlands': 'Chaos', 'The Endless Night': 'Evil', 'The Endless Night City of Zhak-Tor': 'Evil',
    'The Endless Night Sorceror\'s Tower': 'Evil', 'The Endless Night Temple of Klotha': 'Evil',
    'The Hive': 'Evil', 'The Orc Caverns': 'Evil', 'The Pyramid of the Sun': 'Evil',
    'The Rowangroves': 'Evil', 'The Southern Dunes': 'Chaos', 'The Underdark': 'Evil',
    'The Underground Temple': 'Evil', 'The Wailing Woods': 'Chaos', 'The Wildwood': 'Good',
    'The snarewoods': 'Evil', 'Thorn Forest': 'Good', 'Thorny Brush Hills': 'Chaos',
    'Throrfiril': 'Chaos', 'Throrfiril City of Trees': 'Chaos', 'Throrfiril Royal Citadel': 'Chaos',
    'Tomb of the Dragon King': 'Chaos', 'Twisted Jungle': 'Good', 'Unicorn Glen': 'Evil',
    'Unicorn Glen Basement': 'Evil', 'Upper Drakurat\'s Spine': 'Chaos', 'Utilla Town': 'Chaos',
    'Valley of Zden': 'Evil', 'Valley of the Giants': 'Evil', 'Valley of the Giants Volcano': 'Evil',
    'Venomvein Timberland': 'Good', 'Verulian Falls': 'Chaos', 'Verulian Forest': 'Chaos',
    'Village of Reilyn': 'Evil', 'Wastes of Olmran': 'Evil', 'Weald of Decdare': 'Chaos',
    'Wretched Forest': 'Chaos',
}

# ─────────────────────────────────────────────────────────────
#  PARSERS  (one class per log type, using actual file patterns)
# ─────────────────────────────────────────────────────────────

class ChatParser:
    """
    Parses CHAT log files.
    Line format examples:
      <00:23:29> You say '11'
      <00:32:27> [Plan] Terminator: 'yo'
      <00:34:26> Zealut sends, 'ball spells...'
      <00:34:34> You send, 'yes' to Zealut
    """
    CHANNEL_RE = re.compile(
        r'^<(\d{2}:\d{2}:\d{2})>\s+'           # timestamp
        r'(?:\[([^\]]+)\]\s+)?'                  # optional [channel]
        r'(.+?)(?:\s+sends?,\s*\'([^\']+)\'|'    # Name sends, '...'
        r'\s*:\s*\'([^\']+)\'|'                   # Name: '...'
        r'\s+says?\s+\'([^\']+)\')'              # Name says '...'
    )
    SEND_RE = re.compile(
        r'^<(\d{2}:\d{2}:\d{2})>\s+You send,\s*\'([^\']+)\'\s+to\s+(\w+)'
    )
    SAY_RE = re.compile(
        r'^<(\d{2}:\d{2}:\d{2})>\s+You say\s+\'([^\']+)\''
    )
    NAMED_RE = re.compile(
        r'^<(\d{2}:\d{2}:\d{2})>\s+(?:\[([^\]]+)\]\s+)?(\w+)\s*(?:sends?,\s*)?[\':]?\s*\'([^\']+)\''
    )

    @staticmethod
    def parse_file(filepath: str) -> list:
        rows = []
        filename = os.path.basename(filepath)
        with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
            for raw in f:
                line = raw.strip()
                if not line:
                    continue
                ts_m = re.match(r'^<(\d{2}:\d{2}:\d{2})>', line)
                if not ts_m:
                    continue
                ts = ts_m.group(1)
                rest = line[len(ts_m.group(0)):].strip()

                # You send '...' to Name
                m = re.match(r"You send,?\s+'([^']+)'\s+to\s+(\w+)", rest, re.I)
                if m:
                    rows.append({'Timestamp': ts, 'Channel': 'Tell-Out',
                                 'Speaker': 'You', 'Target': m.group(2),
                                 'Message': m.group(1), 'File': filename})
                    continue

                # You say '...'
                m = re.match(r"You say\s+'([^']+)'", rest, re.I)
                if m:
                    rows.append({'Timestamp': ts, 'Channel': 'Say',
                                 'Speaker': 'You', 'Target': '',
                                 'Message': m.group(1), 'File': filename})
                    continue

                # [Channel] Name: '...' or Name sends, '...' or Name says '...'
                m = re.match(r"(?:\[([^\]]+)\]\s+)?(\w+)\s+(?:sends?,\s*)?[\"']([^\"']+)[\"']", rest, re.I)
                if m:
                    ch = m.group(1) or 'Say'
                    # Target from "to Name" suffix
                    tgt_m = re.search(r"\bto\s+(\w+)\s*$", rest, re.I)
                    rows.append({'Timestamp': ts, 'Channel': ch,
                                 'Speaker': m.group(2), 'Target': tgt_m.group(1) if tgt_m else '',
                                 'Message': m.group(3), 'File': filename})
                    continue

                # [Plan] Name: 'msg'   or   [Guild] Name: 'msg'
                m = re.match(r"\[([^\]]+)\]\s+(\w+):\s+'([^']+)'", rest, re.I)
                if m:
                    rows.append({'Timestamp': ts, 'Channel': m.group(1),
                                 'Speaker': m.group(2), 'Target': '',
                                 'Message': m.group(3), 'File': filename})
                    continue

        return rows


class CombatParser:
    """
    Parses ACTION log files for combat events.
    Location comes from [Room Name] lines immediately before the block.

    Extracts:
      - Player attacks mob (smash/cast lines)
      - Mob attacks player
      - Kill events + XP gained
      - Spell casts
    """
    TS_RE   = re.compile(r'^<(\d{2}:\d{2}:\d{2})>')
    LOC_RE  = re.compile(r'^\[([^\]]+)\]')

    MOB_ATK = re.compile(
        r'^(A |An |The )?([A-Za-z ]+?) attacks? (?:you|a knight of dread)(?: and (miss(?:es)?)| for (\d+) damage!)'
    )
    YOU_ATK = re.compile(
        r'^You (?:smash|strike|swing at) (?:a |an |the )?(.+?) (?:with .+? for (\d+) damage!|and miss)'
    )
    SPELL   = re.compile(
        r'^You cast a (.+?) spell!'
    )
    FREEZE  = re.compile(
        r'^(.+?) is hit for (\d+) damage!'
    )
    KILL    = re.compile(
        r'^(.+?) dies!'
    )
    XP      = re.compile(
        r'^You gain (\d+) \(\+(\d+)\) experience points\.'
    )
    HIT_REC = re.compile(
        r'^You are hit for (\d+) damage'
    )

    @staticmethod
    def parse_file(filepath: str) -> list:
        rows = []
        filename = os.path.basename(filepath)
        current_location = ''
        current_ts = ''

        with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
            raw_text = f.read()

        # Split into timestamped blocks
        blocks = re.split(r'\n(?=<\d{2}:\d{2}:\d{2}>)', raw_text)

        for block in blocks:
            lines = [l.rstrip('\r') for l in block.split('\n')]
            if not lines:
                continue

            ts_m = CombatParser.TS_RE.match(lines[0])
            if ts_m:
                current_ts = ts_m.group(1)

            # Check for location header in this block
            for l in lines:
                loc_m = CombatParser.LOC_RE.match(l.strip())
                if loc_m:
                    current_location = loc_m.group(1)
                    break

            # Parse each content line
            for l in lines[1:]:
                l = l.strip()
                if not l:
                    continue

                # Mob attacks you
                m = CombatParser.MOB_ATK.match(l)
                if m:
                    dmg = m.group(4) or '0'
                    hit = 'miss' if m.group(3) else 'hit'
                    rows.append({
                        'Timestamp': current_ts,
                        'Location': current_location,
                        'Attacker': (m.group(1) or '') + (m.group(2) or ''),
                        'Target': 'You',
                        'Action': hit,
                        'Damage': int(dmg) if dmg.isdigit() else 0,
                        'Spell': '',
                        'Event': 'mob_attack',
                        'Full': l,
                        'File': filename
                    })
                    continue

                # You attack a mob
                m = CombatParser.YOU_ATK.match(l)
                if m:
                    rows.append({
                        'Timestamp': current_ts,
                        'Location': current_location,
                        'Attacker': 'You',
                        'Target': m.group(1).strip(),
                        'Action': 'smash' if m.group(2) else 'miss',
                        'Damage': int(m.group(2)) if m.group(2) else 0,
                        'Spell': '',
                        'Event': 'player_attack',
                        'Full': l,
                        'File': filename
                    })
                    continue

                # Spell cast
                m = CombatParser.SPELL.match(l)
                if m:
                    rows.append({
                        'Timestamp': current_ts,
                        'Location': current_location,
                        'Attacker': 'You',
                        'Target': '',
                        'Action': 'cast',
                        'Damage': 0,
                        'Spell': m.group(1),
                        'Event': 'spell_cast',
                        'Full': l,
                        'File': filename
                    })
                    continue

                # Mob/target hit for damage (spell AoE)
                m = CombatParser.FREEZE.match(l)
                if m:
                    rows.append({
                        'Timestamp': current_ts,
                        'Location': current_location,
                        'Attacker': 'You (spell)',
                        'Target': m.group(1).strip(),
                        'Action': 'spell_hit',
                        'Damage': int(m.group(2)),
                        'Spell': '',
                        'Event': 'spell_damage',
                        'Full': l,
                        'File': filename
                    })
                    continue

                # Kill
                m = CombatParser.KILL.match(l)
                if m:
                    rows.append({
                        'Timestamp': current_ts,
                        'Location': current_location,
                        'Attacker': 'You',
                        'Target': m.group(1).strip(),
                        'Action': 'kill',
                        'Damage': 0,
                        'Spell': '',
                        'Event': 'kill',
                        'Full': l,
                        'File': filename
                    })
                    continue

                # XP gained
                m = CombatParser.XP.match(l)
                if m:
                    rows.append({
                        'Timestamp': current_ts,
                        'Location': current_location,
                        'Attacker': '',
                        'Target': '',
                        'Action': 'xp_gain',
                        'Damage': 0,
                        'Spell': '',
                        'Event': f"XP: {m.group(1)} (+{m.group(2)} bonus)",
                        'Full': l,
                        'File': filename
                    })
                    continue

                # Player takes damage (bleed, etc.)
                m = CombatParser.HIT_REC.match(l)
                if m:
                    rows.append({
                        'Timestamp': current_ts,
                        'Location': current_location,
                        'Attacker': 'Environment',
                        'Target': 'You',
                        'Action': 'damage_taken',
                        'Damage': int(m.group(1)),
                        'Spell': '',
                        'Event': 'env_damage',
                        'Full': l,
                        'File': filename
                    })
                    continue

        return rows


class LootParser:
    """
    Parses ACTION log files for loot/drop events.
    Location comes from [Room Name] bracket in same block.

    Patterns from actual log:
      A sand giant drops a shining cracked gauntlets of the sands.
      An Al'Tizor assassin drops a glowing mithril enforcer's bladed spear.
      [GSEND] 'Cardova has completed the 'Level 10 Sin Gear' player achievement!'
    """
    TS_RE   = re.compile(r'^<(\d{2}:\d{2}:\d{2})>')
    LOC_RE  = re.compile(r'^\[([^\]]+)\]')

    DROP_RE  = re.compile(r'^(A |An |The )?(.+?) drops (.+?)\.')
    ACHIEV_RE = re.compile(r"\[GSEND\] '(.+?) has completed the '(.+?)' player achievement!'")
    
    # Item quality and material prefixes to remove
    QUALITY_PREFIXES = re.compile(r'^(silvered|bright|shining|glowing|lustrous|brilliant)\s+', re.IGNORECASE)
    MATERIAL_PREFIXES = re.compile(r'^(bronze|iron|steel|alloy|mithril|laen|wool|cotton|silk|gossamer|wispweave|ebonweave|leather|rough|embossed|suede|wyvern scale|enchanted|maple|oak|yew|rosewood|ironwood|ebony)\s+', re.IGNORECASE)
    
    @staticmethod
    def clean_item_name(item_name: str) -> str:
        """Remove quality and material prefixes from item names."""
        original = item_name
        
        # Remove leading article (a, an, the)
        cleaned = re.sub(r'^(a|an|the)\s+', '', item_name, flags=re.IGNORECASE)
        
        # Try removing quality prefixes
        after_quality = LootParser.QUALITY_PREFIXES.sub('', cleaned)
        
        # Try removing material prefixes
        after_material = LootParser.MATERIAL_PREFIXES.sub('', after_quality)
        
        result = after_material.strip()
        
        # If we're down to just one word, that's probably wrong
        # Keep the material prefix instead
        if result and len(result.split()) == 1:
            # Just remove quality, keep material
            result = after_quality.strip()
        
        return result if result else cleaned.strip()
    
    @staticmethod
    def parse_delve_stats(lines: list) -> dict:
        """Extract stats from delve output in the log."""
        stats = {
            'Slot': '',
            'Type': '',
            'Spell': '',
            'Level': '',
            'Damage': '',
            'Timer': '',
            'Fumble': '',
            'Accuracy': '',
            'Defense': '',
            'Sigil': '',
            'SigilLvl': '',
            'Weight': '',
            'Holdable': False,
            'HoldableWhileCasting': False,
            'Hands': '',
            'Worth': ''
        }
        
        for line in lines:
            line = line.strip()
            
            # Weapon type: "This object is a weapon of the two-handed thrusting type."
            m = re.search(r'weapon of the (.+?)\s+type', line, re.I)
            if m:
                weapon_desc = m.group(1).strip()
                # Extract just the last word (the actual type: thrusting, slashing, etc.)
                type_match = re.search(r'(\w+)$', weapon_desc)
                if type_match:
                    stats['Type'] = type_match.group(1)
                stats['Slot'] = 'weapon'
                continue
            
            # Hands: "This object takes two hands to wield."
            if re.search(r'two hands to wield', line, re.I):
                stats['Hands'] = '2h'
                continue
            
            # Armor slot: "This object can be worn on the Hands."
            m = re.search(r'worn on the (\w+)', line, re.I)
            if m:
                stats['Slot'] = m.group(1).lower()
                continue
            
            # Armor type: "This object is armor of the cloth type."
            m = re.search(r'armor of the (\w+) type', line, re.I)
            if m:
                stats['Type'] = m.group(1)
                continue
            
            # Spell: "This object seems to cast the spell: Agility.II."
            # Preserve case to keep .I .II .III intact
            m = re.search(r'cast the spell:\s*(.+?)\.?\s*$', line, re.I)
            if m:
                spell_name = m.group(1).strip()
                # Don't add generic words as spells
                if spell_name.lower() not in ['melee', 'shield', 'armor', 'weapon']:
                    stats['Spell'] = spell_name  # Keep original case
                continue
            
            # Sigil: "This item seems to be enchanted with a tier 2 sigil of Psionic Protection!"
            m = re.search(r'tier (\d+) sigil of ([^!]+)', line, re.I)
            if m:
                stats['SigilLvl'] = m.group(1)
                stats['Sigil'] = m.group(2).strip()
                continue
            
            # Level: "This item seems to require level 60 to use effectively."
            m = re.search(r'require level (\d+)', line, re.I)
            if m:
                stats['Level'] = m.group(1)
                continue
            
            # Damage: "This item seems to do a very good amount of damage."
            m = re.search(r'do a (.+?) amount of damage', line, re.I)
            if m:
                stats['Damage'] = m.group(1).strip()
                continue
            
            # Fumble: "This item seems to fumble rarely."
            m = re.search(r'fumble (\w+)', line, re.I)
            if m:
                stats['Fumble'] = m.group(1).strip()
                continue
            
            # Accuracy: "This item seems to hit in combat better than normal."
            m = re.search(r'hit in combat (.+?)\.', line, re.I)
            if m:
                accuracy_desc = m.group(1).strip()
                # Clean up "better than normal" -> "better", "about like normal" -> "normal"
                accuracy_desc = accuracy_desc.replace(' than normal', '').replace('about like ', '')
                stats['Accuracy'] = accuracy_desc.strip()
                continue
            
            # Defense/Protection: "This item seems to protect in combat about like normal."
            m = re.search(r'protect in combat (.+?)\.', line, re.I)
            if m:
                defense_desc = m.group(1).strip()
                # Clean up "about like normal" -> "normal"
                defense_desc = defense_desc.replace('about like ', '').replace(' than normal', '')
                stats['Defense'] = defense_desc.strip()
                continue
            
            # Weight: "This item seems to weigh 6 pounds." or "This item seems to weigh 2 pounds."
            m = re.search(r'weigh (\d+)', line, re.I)
            if m:
                stats['Weight'] = m.group(1)
                continue
            
            # Holdable while casting: "This item seems to be holdable while casting spells."
            if re.search(r'holdable while casting', line, re.I):
                stats['HoldableWhileCasting'] = True
                stats['Holdable'] = True
                continue
            
            # Holdable: "This object can be held"
            if re.search(r'can be held|holdable', line, re.I):
                stats['Holdable'] = True
                continue
            
            # Worth/Value: "This item seems to be worth between 23000 and 24000 silver."
            m = re.search(r'worth between (\d+) and (\d+) silver', line, re.I)
            if m:
                # Use the average of the range
                low = int(m.group(1))
                high = int(m.group(2))
                avg = (low + high) // 2
                stats['Worth'] = str(avg)
                continue
        
        # Post-process: Build Type field for weapons
        if stats['Slot'] == 'weapon' and stats['Type']:
            # Build Type: [direct] <type> <hands>
            type_parts = []
            
            # Add "direct" prefix if holdable while casting
            if stats['HoldableWhileCasting']:
                type_parts.append('direct')
            
            # Add weapon type (slashing, crushing, thrusting, etc.)
            type_parts.append(stats['Type'])
            
            # Add hands (1h or 2h)
            if stats['Hands']:
                type_parts.append(stats['Hands'])
            else:
                # Default to 1h if not specified
                type_parts.append('1h')
            
            stats['Type'] = ' '.join(type_parts)
        
        return stats

    @staticmethod
    def parse_file(filepath: str) -> list:
        rows = []
        filename = os.path.basename(filepath)

        with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
            raw_text = f.read()

        # Split into timestamp blocks
        blocks = re.split(r'\n(?=<\d{2}:\d{2}:\d{2}>)', raw_text)
        
        # First pass: Find all delves and extract stats
        delved_items = {}  # clean_item_name -> delve_stats
        
        for block in blocks:
            lines = [l.rstrip('\r') for l in block.split('\n')]
            
            # Look for "You examine a <item> closely." which indicates delve output
            delve_item_name = None
            for line in lines:
                m = re.search(r'You examine (?:a |an |the )?(.+?) closely', line, re.I)
                if m:
                    full_item_name = m.group(1).strip()
                    clean_item_name = LootParser.clean_item_name(full_item_name)
                    delve_item_name = clean_item_name
                    break
            
            if delve_item_name:
                # Extract stats from this delve block
                stats = LootParser.parse_delve_stats(lines)
                delved_items[delve_item_name] = stats
        
        # Second pass: Find drops of delved items
        # Build a list of all area markers and drop events with their positions
        area_markers = []  # (line_index, area_name)
        drop_events = []   # (line_index, drop_line, item_clean, mob_name, timestamp)
        
        line_index = 0
        current_ts = ''
        
        for block in blocks:
            lines = [l.rstrip('\r') for l in block.split('\n')]
            
            # Get timestamp from first line
            ts_m = LootParser.TS_RE.match(lines[0])
            if ts_m:
                current_ts = ts_m.group(1)
            
            for line in lines:
                # Track area markers
                loc_m = LootParser.LOC_RE.match(line.strip())
                if loc_m:
                    area_markers.append((line_index, loc_m.group(1)))
                
                # Track drops
                m = LootParser.DROP_RE.match(line.strip())
                if m:
                    item_raw = m.group(3).strip()
                    
                    # Skip silver/coins
                    if re.match(r'a bag of (silver|coins?|gold)', item_raw, re.I):
                        line_index += 1
                        continue
                    
                    item_clean = LootParser.clean_item_name(item_raw)
                    
                    # Only care about delved items
                    if item_clean in delved_items:
                        mob_name = ((m.group(1) or '') + m.group(2)).strip()
                        mob_name = re.sub(r'^(A|An|The)\s+', '', mob_name, flags=re.IGNORECASE).strip()
                        drop_events.append((line_index, line.strip(), item_clean, mob_name, current_ts))
                
                line_index += 1
        
        # Third pass: Match drops to their nearest preceding area
        for drop_line_idx, drop_line, item_name, mob_name, timestamp in drop_events:
            # Find the closest area marker before this drop
            closest_area = ''
            for area_idx, area_name in area_markers:
                if area_idx < drop_line_idx:
                    closest_area = area_name
                else:
                    break  # Past the drop point
            
            # Get delve stats for this item
            delve_stats = delved_items.get(item_name, {})
            
            # Check if this is a fodder/crafting item (has only weight, no other stats)
            has_only_weight = (
                delve_stats.get('Weight') and
                not delve_stats.get('Slot') and
                not delve_stats.get('Type') and
                not delve_stats.get('Spell') and
                not delve_stats.get('Level') and
                not delve_stats.get('Damage') and
                not delve_stats.get('Defense')
            )
            
            if has_only_weight:
                if delve_stats.get('Holdable'):
                    # Holdable crafting item
                    delve_stats['Slot'] = 'possible'
                    delve_stats['Type'] = 'crafting'
                    delve_stats['Spell'] = 'item'
                else:
                    # Non-holdable fodder
                    delve_stats['Slot'] = 'fodder'
                    delve_stats['Type'] = 'fodder'
            
            realm = AREA_TO_REALM.get(closest_area, '') if closest_area else ''
            
            # Build notes with sell price if available
            notes = ''
            if delve_stats.get('Worth'):
                notes = f"sells for {delve_stats['Worth']}"
            
            rows.append({
                'Realm': realm,
                'Area': closest_area,
                'Mob': mob_name,
                'Item': item_name,
                'Slot': delve_stats.get('Slot', ''),
                'Type': delve_stats.get('Type', ''),
                'Spell': delve_stats.get('Spell', ''),
                'Level': delve_stats.get('Level', ''),
                'Damage': delve_stats.get('Damage', ''),
                'Timer': '',
                'Fumble': delve_stats.get('Fumble', ''),
                'Accuracy': delve_stats.get('Accuracy', ''),
                'Defense': delve_stats.get('Defense', ''),
                'Sigil': delve_stats.get('Sigil', ''),
                'SigilLvl': delve_stats.get('SigilLvl', ''),
                'Weight': delve_stats.get('Weight', ''),
                'Notes': notes,
                'Timestamp': timestamp,
                'Location': closest_area,
                'Source': mob_name,
                'Quantity': '1',
                'Silver': '',
                'Event': 'mob_drop',
                'Full': drop_line,
                'File': filename
            })

        # Deduplicate rows based on Realm, Area, Mob, Item
        # Keep first occurrence of each unique combination
        seen = set()
        deduplicated = []
        for row in rows:
            key = (row['Realm'], row['Area'], row['Mob'], row['Item'])
            if key not in seen:
                seen.add(key)
                deduplicated.append(row)

        return deduplicated


# ─────────────────────────────────────────────────────────────
#  EXCEL EXPORT
# ─────────────────────────────────────────────────────────────

COLORS = {
    'chat':   'FF4472C4',  # blue
    'combat': 'FFED7D31',  # orange
    'loot':   'FF70AD47',  # green
}

def write_sheet(ws, rows: list, fields: list, sheet_type: str):
    """Write rows to a worksheet using the field configuration."""
    if not rows:
        ws.append(['No data found for this log type.'])
        return

    accent = COLORS.get(sheet_type, 'FF4472C4')
    hdr_font  = Font(name='Arial', bold=True, color='FFFFFFFF', size=10)
    hdr_fill  = PatternFill(fill_type='solid', start_color=accent, end_color=accent)
    hdr_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    data_font = Font(name='Arial', size=10)
    thin      = Side(style='thin', color='FFD9D9D9')
    border    = Border(left=thin, right=thin, top=thin, bottom=thin)

    # Build header row using field positions
    # fields = list of {'label': str, 'source_key': str, 'col': int, 'row_mode': bool}
    max_col = max(f['col'] for f in fields) if fields else 1
    headers = [''] * max_col
    for f in fields:
        headers[f['col'] - 1] = f['label']

    ws.append(headers)
    for ci, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=ci)
        cell.font = hdr_font
        cell.fill = hdr_fill
        cell.alignment = hdr_align
        cell.border = border

    ws.row_dimensions[1].height = 22

    # Data rows
    alt_fill = PatternFill(fill_type='solid', start_color='FFF2F2F2', end_color='FFF2F2F2')
    for ri, row_data in enumerate(rows, 2):
        values = [''] * max_col
        for f in fields:
            values[f['col'] - 1] = str(row_data.get(f['source_key'], ''))
        ws.append(values)
        for ci in range(1, max_col + 1):
            cell = ws.cell(row=ri, column=ci)
            cell.font = data_font
            cell.border = border
            cell.alignment = Alignment(vertical='top', wrap_text=False)
            if ri % 2 == 0:
                cell.fill = alt_fill

    ws.freeze_panes = 'A2'

    # Auto-width
    for ci in range(1, max_col + 1):
        col_letter = get_column_letter(ci)
        best = len(str(headers[ci - 1])) + 2
        for ri in range(2, min(len(rows) + 2, 200)):
            val = ws.cell(row=ri, column=ci).value or ''
            best = max(best, min(len(str(val)) + 2, 60))
        ws.column_dimensions[col_letter].width = best


# ─────────────────────────────────────────────────────────────
#  GUI
# ─────────────────────────────────────────────────────────────

DEFAULT_FIELDS = {
    'chat': [
        {'label': 'Timestamp',  'source_key': 'Timestamp', 'col': 1},
        {'label': 'Channel',    'source_key': 'Channel',   'col': 2},
        {'label': 'Speaker',    'source_key': 'Speaker',   'col': 3},
        {'label': 'Target',     'source_key': 'Target',    'col': 4},
        {'label': 'Message',    'source_key': 'Message',   'col': 5},
        {'label': 'File',       'source_key': 'File',      'col': 6},
    ],
    'combat': [
        {'label': 'Timestamp',  'source_key': 'Timestamp', 'col': 1},
        {'label': 'Location',   'source_key': 'Location',  'col': 2},
        {'label': 'Attacker',   'source_key': 'Attacker',  'col': 3},
        {'label': 'Target',     'source_key': 'Target',    'col': 4},
        {'label': 'Action',     'source_key': 'Action',    'col': 5},
        {'label': 'Damage',     'source_key': 'Damage',    'col': 6},
        {'label': 'Spell',      'source_key': 'Spell',     'col': 7},
        {'label': 'Event',      'source_key': 'Event',     'col': 8},
        {'label': 'File',       'source_key': 'File',      'col': 9},
    ],
    'loot': [
        {'label': 'Realm',      'source_key': 'Realm',      'col': 1},
        {'label': 'Area',       'source_key': 'Area',       'col': 2},
        {'label': 'Mob',        'source_key': 'Mob',        'col': 3},
        {'label': 'Item',       'source_key': 'Item',       'col': 4},
        {'label': 'Slot',       'source_key': 'Slot',       'col': 5},
        {'label': 'Type',       'source_key': 'Type',       'col': 6},
        {'label': 'Spell',      'source_key': 'Spell',      'col': 7},
        {'label': 'Level',      'source_key': 'Level',      'col': 8},
        {'label': 'Damage',     'source_key': 'Damage',     'col': 9},
        {'label': 'Timer',      'source_key': 'Timer',      'col': 10},
        {'label': 'Fumble',     'source_key': 'Fumble',     'col': 11},
        {'label': 'Accuracy',   'source_key': 'Accuracy',   'col': 12},
        {'label': 'Defense',    'source_key': 'Defense',    'col': 13},
        {'label': 'Sigil',      'source_key': 'Sigil',      'col': 14},
        {'label': 'SigilLvl',   'source_key': 'SigilLvl',   'col': 15},
        {'label': 'Weight',     'source_key': 'Weight',     'col': 16},
        {'label': 'Notes',      'source_key': 'Notes',      'col': 17},
    ],
}

CHAT_SOURCES   = ['Timestamp','Channel','Speaker','Target','Message','File']
COMBAT_SOURCES = ['Timestamp','Location','Attacker','Target','Action','Damage','Spell','Event','Full','File']
LOOT_SOURCES   = ['Realm','Area','Mob','Item','Slot','Type','Spell','Level','Damage','Timer','Fumble','Accuracy','Defense','Sigil','SigilLvl','Weight','Notes','Timestamp','Location','Source','Quantity','Silver','Event','Full','File']

# Spell name options for the Build tab category dropdowns. Populate the empty
# lists (Class Specific, Other1, Other2) whenever the full spell lists are
# available - no other code changes are needed, the dropdowns pick them up automatically.
SPELL_CATEGORIES = {
    'Basic':          ['Agility', 'Dexterity', 'Constitution', 'Intelligence', 'Wisdom', 'Strength', 'Bless', 'Evade', 'Combat'],
    'Class Specific': [],
    'Shields/Buffs':  ['Protect', 'Blur', 'Shield', 'Tough.skin', 'Vitalize', 'Regenerate'],
    'Other1':         [],
    'Other2':         [],
}

SPELL_TIERS = ['(any)', 'i', 'ii', 'iii']

# Categories whose list is populated but not yet exhaustive - shown next to the dropdown.
SPELL_CATEGORY_PARTIAL_NOTE = {
    'Shields/Buffs': '(more added soon)',
}

# Spells (lowercase) that only go up to a certain tier - the tier dropdown is
# restricted to these values whenever the spell is selected. Default is SPELL_TIERS.
SPELL_TIER_RESTRICTIONS = {
    'agility': ['(any)', 'i', 'ii'],
    'bless': ['(any)', 'i', 'ii'],
}

# A build can include at most this many items from the "Crafted" realm.
MAX_CRAFTED_ITEMS = 1

# Cap on how many alternate full-build variants "Generate multiple build
# options" will produce (in addition to the primary optimal build).
MAX_BUILD_VARIANTS = 5

# Spells (lowercase) whose underlying spell-column value differs from the
# display name itself, e.g. Evade is stored as "evade.enhance".
SPELL_VALUE_OVERRIDES = {
    'evade': 'evade.enhance',
}

_TIER_RANK = {'i': 1, 'ii': 2, 'iii': 3}


def _spell_base(spell):
    """Strip a trailing .i/.ii/.iii tier suffix so different tiers of the same
    wanted spell are recognized as one requirement, not two."""
    m = re.match(r'^(.*)\.(i|ii|iii)$', spell.strip().lower())
    return m.group(1) if m else spell.strip().lower()


def _spell_tier_rank(variant):
    """Rank of the tier suffix on a wanted-spell string (0 if untiered/any)."""
    m = re.match(r'^.*\.(i|ii|iii)$', variant.strip().lower())
    return _TIER_RANK[m.group(1)] if m else 0


class ItemMatchDialog(tk.Toplevel):
    """Modal dialog listing close-spelling matches for a searched item name,
    letting the user confirm which one they actually meant."""
    def __init__(self, parent, query: str, matches: list):
        super().__init__(parent)
        self.result = None
        self.title("Did you mean...?")
        self.resizable(False, False)
        self.grab_set()

        ttk.Label(self, text=f"No exact match for '{query}'. Closest matches:",
                 font=('Arial', 9)).pack(anchor='w', padx=10, pady=(10,4))

        self.listbox = tk.Listbox(self, height=min(6, len(matches)), width=50)
        for m in matches:
            self.listbox.insert(tk.END, m)
        self.listbox.selection_set(0)
        self.listbox.pack(padx=10, pady=4)

        bf = ttk.Frame(self)
        bf.pack(pady=10)
        ttk.Button(bf, text="Use Selected", command=self._confirm).pack(side='left', padx=8)
        ttk.Button(bf, text="Cancel", command=self.destroy).pack(side='left', padx=8)

    def _confirm(self):
        sel = self.listbox.curselection()
        if sel:
            self.result = self.listbox.get(sel[0])
        self.destroy()


class FieldEditorDialog(tk.Toplevel):
    """Modal dialog to add or edit a single field."""
    def __init__(self, parent, mode: str, existing: dict = None):
        super().__init__(parent)
        self.result = None
        self.mode = mode
        source_opts = {'chat': CHAT_SOURCES, 'combat': COMBAT_SOURCES, 'loot': LOOT_SOURCES}[mode]

        self.title("Edit Field" if existing else "Add Field")
        self.resizable(False, False)
        self.grab_set()

        pad = {'padx': 10, 'pady': 6}

        ttk.Label(self, text="Column label (header text):").grid(row=0, column=0, sticky='w', **pad)
        self.lbl_var = tk.StringVar(value=existing['label'] if existing else '')
        ttk.Entry(self, textvariable=self.lbl_var, width=26).grid(row=0, column=1, **pad)

        ttk.Label(self, text="Data source:").grid(row=1, column=0, sticky='w', **pad)
        self.src_var = tk.StringVar(value=existing['source_key'] if existing else source_opts[0])
        ttk.Combobox(self, textvariable=self.src_var, values=source_opts,
                     state='readonly', width=24).grid(row=1, column=1, **pad)

        ttk.Label(self, text="Column position (1=A, 2=B…):").grid(row=2, column=0, sticky='w', **pad)
        self.col_var = tk.IntVar(value=existing['col'] if existing else 1)
        ttk.Spinbox(self, textvariable=self.col_var, from_=1, to=30,
                    width=6).grid(row=2, column=1, sticky='w', **pad)

        bf = ttk.Frame(self)
        bf.grid(row=3, column=0, columnspan=2, pady=10)
        ttk.Button(bf, text="Save", command=self._save).pack(side='left', padx=8)
        ttk.Button(bf, text="Cancel", command=self.destroy).pack(side='left', padx=8)

    def _save(self):
        lbl = self.lbl_var.get().strip()
        if not lbl:
            messagebox.showwarning("Missing", "Enter a label.", parent=self)
            return
        self.result = {'label': lbl, 'source_key': self.src_var.get(), 'col': self.col_var.get()}
        self.destroy()


class SnapshotViewer(tk.Toplevel):
    """Shows the raw parsed rows for one sheet type so user can verify."""
    def __init__(self, parent, title: str, rows: list, fields: list):
        super().__init__(parent)
        self.title(f"Snapshot — {title}  ({len(rows)} rows)")
        self.geometry("1100x540")

        if not rows:
            ttk.Label(self, text="No data parsed yet.  Run Parse first.",
                      font=('Arial', 12)).pack(expand=True)
            return

        cols = [f['label'] for f in fields]
        keys = [f['source_key'] for f in fields]

        frame = ttk.Frame(self)
        frame.pack(fill='both', expand=True)

        tv = ttk.Treeview(frame, columns=cols, show='headings', height=22)
        for c in cols:
            tv.heading(c, text=c)
            tv.column(c, width=max(80, len(c) * 9), stretch=True)

        vsb = ttk.Scrollbar(frame, orient='vertical',   command=tv.yview)
        hsb = ttk.Scrollbar(frame, orient='horizontal', command=tv.xview)
        tv.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        tv.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)

        for row in rows[:2000]:
            tv.insert('', 'end', values=[str(row.get(k, '')) for k in keys])

        ttk.Label(self, text=f"Showing up to 2 000 of {len(rows)} rows",
                  foreground='gray').pack(pady=4)


class LogSearchViewer(tk.Toplevel):
    """Shows every raw log line matching a search term, across all loaded files."""
    def __init__(self, parent, query: str, results: list):
        super().__init__(parent)
        self.title(f"Search Results — '{query}'  ({len(results)} matches)")
        self.geometry("1150x540")

        cols = ('File', 'Line #', 'Timestamp', 'Text')
        frame = ttk.Frame(self)
        frame.pack(fill='both', expand=True)

        tv = ttk.Treeview(frame, columns=cols, show='headings', height=22)
        widths = {'File': 220, 'Line #': 70, 'Timestamp': 90, 'Text': 700}
        for c in cols:
            tv.heading(c, text=c)
            tv.column(c, width=widths[c], anchor='center' if c == 'Line #' else 'w')

        vsb = ttk.Scrollbar(frame, orient='vertical',   command=tv.yview)
        hsb = ttk.Scrollbar(frame, orient='horizontal', command=tv.xview)
        tv.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        tv.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)

        for file_name, line_num, timestamp, text in results[:5000]:
            tv.insert('', 'end', values=(file_name, line_num, timestamp, text))

        shown = min(len(results), 5000)
        ttk.Label(self, text=f"Showing {shown} of {len(results)} matching lines",
                  foreground='gray').pack(pady=4)


class DropSnapshotViewer(tk.Toplevel):
    """Lists every drop event found for an item; selecting one shows a snapshot
    of the raw log from the last timestamp through that drop line."""
    def __init__(self, parent, query: str, drops: list):
        super().__init__(parent)
        self.title(f"Drop Search — '{query}'  ({len(drops)} drops)")
        self.geometry("1150x650")
        self.drops = drops

        paned = ttk.PanedWindow(self, orient='vertical')
        paned.pack(fill='both', expand=True)

        list_frame = ttk.Frame(paned)
        paned.add(list_frame, weight=1)

        cols = ('File', 'Timestamp', 'Item', 'Mob')
        self.tv = ttk.Treeview(list_frame, columns=cols, show='headings', height=12)
        widths = {'File': 220, 'Timestamp': 90, 'Item': 320, 'Mob': 220}
        for c in cols:
            self.tv.heading(c, text=c)
            self.tv.column(c, width=widths[c], anchor='center' if c == 'Timestamp' else 'w')

        vsb = ttk.Scrollbar(list_frame, orient='vertical', command=self.tv.yview)
        self.tv.configure(yscrollcommand=vsb.set)
        self.tv.pack(side='left', fill='both', expand=True)
        vsb.pack(side='right', fill='y')

        for i, d in enumerate(drops):
            self.tv.insert('', 'end', iid=str(i),
                           values=(d['file'], d['timestamp'], d['item'], d['mob']))
        self.tv.bind('<<TreeviewSelect>>', self._on_select)

        snapshot_frame = ttk.LabelFrame(paned, text="Snapshot (last timestamp → drop)", padding=6)
        paned.add(snapshot_frame, weight=1)

        self.snapshot_text = tk.Text(snapshot_frame, wrap='none', height=10, state='disabled')
        snap_vsb = ttk.Scrollbar(snapshot_frame, orient='vertical', command=self.snapshot_text.yview)
        snap_hsb = ttk.Scrollbar(snapshot_frame, orient='horizontal', command=self.snapshot_text.xview)
        self.snapshot_text.configure(yscrollcommand=snap_vsb.set, xscrollcommand=snap_hsb.set)
        self.snapshot_text.grid(row=0, column=0, sticky='nsew')
        snap_vsb.grid(row=0, column=1, sticky='ns')
        snap_hsb.grid(row=1, column=0, sticky='ew')
        snapshot_frame.rowconfigure(0, weight=1)
        snapshot_frame.columnconfigure(0, weight=1)

        ttk.Label(self, text=f"{len(drops)} drop(s) found — select a row above to view its snapshot",
                  foreground='gray').pack(pady=4)

        if drops:
            self.tv.selection_set('0')
            self._on_select()

    def _on_select(self, event=None):
        sel = self.tv.selection()
        if not sel:
            return
        drop = self.drops[int(sel[0])]
        self.snapshot_text.config(state='normal')
        self.snapshot_text.delete('1.0', tk.END)
        self.snapshot_text.insert('1.0', drop['snapshot'])
        self.snapshot_text.config(state='disabled')


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("🎮 Gaming Log Parser")
        self.geometry("1350x1000")
        self.minsize(1150, 750)
        self.resizable(True, True)

        self.files: list[dict] = []       # {path, name, size, type}
        self.parsed = {'chat': [], 'combat': [], 'loot': []}
        self.fields = {k: [dict(f) for f in v] for k, v in DEFAULT_FIELDS.items()}
        
        # Config file for persistent settings
        self.config_file = os.path.join(os.path.expanduser("~"), ".gaming_log_parser_config.json")
        
        # Load saved directories or use home as default
        self._load_config()

        self._build_ui()
        
        # Save config on close
        self.protocol("WM_DELETE_WINDOW", self._on_closing)
    
    def _load_config(self):
        """Load saved configuration from file"""
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r') as f:
                    config = json.load(f)
                    self.last_open_dir = config.get('last_open_dir', os.path.expanduser("~"))
                    self.last_save_dir = config.get('last_save_dir', os.path.expanduser("~"))
                    self.armor_defaults = config.get('armor_defaults', {})
                    self.weapon_defaults = config.get('weapon_defaults', {})
                    self.weapon_style_default = config.get('weapon_style_default', '')
                    self.build_config_defaults = config.get('build_config_defaults', {})
                    self.dual_wield_default = config.get('dual_wield_default', False)
            else:
                self.last_open_dir = os.path.expanduser("~")
                self.last_save_dir = os.path.expanduser("~")
                self.armor_defaults = {}
                self.weapon_defaults = {}
                self.weapon_style_default = ''
                self.build_config_defaults = {}
                self.dual_wield_default = False
        except Exception:
            self.last_open_dir = os.path.expanduser("~")
            self.last_save_dir = os.path.expanduser("~")
            self.armor_defaults = {}
            self.weapon_defaults = {}
            self.weapon_style_default = ''
            self.build_config_defaults = {}
            self.dual_wield_default = False
    
    def _save_config(self):
        """Save configuration to file"""
        try:
            config = {
                'last_open_dir': self.last_open_dir,
                'last_save_dir': self.last_save_dir,
                'armor_defaults': getattr(self, 'armor_defaults', {}),
                'weapon_defaults': getattr(self, 'weapon_defaults', {}),
                'weapon_style_default': getattr(self, 'weapon_style_default', ''),
                'build_config_defaults': getattr(self, 'build_config_defaults', {}),
                'dual_wield_default': getattr(self, 'dual_wield_default', False)
            }
            with open(self.config_file, 'w') as f:
                json.dump(config, f)
        except Exception:
            pass  # Silently fail if can't save config
    
    def _on_closing(self):
        """Handle window closing"""
        self._save_config()
        self.destroy()

    # ── BUILD UI ──────────────────────────────────────────────
    def _build_ui(self):
        style = ttk.Style(self)
        style.theme_use('clam')
        style.configure('TNotebook.Tab', font=('Arial', 10, 'bold'), padding=[12, 6])
        style.configure('Accent.TButton', font=('Arial', 10, 'bold'))

        nb = ttk.Notebook(self)
        nb.pack(fill='both', expand=True, padx=8, pady=8)

        self.tab_parse  = ttk.Frame(nb, padding=12)
        self.tab_fields = ttk.Frame(nb, padding=12)
        self.tab_export = ttk.Frame(nb, padding=12)
        self.tab_build = ttk.Frame(nb, padding=12)
        self.tab_results = ttk.Frame(nb, padding=12)

        nb.add(self.tab_parse,  text='▶  Parse')
        nb.add(self.tab_fields, text='⚙  Fields')
        nb.add(self.tab_export, text='💾  Export')
        nb.add(self.tab_build, text='🔨  Build')
        nb.add(self.tab_results, text='📊  Results')  # Always visible
        
        self.notebook = nb  # Store reference to notebook

        self._build_parse_tab()
        self._build_fields_tab()
        self._build_export_tab()
        self._build_build_tab()
        self._build_results_tab()

    # ── PARSE TAB (FILES + PARSING) ───────────────────────────
    def _build_parse_tab(self):
        t = self.tab_parse
        
        # ═══ FILES SECTION ═══
        ttk.Label(t, text="Load Log Files",
                  font=('Arial', 13, 'bold')).pack(anchor='w', pady=(0, 8))

        top = ttk.Frame(t)
        top.pack(fill='x')
        ttk.Button(top, text="➕ Add Files",   command=self._add_files).pack(side='left', padx=(0,6))
        ttk.Button(top, text="📁 Add Folder",  command=self._add_folder).pack(side='left', padx=(0,6))
        ttk.Button(top, text="❌ Remove",      command=self._remove_files).pack(side='left', padx=(0,6))
        ttk.Button(top, text="🗑 Clear All",   command=self._clear_files).pack(side='left')
        self.file_count_lbl = ttk.Label(top, text="", foreground='#444')
        self.file_count_lbl.pack(side='right')

        cols = ('File', 'Type', 'Size')
        frame = ttk.Frame(t)
        frame.pack(fill='both', expand=True, pady=8)

        self.file_tv = ttk.Treeview(frame, columns=cols, show='headings', height=8)
        self.file_tv.heading('File', text='Filename')
        self.file_tv.heading('Type', text='Auto-Detected Type')
        self.file_tv.heading('Size', text='Size')
        self.file_tv.column('File', width=380)
        self.file_tv.column('Type', width=150)
        self.file_tv.column('Size', width=90, anchor='e')

        vsb = ttk.Scrollbar(frame, command=self.file_tv.yview)
        self.file_tv.configure(yscrollcommand=vsb.set)
        self.file_tv.pack(side='left', fill='both', expand=True)
        vsb.pack(side='right', fill='y')

        ttk.Label(t, text="💡 File types are auto-detected (CHAT = Chat logs, ACTION = Action logs)",
                 font=('Arial', 8, 'italic'), foreground='#666').pack(anchor='w', pady=(0,12))

        # ═══ SEARCH SECTION ═══
        # Search across every loaded file's actual lines - independent of Run
        # Parse, so you can e.g. find every time an item dropped across a large
        # batch of logs without parsing them all first.
        search_frame = ttk.LabelFrame(t, text="Search Logs", padding=10)
        search_frame.pack(fill='x', pady=(0,12))

        search_row = ttk.Frame(search_frame)
        search_row.pack(fill='x')
        ttk.Label(search_row, text="Find:", width=8).pack(side='left')
        self.search_query_var = tk.StringVar(value='')
        search_entry = ttk.Entry(search_row, textvariable=self.search_query_var, width=40)
        search_entry.pack(side='left', padx=4)
        search_entry.bind('<Return>', lambda e: self._search_logs())
        self.search_case_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(search_row, text="Case sensitive",
                       variable=self.search_case_var).pack(side='left', padx=8)
        self.search_drops_only_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(search_row, text="Drops only (with snapshot)",
                       variable=self.search_drops_only_var).pack(side='left', padx=8)
        ttk.Button(search_row, text="🔍 Search",
                  command=self._search_logs).pack(side='left', padx=4)

        ttk.Label(search_frame,
                 text="💡 \"Drops only\" catches every instance the item actually dropped (ignoring quality/material "
                      "prefixes like \"bright\" or \"glowing\", same as the Loot parser), and for each one shows a "
                      "snapshot of everything from the last timestamp through the drop line. Uncheck it to just "
                      "search every loaded file's raw text for any mention of the term.",
                 font=('Arial', 8, 'italic'), foreground='#666', wraplength=900,
                 justify='left').pack(anchor='w', pady=(4,0))

        # ═══ PARSING SECTION ═══
        parse_frame = ttk.LabelFrame(t, text="Parse Options", padding=10)
        parse_frame.pack(fill='x', pady=(0,12))
        
        ttk.Label(parse_frame, text="What to Parse:", 
                 font=('Arial', 9)).pack(anchor='w', pady=(0,4))
        
        opts_frame = ttk.Frame(parse_frame)
        opts_frame.pack(anchor='w', padx=20)
        
        self.do_chat = tk.BooleanVar(value=False)
        self.do_combat = tk.BooleanVar(value=False)
        self.do_loot   = tk.BooleanVar(value=False)
        
        self.chat_cb = ttk.Checkbutton(opts_frame, text="💬 Chat", 
                                       variable=self.do_chat)
        self.chat_cb.pack(side='left', padx=(0,20))
        
        self.combat_cb = ttk.Checkbutton(opts_frame, text="⚔️ Combat", 
                                         variable=self.do_combat)
        self.combat_cb.pack(side='left', padx=(0,20))
        
        self.loot_cb = ttk.Checkbutton(opts_frame, text="💎 Loot", 
                                       variable=self.do_loot)
        self.loot_cb.pack(side='left')

        run_frame = ttk.Frame(t)
        run_frame.pack(pady=8)
        ttk.Button(run_frame, text="▶  Run Parse",
                   command=self._run_parse, width=22, style='Accent.TButton').pack(side='left', padx=6)
        ttk.Button(run_frame, text="👁  Snapshot: Chat",
                   command=lambda: self._show_snapshot('chat')).pack(side='left', padx=6)
        ttk.Button(run_frame, text="👁  Snapshot: Combat",
                   command=lambda: self._show_snapshot('combat')).pack(side='left', padx=6)
        ttk.Button(run_frame, text="👁  Snapshot: Loot",
                   command=lambda: self._show_snapshot('loot')).pack(side='left', padx=6)

        self.status_var = tk.StringVar(value="Load files and click 'Run Parse' to begin")
        ttk.Label(t, textvariable=self.status_var,
                  font=('Arial', 10), foreground='#222').pack(anchor='w', pady=(12,0))

        # Mini stats
        self.stats_frame = ttk.LabelFrame(t, text="Parse Results", padding=10)
        self.stats_frame.pack(fill='x', pady=(12,0))
        
        self.stat_labels = {}
        for i, (k, icon) in enumerate([('chat','💬'),('combat','⚔️'),('loot','💎')]):
            f = ttk.Frame(self.stats_frame)
            f.grid(row=0, column=i, padx=20, pady=6)
            ttk.Label(f, text=f"{icon} {k.title()}", font=('Arial', 10, 'bold')).pack()
            lbl = ttk.Label(f, text="0 rows", foreground='#555')
            lbl.pack()
            self.stat_labels[k] = lbl

    # ── FIELDS TAB ────────────────────────────────────────────
    def _build_fields_tab(self):
        t = self.tab_fields
        ttk.Label(t, text="Customize Output Columns",
                  font=('Arial', 13, 'bold')).pack(anchor='w', pady=(0,6))
        ttk.Label(t, text="Each row below is one column in the exported spreadsheet.\n"
                           "Drag rows to reorder, or use Add/Edit/Remove. Column position sets the Excel column (1=A).",
                  foreground='#444').pack(anchor='w', pady=(0,8))

        ttk.Label(t, 
                 text="💡 Editing fields for types selected in Parse tab (Chat/Combat/Loot checkboxes)", 
                 font=('Arial', 9, 'italic'), foreground='#666').pack(anchor='w', pady=(0,8))

        # Create notebook for multiple field types
        self.fields_notebook = ttk.Notebook(t)
        self.fields_notebook.pack(fill='both', expand=True, pady=(0,6))
        
        # Create tabs for each type
        self.field_frames = {}
        self.field_tvs = {}
        
        for field_type, icon in [('chat', '💬'), ('combat', '⚔️'), ('loot', '💎')]:
            frame = ttk.Frame(self.fields_notebook, padding=12)
            self.fields_notebook.add(frame, text=f"{icon} {field_type.title()}")
            
            # Treeview for this type
            tree_frame = ttk.Frame(frame)
            tree_frame.pack(fill='both', expand=True, pady=(0,6))
            
            cols = ('Col #', 'Header Label', 'Data Source')
            tv = ttk.Treeview(tree_frame, columns=cols, show='headings', height=13)
            for c, w in zip(cols, [70, 220, 220]):
                tv.heading(c, text=c)
                tv.column(c, width=w)
            vsb = ttk.Scrollbar(tree_frame, command=tv.yview)
            tv.configure(yscrollcommand=vsb.set)
            tv.pack(side='left', fill='both', expand=True)
            vsb.pack(side='right', fill='y')
            
            self.field_frames[field_type] = frame
            self.field_tvs[field_type] = tv

        # Buttons (work on currently visible tab)
        btns = ttk.Frame(t)
        btns.pack(anchor='w')
        ttk.Button(btns, text="➕ Add",          command=self._add_field).pack(side='left', padx=4)
        ttk.Button(btns, text="✏️ Edit",         command=self._edit_field).pack(side='left', padx=4)
        ttk.Button(btns, text="❌ Remove",       command=self._remove_field).pack(side='left', padx=4)
        ttk.Button(btns, text="⬆ Move Up",      command=self._move_up).pack(side='left', padx=4)
        ttk.Button(btns, text="⬇ Move Down",    command=self._move_down).pack(side='left', padx=4)
        ttk.Button(btns, text="🔄 Reset Defaults", command=self._reset_fields).pack(side='left', padx=4)

        self._refresh_all_fields()

    # ── EXPORT TAB ────────────────────────────────────────────
    def _build_export_tab(self):
        t = self.tab_export
        ttk.Label(t, text="Export to Excel",
                  font=('Arial', 13, 'bold')).pack(anchor='w', pady=(0,10))

        info_frame = ttk.Frame(t)
        info_frame.pack(fill='x', pady=(0,12))
        ttk.Label(info_frame, 
                 text="💡 Export settings follow Parse tab selections (Chat/Combat/Loot checkboxes)", 
                 font=('Arial', 9, 'italic'), foreground='#666').pack(anchor='w')
        
        # Summary sheet option
        summary_frame = ttk.Frame(t)
        summary_frame.pack(fill='x', pady=(0,12))
        self.exp_summary = tk.BooleanVar(value=False)
        ttk.Checkbutton(summary_frame, text="📋 Include Summary Sheet", 
                       variable=self.exp_summary).pack(anchor='w')

        # Master file option
        master_frame = ttk.LabelFrame(t, text="Master Database (Loot Only)", padding=10)
        master_frame.pack(fill='x', pady=(0,12))
        
        checkbox_frame = ttk.Frame(master_frame)
        checkbox_frame.pack(anchor='w', pady=4)
        
        self.use_master = tk.BooleanVar(value=False)
        ttk.Checkbutton(checkbox_frame, text="📚 Append to Master Database", 
                       variable=self.use_master).pack(side='left', padx=(0, 20))
        
        self.exclude_fodder = tk.BooleanVar(value=True)
        ttk.Checkbutton(checkbox_frame, text="🚫 Exclude Fodder", 
                       variable=self.exclude_fodder).pack(side='left')
        
        ttk.Label(master_frame, text="Master file:", 
                 font=('Arial', 9)).pack(side='left', padx=(0,8))
        self.master_path_var = tk.StringVar(value='')
        ttk.Entry(master_frame, textvariable=self.master_path_var, 
                 width=35, state='readonly').pack(side='left', padx=4)
        ttk.Button(master_frame, text="Browse...", 
                  command=self._browse_master).pack(side='left', padx=2)
        ttk.Button(master_frame, text="Create New", 
                  command=self._create_new_master).pack(side='left', padx=2)
        
        ttk.Label(master_frame, 
                 text="💡 Master database builds a complete item list across all sessions",
                 font=('Arial', 8, 'italic'), foreground='#666').pack(anchor='w', pady=(4,0))

        fname_frame = ttk.Frame(t)
        fname_frame.pack(fill='x', pady=6)
        ttk.Label(fname_frame, text="Output filename:").pack(side='left')
        self.fname_var = tk.StringVar(value='game_log_export')
        ttk.Entry(fname_frame, textvariable=self.fname_var, width=32).pack(side='left', padx=8)
        ttk.Label(fname_frame, text=".xlsx").pack(side='left')
        
        # Export options
        export_opts_frame = ttk.Frame(t)
        export_opts_frame.pack(fill='x', pady=(0,8))
        
        self.separate_action_exports = tk.BooleanVar(value=False)
        ttk.Checkbutton(export_opts_frame, 
                       text="📂 Export Combat and Loot as separate files (when both selected)", 
                       variable=self.separate_action_exports).pack(anchor='w')

        ttk.Button(t, text="💾  Export to Excel",
                   command=self._export, width=26).pack(pady=8)

        self.export_status = ttk.Label(t, text="", foreground='#226622', font=('Arial', 10))
        self.export_status.pack(anchor='w')
    
    def _browse_master(self):
        """Browse for master database file"""
        path = filedialog.askopenfilename(
            title='Select Master Database File',
            initialdir=self.last_save_dir,
            filetypes=[('Excel', '*.xlsx'), ('All', '*.*')])
        if path:
            self.master_path_var.set(path)
            self.last_save_dir = os.path.dirname(path)
    
    def _create_new_master(self):
        """Create a new master database file"""
        # Suggest filename
        from datetime import datetime
        timestamp = datetime.now().strftime('%Y%m%d')
        default_name = f"Master_Loot_Database_{timestamp}"
        
        path = filedialog.asksaveasfilename(
            title='Create New Master Database',
            initialdir=self.last_save_dir,
            initialfile=default_name + '.xlsx',
            defaultextension='.xlsx',
            filetypes=[('Excel', '*.xlsx'), ('All', '*.*')])
        
        if not path:
            return
        
        try:
            # Create new Excel file with Loot sheet
            wb = openpyxl.Workbook()
            wb.remove(wb.active)  # Remove default sheet
            
            ws = wb.create_sheet("Loot")
            
            # Write header row using loot fields
            from openpyxl.styles import Font, PatternFill, Alignment
            
            headers = []
            max_col = max(f['col'] for f in self.fields['loot'])
            headers = [''] * max_col
            for f in self.fields['loot']:
                headers[f['col'] - 1] = f['label']
            
            ws.append(headers)
            
            # Style header
            hdr_font = Font(name='Arial', size=11, bold=True, color='FFFFFF')
            hdr_fill = PatternFill(fill_type='solid', start_color='4472C4', end_color='4472C4')
            for col_idx, header in enumerate(headers, 1):
                cell = ws.cell(1, col_idx)
                cell.font = hdr_font
                cell.fill = hdr_fill
                cell.alignment = Alignment(horizontal='center', vertical='center')
            
            ws.freeze_panes = 'A2'
            
            # Save
            wb.save(path)
            
            # Set as master file
            self.master_path_var.set(path)
            self.use_master.set(True)
            self.last_save_dir = os.path.dirname(path)
            
            messagebox.showinfo("Master Created", 
                f"✓ New master database created!\n\n{path}\n\n"
                "Now parse your logs and export - new items will be added to this master file.")
                
        except Exception as e:
            messagebox.showerror("Create Error", f"Failed to create master file:\n{str(e)}")
    
    def _append_to_master(self):
        """Append new loot items to master database, avoiding duplicates"""
        master_path = self.master_path_var.get()
        if not os.path.exists(master_path):
            messagebox.showwarning("Master File Not Found", 
                f"Master database file not found:\n{master_path}\n\nSkipping master append.")
            return
        
        try:
            # Load existing master file
            wb = openpyxl.load_workbook(master_path)
            
            # Find or create Loot sheet
            if 'Loot' in wb.sheetnames:
                ws = wb['Loot']
            else:
                messagebox.showwarning("No Loot Sheet", 
                    "Master file has no 'Loot' sheet. Skipping master append.")
                return
            
            # Read existing items to avoid duplicates
            # Key: (Realm, Area, Mob, Item)
            existing_items = set()
            for row in range(2, ws.max_row + 1):  # Skip header
                realm = ws.cell(row, 1).value or ''
                area = ws.cell(row, 2).value or ''
                mob = ws.cell(row, 3).value or ''
                item = ws.cell(row, 4).value or ''
                existing_items.add((realm, area, mob, item))
            
            # Add new items that aren't already in master
            new_count = 0
            skipped_fodder = 0
            for loot_item in self.parsed['loot']:
                # Check if we should exclude fodder
                if self.exclude_fodder.get():
                    slot = loot_item.get('Slot', '').lower()
                    if slot == 'fodder':
                        skipped_fodder += 1
                        continue
                
                key = (loot_item['Realm'], loot_item['Area'], 
                      loot_item['Mob'], loot_item['Item'])
                if key not in existing_items:
                    # Append new row
                    row_num = ws.max_row + 1
                    for field in self.fields['loot']:
                        col = field['col']
                        value = loot_item.get(field['source_key'], '')
                        ws.cell(row_num, col, value)
                    new_count += 1
                    existing_items.add(key)  # Prevent duplicates within this session
            
            # Save master file
            wb.save(master_path)
            
            # Show results
            message = f"✓ Added {new_count} new items to master database"
            if skipped_fodder > 0:
                message += f"\n🚫 Excluded {skipped_fodder} fodder items"
            message += f"\n\n{master_path}"
            
            if new_count > 0:
                messagebox.showinfo("Master Updated", message)
            else:
                if skipped_fodder > 0:
                    messagebox.showinfo("Master Up-to-Date", 
                        f"No new items to add\n🚫 Excluded {skipped_fodder} fodder items")
                else:
                    messagebox.showinfo("Master Up-to-Date", 
                        "No new items to add (all items already in master)")
                    
        except Exception as e:
            messagebox.showerror("Master Append Error", 
                f"Failed to append to master:\n{str(e)}")


    # ── FILE HELPERS ──────────────────────────────────────────
    def _add_files(self):
        paths = filedialog.askopenfilenames(
            title='Select log files',
            initialdir=self.last_open_dir,
            filetypes=[('Log/Text', '*.log *.txt'), ('All', '*.*')])
        
        if paths:
            # Save the directory for next time
            self.last_open_dir = os.path.dirname(paths[0])
        
        for p in paths:
            if not any(f['path'] == p for f in self.files):
                name = os.path.basename(p)
                # Default to loot for ACTION files, chat for CHAT files
                ftype = 'loot' if 'ACTION' in name.upper() else \
                        'chat' if 'CHAT' in name.upper() else 'loot'
                self.files.append({'path': p, 'name': name,
                                   'size': os.path.getsize(p), 'type': ftype})
        self._refresh_file_list()
        self._update_parse_options()

    def _add_folder(self):
        folder = filedialog.askdirectory(
            title='Select folder',
            initialdir=self.last_open_dir)
        if not folder:
            return
        
        # Save the directory for next time
        self.last_open_dir = folder
        
        for ext in ('*.log', '*.txt'):
            for p in Path(folder).glob(ext):
                sp = str(p)
                if not any(f['path'] == sp for f in self.files):
                    name = p.name
                    # Default to loot for ACTION files
                    ftype = 'loot' if 'ACTION' in name.upper() else \
                            'chat' if 'CHAT' in name.upper() else 'loot'
                    self.files.append({'path': sp, 'name': name,
                                       'size': os.path.getsize(sp), 'type': ftype})
        self._refresh_file_list()
        self._update_parse_options()

    def _remove_files(self):
        for iid in self.file_tv.selection():
            idx = self.file_tv.index(iid)
            if 0 <= idx < len(self.files):
                del self.files[idx]
        self._refresh_file_list()
        self._update_parse_options()

    def _clear_files(self):
        if self.files and messagebox.askyesno("Clear", "Remove all loaded files?"):
            self.files.clear()
            self._refresh_file_list()
            self._update_parse_options()

    def _update_parse_options(self):
        """Update parse checkboxes based on loaded files"""
        # Check what types of files are loaded
        has_chat = any(f['type'] == 'chat' for f in self.files)
        has_action = any(f['type'] in ('combat', 'loot') for f in self.files)
        
        if not self.files:
            # No files: uncheck and enable all
            self.do_chat.set(False)
            self.do_combat.set(False)
            self.do_loot.set(False)
            self.chat_cb.config(state='normal')
            self.combat_cb.config(state='normal')
            self.loot_cb.config(state='normal')
        elif has_chat and not has_action:
            # Only chat files: check chat, disable action options
            self.do_chat.set(True)
            self.do_combat.set(False)
            self.do_loot.set(False)
            self.chat_cb.config(state='normal')
            self.combat_cb.config(state='disabled')
            self.loot_cb.config(state='disabled')
        elif has_action and not has_chat:
            # Only action files: check loot, disable chat
            self.do_chat.set(False)
            self.do_combat.set(False)
            self.do_loot.set(True)
            self.chat_cb.config(state='disabled')
            self.combat_cb.config(state='normal')
            self.loot_cb.config(state='normal')
        else:
            # Both types: check both, enable all
            self.do_chat.set(True)
            self.do_combat.set(False)
            self.do_loot.set(True)
            self.chat_cb.config(state='normal')
            self.combat_cb.config(state='normal')
            self.loot_cb.config(state='normal')
    
    def _select_all_files(self):
        """Select all files in the file list"""
        for item in self.file_tv.get_children():
            self.file_tv.selection_add(item)
    
    def _apply_type(self):
        t = self.type_var.get()
        for iid in self.file_tv.selection():
            idx = self.file_tv.index(iid)
            if 0 <= idx < len(self.files):
                self.files[idx]['type'] = t
        self._refresh_file_list()

    def _refresh_file_list(self):
        self.file_tv.delete(*self.file_tv.get_children())
        for f in self.files:
            kb = f['size'] / 1024
            self.file_tv.insert('', 'end',
                values=(f['name'], f['type'], f"{kb:.1f} KB"))
        n = len(self.files)
        self.file_count_lbl.config(text=f"{n} file{'s' if n != 1 else ''} loaded")

    # ── PARSE HELPERS ─────────────────────────────────────────
    def _run_parse(self):
        if not self.files:
            messagebox.showwarning("No Files", "Load some log files first.")
            return

        self.parsed = {'chat': [], 'chat_files': {}, 'combat': [], 'loot': []}
        errors = []

        for f in self.files:
            ftype = f['type']
            try:
                # Auto-parse based on detected file type and checkboxes
                if ftype == 'chat' and self.do_chat.get():
                    chat_data = ChatParser.parse_file(f['path'])
                    self.parsed['chat'] += chat_data
                    # Store separately by filename for individual sheets
                    self.parsed['chat_files'][f['name']] = chat_data
                elif ftype == 'combat' and self.do_combat.get():
                    self.parsed['combat'] += CombatParser.parse_file(f['path'])
                elif ftype == 'loot' and self.do_loot.get():
                    self.parsed['loot'] += LootParser.parse_file(f['path'])
            except Exception as e:
                errors.append(f"{f['name']}: {e}")

        for k, lbl in self.stat_labels.items():
            if k in self.parsed:
                lbl.config(text=f"{len(self.parsed[k])} rows")

        msg = (f"✓  Chat: {len(self.parsed['chat'])} rows  |  "
               f"Combat: {len(self.parsed['combat'])} rows  |  "
               f"Loot: {len(self.parsed['loot'])} rows")
        if errors:
            msg += f"\n⚠ Errors: {'; '.join(errors)}"
        self.status_var.set(msg)

    def _show_snapshot(self, mode: str):
        rows = self.parsed.get(mode, [])
        if not rows:
            messagebox.showinfo("Snapshot",
                f"No {mode} data yet.\nRun Parse first (with {mode} files loaded).")
            return
        SnapshotViewer(self, mode.title(), rows, self.fields[mode])

    def _search_logs(self):
        """Search every loaded file for a term, e.g. an item name. In "Drops
        only" mode (default), finds every actual drop event of that item and
        shows a snapshot from the last timestamp through the drop; otherwise
        falls back to a plain raw-text search of every line."""
        query = self.search_query_var.get().strip()
        if not query:
            messagebox.showwarning("No Search Term", "Enter something to search for.")
            return
        if not self.files:
            messagebox.showwarning("No Files", "Load some log files first.")
            return

        if self.search_drops_only_var.get():
            drops, errors = self._find_item_drops(query)
            if not drops:
                msg = f"No drops of '{query}' were found in {len(self.files)} file(s)."
                if errors:
                    msg += f"\n\n{len(errors)} file(s) could not be read: " + "; ".join(errors)
                messagebox.showinfo("No Matches", msg)
                return
            DropSnapshotViewer(self, query, drops)
        else:
            results, errors = self._find_raw_lines(query)
            if not results:
                msg = f"No lines containing '{query}' were found in {len(self.files)} file(s)."
                if errors:
                    msg += f"\n\n{len(errors)} file(s) could not be read: " + "; ".join(errors)
                messagebox.showinfo("No Matches", msg)
                return
            LogSearchViewer(self, query, results)

        if errors:
            messagebox.showwarning("Some Files Skipped",
                f"{len(errors)} file(s) could not be read:\n" + "\n".join(errors))

    def _find_raw_lines(self, query):
        """Plain substring search of every loaded file's raw lines"""
        case_sensitive = self.search_case_var.get()
        needle = query if case_sensitive else query.lower()

        results = []
        errors = []
        for f in self.files:
            try:
                with open(f['path'], 'r', encoding='utf-8', errors='ignore') as fh:
                    for line_num, line in enumerate(fh, 1):
                        haystack = line if case_sensitive else line.lower()
                        if needle in haystack:
                            ts_match = re.match(r'^<(\d{2}:\d{2}:\d{2})>', line)
                            timestamp = ts_match.group(1) if ts_match else ''
                            results.append((f['name'], line_num, timestamp, line.strip()))
            except Exception as e:
                errors.append(f"{f['name']}: {e}")
        return results, errors

    def _find_item_drops(self, query):
        """Find every actual drop event of an item across all loaded files,
        matching the same way the Loot parser identifies drops (so quality
        and material prefixes like "bright" or "glowing" don't cause misses).
        Each result includes a snapshot of every line from the last timestamp
        seen through the drop line itself."""
        case_sensitive = self.search_case_var.get()
        needle = query if case_sensitive else query.lower()

        drops = []
        errors = []
        for f in self.files:
            try:
                with open(f['path'], 'r', encoding='utf-8', errors='ignore') as fh:
                    lines = fh.readlines()
            except Exception as e:
                errors.append(f"{f['name']}: {e}")
                continue

            current_ts = ''
            block_start = 0
            for idx, line in enumerate(lines):
                ts_m = LootParser.TS_RE.match(line)
                if ts_m:
                    current_ts = ts_m.group(1)
                    block_start = idx

                m = LootParser.DROP_RE.match(line.strip())
                if not m:
                    continue

                item_raw = m.group(3).strip()
                if re.match(r'a bag of (silver|coins?|gold)', item_raw, re.I):
                    continue

                item_clean = LootParser.clean_item_name(item_raw)
                haystack = item_clean if case_sensitive else item_clean.lower()
                if needle not in haystack:
                    continue

                mob_name = ((m.group(1) or '') + m.group(2)).strip()
                mob_name = re.sub(r'^(A|An|The)\s+', '', mob_name, flags=re.IGNORECASE).strip()

                snapshot = ''.join(lines[block_start:idx + 1]).rstrip('\n')
                drops.append({
                    'file': f['name'],
                    'timestamp': current_ts,
                    'item': item_clean,
                    'mob': mob_name,
                    'snapshot': snapshot,
                })

        return drops, errors

    # ── FIELD HELPERS ─────────────────────────────────────────
    def _get_current_field_mode(self):
        """Get the currently selected field type from notebook tab"""
        tab_index = self.fields_notebook.index(self.fields_notebook.select())
        return ['chat', 'combat', 'loot'][tab_index]
    
    def _refresh_all_fields(self):
        """Refresh all field treeviews"""
        for mode in ['chat', 'combat', 'loot']:
            tv = self.field_tvs[mode]
            tv.delete(*tv.get_children())
            for f in sorted(self.fields[mode], key=lambda x: x['col']):
                tv.insert('', 'end',
                         values=(f['col'], f['label'], f['source_key']))
    
    def _refresh_fields(self):
        """Refresh current field treeview"""
        mode = self._get_current_field_mode()
        tv = self.field_tvs[mode]
        tv.delete(*tv.get_children())
        for f in sorted(self.fields[mode], key=lambda x: x['col']):
            tv.insert('', 'end',
                     values=(f['col'], f['label'], f['source_key']))

    def _add_field(self):
        mode = self._get_current_field_mode()
        dlg = FieldEditorDialog(self, mode)
        self.wait_window(dlg)
        if dlg.result:
            self.fields[mode].append(dlg.result)
            self._refresh_fields()

    def _edit_field(self):
        mode = self._get_current_field_mode()
        tv = self.field_tvs[mode]
        sel = tv.selection()
        if not sel:
            return
        idx = tv.index(sel[0])
        flist = sorted(self.fields[mode], key=lambda x: x['col'])
        if idx >= len(flist):
            return
        dlg = FieldEditorDialog(self, mode, existing=flist[idx])
        self.wait_window(dlg)
        if dlg.result:
            orig = flist[idx]
            orig_idx = self.fields[mode].index(orig)
            self.fields[mode][orig_idx] = dlg.result
            self._refresh_fields()

    def _remove_field(self):
        mode = self._get_current_field_mode()
        tv = self.field_tvs[mode]
        sel = tv.selection()
        if not sel:
            return
        idx = tv.index(sel[0])
        flist = sorted(self.fields[mode], key=lambda x: x['col'])
        if idx < len(flist):
            self.fields[mode].remove(flist[idx])
            self._refresh_fields()

    def _move_up(self):
        mode = self._get_current_field_mode()
        tv = self.field_tvs[mode]
        sel = tv.selection()
        if not sel:
            return
        idx = tv.index(sel[0])
        flist = sorted(self.fields[mode], key=lambda x: x['col'])
        if idx > 0:
            flist[idx]['col'], flist[idx-1]['col'] = flist[idx-1]['col'], flist[idx]['col']
            self._refresh_fields()

    def _move_down(self):
        mode = self._get_current_field_mode()
        tv = self.field_tvs[mode]
        sel = tv.selection()
        if not sel:
            return
        idx = tv.index(sel[0])
        flist = sorted(self.fields[mode], key=lambda x: x['col'])
        if idx < len(flist) - 1:
            flist[idx]['col'], flist[idx+1]['col'] = flist[idx+1]['col'], flist[idx]['col']
            self._refresh_fields()

    def _reset_fields(self):
        mode = self._get_current_field_mode()
        self.fields[mode] = [dict(f) for f in DEFAULT_FIELDS[mode]]
        self._refresh_fields()

    # ── EXPORT ────────────────────────────────────────────────
    def _export(self):
        total = (len(self.parsed['chat']) + len(self.parsed['combat']) +
                 len(self.parsed['loot']))
        if total == 0:
            messagebox.showwarning("No Data", "Nothing parsed yet. Run Parse first.")
            return

        # Auto-generate filename based on what's being exported
        fname = self.fname_var.get().strip()
        if not fname:
            # Extract character name from ACTION log files
            char_name = None
            for f in self.files:
                name = f['name'].upper()
                if 'ACTION' in name:
                    # Pattern: ACTION-PDF-CHARACTERNAME-timestamp.log
                    parts = f['name'].split('-')
                    if len(parts) >= 3:
                        char_name = parts[2].title()  # Capitalize properly
                        break
            
            # Generate filename based on what's being parsed
            if self.do_combat.get() and len(self.parsed['combat']) > 0:
                # Combat export: CharacterName_Combat_timestamp
                from datetime import datetime
                timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                fname = f"{char_name}_Combat_{timestamp}" if char_name else f"Combat_{timestamp}"
            elif self.do_loot.get() and len(self.parsed['loot']) > 0:
                # Loot export: CharacterName_Loot_timestamp
                from datetime import datetime
                timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                fname = f"{char_name}_Loot_{timestamp}" if char_name else f"Loot_{timestamp}"
            elif self.do_chat.get() and len(self.parsed['chat']) > 0:
                # Chat export: CharacterName_Chat_timestamp
                from datetime import datetime
                timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                fname = f"{char_name}_Chat_{timestamp}" if char_name else f"Chat_{timestamp}"
            else:
                fname = 'game_log_export'
        
        path  = filedialog.asksaveasfilename(
            defaultextension='.xlsx',
            initialdir=self.last_save_dir,
            initialfile=fname + '.xlsx',
            filetypes=[('Excel', '*.xlsx'), ('All', '*.*')])
        if not path:
            return
        
        # Save the directory for next time
        self.last_save_dir = os.path.dirname(path)

        try:
            # Handle master database append if enabled
            if self.use_master.get() and self.master_path_var.get() and self.parsed['loot']:
                self._append_to_master()
            
            # Check if we should export Combat and Loot as separate files
            if (self.separate_action_exports.get() and 
                self.do_combat.get() and self.do_loot.get() and
                len(self.parsed['combat']) > 0 and len(self.parsed['loot']) > 0):
                # Export as two separate files
                self._export_separate_action_files(path)
                return
            
            # Normal single-file export
            wb = openpyxl.Workbook()
            wb.remove(wb.active)

            if self.exp_summary.get():
                ws = wb.create_sheet("Summary")
                self._write_summary(ws)

            # Export chat - create a separate sheet for each chat file
            if self.do_chat.get() and self.parsed['chat_files']:
                for filename, chat_data in self.parsed['chat_files'].items():
                    # Create sheet name from filename (remove extension, truncate if needed)
                    sheet_name = os.path.splitext(filename)[0]
                    # Excel sheet names max 31 chars
                    if len(sheet_name) > 31:
                        sheet_name = sheet_name[:28] + "..."
                    ws = wb.create_sheet(sheet_name)
                    write_sheet(ws, chat_data, self.fields['chat'], 'chat')

            # Export combat if checkbox is checked and data exists
            if self.do_combat.get() and self.parsed['combat']:
                ws = wb.create_sheet("Combat")
                write_sheet(ws, self.parsed['combat'], self.fields['combat'], 'combat')

            # Export loot if checkbox is checked and data exists
            if self.do_loot.get() and self.parsed['loot']:
                ws = wb.create_sheet("Loot")
                write_sheet(ws, self.parsed['loot'], self.fields['loot'], 'loot')

            wb.save(path)
            self.export_status.config(
                text=f"✓ Saved to {os.path.basename(path)}")
            messagebox.showinfo("Exported",
                f"File saved!\n{path}\n\nSheets written:\n"
                f"  Chat:   {len(self.parsed['chat'])} rows\n"
                f"  Combat: {len(self.parsed['combat'])} rows\n"
                f"  Loot:   {len(self.parsed['loot'])} rows")
        except Exception as e:
            messagebox.showerror("Export Error", str(e))

    def _export_separate_action_files(self, base_path):
        """Export Combat and Loot as two separate files"""
        from datetime import datetime
        
        # Get character name for filenames
        char_name = None
        for f in self.files:
            name = f['name'].upper()
            if 'ACTION' in name:
                parts = f['name'].split('-')
                if len(parts) >= 3:
                    char_name = parts[2].title()
                    break
        
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        base_dir = os.path.dirname(base_path)
        
        # Create Combat file
        combat_filename = f"{char_name}_Combat_{timestamp}.xlsx" if char_name else f"Combat_{timestamp}.xlsx"
        combat_path = os.path.join(base_dir, combat_filename)
        
        wb_combat = openpyxl.Workbook()
        wb_combat.remove(wb_combat.active)
        if self.exp_summary.get():
            ws = wb_combat.create_sheet("Summary")
            self._write_summary(ws)
        ws_combat = wb_combat.create_sheet("Combat")
        write_sheet(ws_combat, self.parsed['combat'], self.fields['combat'], 'combat')
        wb_combat.save(combat_path)
        
        # Create Loot file
        loot_filename = f"{char_name}_Loot_{timestamp}.xlsx" if char_name else f"Loot_{timestamp}.xlsx"
        loot_path = os.path.join(base_dir, loot_filename)
        
        wb_loot = openpyxl.Workbook()
        wb_loot.remove(wb_loot.active)
        if self.exp_summary.get():
            ws = wb_loot.create_sheet("Summary")
            self._write_summary(ws)
        ws_loot = wb_loot.create_sheet("Loot")
        write_sheet(ws_loot, self.parsed['loot'], self.fields['loot'], 'loot')
        wb_loot.save(loot_path)
        
        self.export_status.config(text=f"✓ Saved 2 files: {combat_filename}, {loot_filename}")
        messagebox.showinfo("Exported",
            f"Files saved as separate exports!\n\n"
            f"Combat: {combat_path}\n"
            f"  {len(self.parsed['combat'])} rows\n\n"
            f"Loot: {loot_path}\n"
            f"  {len(self.parsed['loot'])} rows")
    
    def _write_summary(self, ws):
        ws.title = "Summary"
        now = datetime.now().strftime('%Y-%m-%d %H:%M')
        hdr = Font(name='Arial', bold=True, size=11)
        ws['A1'] = "Gaming Log Parser — Export Summary"
        ws['A1'].font = Font(name='Arial', bold=True, size=13)
        ws['A3'] = "Generated:";  ws['B3'] = now
        ws['A4'] = "Files processed:"; ws['B4'] = len(self.files)
        ws['A5'] = "Chat rows:";    ws['B5'] = len(self.parsed['chat'])
        ws['A6'] = "Combat rows:";  ws['B6'] = len(self.parsed['combat'])
        ws['A7'] = "Loot rows:";    ws['B7'] = len(self.parsed['loot'])
        ws['A9'] = "Files"
        ws['A9'].font = hdr
        ws['B9'] = "Type"
        ws['B9'].font = hdr
        for i, f in enumerate(self.files, 10):
            ws.cell(row=i, column=1).value = f['name']
            ws.cell(row=i, column=2).value = f['type']
        ws.column_dimensions['A'].width = 44
        ws.column_dimensions['B'].width = 16
    
    # ── BUILD TAB ─────────────────────────────────────────────
    def _build_build_tab(self):
        # The Build tab has more controls than fit in one screen (armor,
        # weapon, realm constraints, results...), so it's wrapped in a
        # scrollable canvas rather than clipping anything below Damage Type.
        outer = self.tab_build
        build_canvas = tk.Canvas(outer, highlightthickness=0)
        build_vsb = ttk.Scrollbar(outer, orient='vertical', command=build_canvas.yview)
        build_canvas.configure(yscrollcommand=build_vsb.set)
        build_canvas.pack(side='left', fill='both', expand=True)
        build_vsb.pack(side='right', fill='y')

        t = ttk.Frame(build_canvas)
        build_canvas_window = build_canvas.create_window((0, 0), window=t, anchor='nw')

        def _on_build_frame_configure(event):
            build_canvas.configure(scrollregion=build_canvas.bbox('all'))
        t.bind('<Configure>', _on_build_frame_configure)

        def _on_build_canvas_configure(event):
            build_canvas.itemconfig(build_canvas_window, width=event.width)
        build_canvas.bind('<Configure>', _on_build_canvas_configure)

        def _on_build_mousewheel(event):
            build_canvas.yview_scroll(int(-1 * (event.delta / 120)), 'units')
        build_canvas.bind('<Enter>', lambda e: build_canvas.bind_all('<MouseWheel>', _on_build_mousewheel))
        build_canvas.bind('<Leave>', lambda e: build_canvas.unbind_all('<MouseWheel>'))

        ttk.Label(t, text="Build Creator",
                  font=('Arial', 13, 'bold')).pack(anchor='w', pady=(0,10))
        
        # Master file selection
        file_frame = ttk.LabelFrame(t, text="Master Database File", padding=10)
        file_frame.pack(fill='x', pady=(0,12))
        
        ttk.Label(file_frame, text="File:").pack(side='left', padx=(0,8))
        self.search_master_path = tk.StringVar(value='')
        ttk.Entry(file_frame, textvariable=self.search_master_path, 
                 width=40, state='readonly').pack(side='left', padx=4)
        ttk.Button(file_frame, text="Browse...", 
                  command=self._browse_search_master).pack(side='left', padx=2)
        ttk.Button(file_frame, text="Create New", 
                  command=self._create_new_search_master).pack(side='left', padx=2)
        ttk.Button(file_frame, text="Use Community List", 
                  command=self._use_community_list).pack(side='left', padx=2)
        ttk.Button(file_frame, text="Load", 
                  command=self._load_master_for_search).pack(side='left', padx=2)
        
        # Build constraints
        constraints_frame = ttk.LabelFrame(t, text="Build Constraints", padding=10)
        constraints_frame.pack(fill='x', pady=(0,12))
        
        # Desired spells (left) and Realm filter (right, in the space next to
        # the spell dropdowns) share a row so the realm checkboxes aren't stuck
        # way down below the weapon constraints in unused whitespace.
        spell_and_realm_frame = ttk.Frame(constraints_frame)
        spell_and_realm_frame.pack(fill='x')

        spell_block = ttk.Frame(spell_and_realm_frame)
        spell_block.pack(side='left', anchor='n')

        # Desired spells - one dropdown row per spell category, each paired with a tier dropdown
        self.category_spell_vars = {}
        self.category_tier_vars = {}
        self.category_spell_combos = {}
        self.category_tier_combos = {}

        for category, spells in SPELL_CATEGORIES.items():
            spell_input_frame = ttk.Frame(spell_block)
            spell_input_frame.pack(fill='x', pady=2)
            ttk.Label(spell_input_frame, text=f"{category}:", width=12).pack(side='left')

            spell_var = tk.StringVar(value='')
            self.category_spell_vars[category] = spell_var
            spell_combo = ttk.Combobox(spell_input_frame, textvariable=spell_var,
                                       values=spells, state='readonly', width=22)
            spell_combo.pack(side='left', padx=4)
            self.category_spell_combos[category] = spell_combo

            tier_var = tk.StringVar(value=SPELL_TIERS[0])
            self.category_tier_vars[category] = tier_var
            tier_combo = ttk.Combobox(spell_input_frame, textvariable=tier_var,
                        values=SPELL_TIERS, state='readonly', width=6)
            tier_combo.pack(side='left', padx=4)
            self.category_tier_combos[category] = tier_combo

            spell_var.trace_add('write', lambda *args, c=category: self._update_tier_options(c))

            ttk.Button(spell_input_frame, text="Add to List",
                      command=lambda c=category: self._add_categorized_spell(c)).pack(side='left', padx=4)

            if not spells:
                ttk.Label(spell_input_frame, text="(list coming soon)",
                         font=('Arial', 8, 'italic'), foreground='#888').pack(side='left', padx=4)
            elif category in SPELL_CATEGORY_PARTIAL_NOTE:
                ttk.Label(spell_input_frame, text=SPELL_CATEGORY_PARTIAL_NOTE[category],
                         font=('Arial', 8, 'italic'), foreground='#888').pack(side='left', padx=4)

        # TEMPORARY manual entry - covers spells not yet in a category dropdown
        # (mainly Class Specific / Other1 / Other2). Remove this row once those
        # lists are fully compiled and every spell has a proper dropdown home.
        manual_frame = ttk.Frame(spell_block)
        manual_frame.pack(fill='x', pady=2)
        ttk.Label(manual_frame, text="Manual:", width=12).pack(side='left')
        self.manual_spell_var = tk.StringVar(value='')
        ttk.Entry(manual_frame, textvariable=self.manual_spell_var, width=22).pack(side='left', padx=4)
        ttk.Button(manual_frame, text="Add to List",
                  command=self._add_manual_spell).pack(side='left', padx=4)
        ttk.Label(manual_frame, text="(temporary, until Class Specific/Other1/Other2 lists are compiled)",
                 font=('Arial', 8, 'italic'), foreground='#a33').pack(side='left', padx=4)

        # Wanted Spells (narrower, left) and Required Items (right, next to it)
        # share a row instead of each spanning the full width stacked vertically.
        wanted_and_required_frame = ttk.Frame(constraints_frame)
        wanted_and_required_frame.pack(fill='x', pady=4)

        # Wanted Spells - chips flow left-to-right and wrap to new lines. This
        # block expands to take the extra row width, making it the wider of
        # the two (Required Items stays a fixed, narrower width beside it).
        wanted_block = ttk.Frame(wanted_and_required_frame)
        wanted_block.pack(side='left', anchor='n', fill='both', expand=True)
        ttk.Label(wanted_block, text="Wanted Spells:").pack(anchor='w')

        spell_scroll_frame = ttk.Frame(wanted_block)
        spell_scroll_frame.pack(fill='both', expand=True)

        self.wanted_spells_data = []
        self.spell_chips_text = tk.Text(spell_scroll_frame, height=4, width=45, wrap='word',
                                        cursor='arrow', state='disabled')
        spell_scroll = ttk.Scrollbar(spell_scroll_frame, orient='vertical',
                                    command=self.spell_chips_text.yview)
        self.spell_chips_text.configure(yscrollcommand=spell_scroll.set)
        self.spell_chips_text.pack(side='left', fill='both', expand=True)
        spell_scroll.pack(side='right', fill='y')

        ttk.Button(wanted_block, text="Clear All",
                  command=self._clear_spell_list).pack(anchor='w', pady=(2,0))

        # Required Items - force specific gear into the build, calculating the
        # rest of the build around it. Tolerates spelling errors via fuzzy match.
        # Label + chip box come first (matching Wanted Spells' layout so both
        # chip areas line up in the same row); the entry to add a new one sits
        # below the chips instead of above them.
        required_block = ttk.Frame(wanted_and_required_frame)
        required_block.pack(side='left', anchor='n', padx=(20,0))

        ttk.Label(required_block, text="Required Items:").pack(anchor='w')
        required_scroll_frame = ttk.Frame(required_block)
        required_scroll_frame.pack(fill='both', expand=True)

        self.required_items = []
        self.required_items_text = tk.Text(required_scroll_frame, height=4, width=25, wrap='word',
                                           cursor='arrow', state='disabled')
        required_scroll = ttk.Scrollbar(required_scroll_frame, orient='vertical',
                                       command=self.required_items_text.yview)
        self.required_items_text.configure(yscrollcommand=required_scroll.set)
        self.required_items_text.pack(side='left', fill='both', expand=True)
        required_scroll.pack(side='right', fill='y')

        required_input_frame = ttk.Frame(required_block)
        required_input_frame.pack(fill='x', pady=(4,0))
        ttk.Label(required_input_frame, text="Require Item:").pack(side='left')
        self.required_item_var = tk.StringVar(value='')
        required_entry = ttk.Entry(required_input_frame, textvariable=self.required_item_var, width=26)
        required_entry.pack(side='left', padx=4)
        required_entry.bind('<Return>', lambda e: self._add_required_item())
        ttk.Button(required_input_frame, text="Add to Build",
                  command=self._add_required_item).pack(side='left', padx=4)
        ttk.Button(required_input_frame, text="Clear All",
                  command=self._clear_required_items).pack(side='left', padx=4)

        ttk.Label(required_block, text="(e.g. a specific weapon/armor piece - typos are OK, it'll offer close matches)",
                 font=('Arial', 8, 'italic'), foreground='#666').pack(anchor='w', pady=(2,0))

        # Level filters (min/max/specific with mutual exclusion)
        level_frame = ttk.Frame(constraints_frame)
        level_frame.pack(fill='x', pady=(8,4))
        
        # Min Level
        ttk.Label(level_frame, text="Min Level:", width=12).pack(side='left')
        self.min_level_var = tk.StringVar(value='')
        self.min_level_entry = ttk.Entry(level_frame, textvariable=self.min_level_var, width=8)
        self.min_level_entry.pack(side='left', padx=(0,8))
        self.min_level_var.trace_add('write', lambda *args: self._update_level_fields())
        
        # Max Level
        ttk.Label(level_frame, text="Max Level:").pack(side='left', padx=(8,4))
        self.max_level_var = tk.StringVar(value='')
        self.max_level_entry = ttk.Entry(level_frame, textvariable=self.max_level_var, width=8)
        self.max_level_entry.pack(side='left', padx=(0,8))
        self.max_level_var.trace_add('write', lambda *args: self._update_level_fields())
        
        # Specific Level
        ttk.Label(level_frame, text="Specific Level:").pack(side='left', padx=(8,4))
        self.specific_level_var = tk.StringVar(value='')
        self.specific_level_entry = ttk.Entry(level_frame, textvariable=self.specific_level_var, width=8)
        self.specific_level_entry.pack(side='left', padx=(0,4))
        self.specific_level_var.trace_add('write', lambda *args: self._update_level_fields())
        
        ttk.Label(level_frame, text="(leave blank for no restriction)", 
                 font=('Arial', 8, 'italic'), foreground='#666').pack(side='left', padx=4)
        
        # Armor type constraints
        armor_header = ttk.Frame(constraints_frame)
        armor_header.pack(fill='x', pady=(8,4))
        ttk.Label(armor_header, text="Armor Type Constraints:", 
                 font=('Arial', 10, 'bold')).pack(side='left')
        ttk.Button(armor_header, text="Set as Default", 
                  command=self._save_armor_defaults).pack(side='left', padx=4)
        ttk.Button(armor_header, text="Clear Default", 
                  command=self._clear_armor_defaults).pack(side='left', padx=4)
        ttk.Button(armor_header, text="Clear All", 
                  command=self._clear_all_armor).pack(side='left', padx=4)
        
        # Create checkbox storage for armor types
        self.armor_checks = {}
        
        # "All" row - controls all slots at once
        all_frame = ttk.Frame(constraints_frame)
        all_frame.pack(fill='x', pady=2)
        ttk.Label(all_frame, text="All:", width=12, font=('Arial', 9, 'bold')).pack(side='left')
        
        self.armor_all_checks = {}
        for armor_type in ['cloth', 'leather', 'studded', 'plate']:
            var = tk.BooleanVar(value=False)
            self.armor_all_checks[armor_type] = var
            ttk.Checkbutton(all_frame, text=armor_type.title(), 
                           variable=var,
                           command=lambda t=armor_type: self._update_all_armor(t)).pack(side='left', padx=4)
        
        # Armor slots
        for slot, label in [('head', 'Head'), ('cloak', 'Cloak'), ('body', 'Body'), 
                           ('hands', 'Hands'), ('legs', 'Legs'), ('feet', 'Feet')]:
            slot_frame = ttk.Frame(constraints_frame)
            slot_frame.pack(fill='x', pady=2)
            ttk.Label(slot_frame, text=f"{label}:", width=12).pack(side='left')
            
            # Create checkboxes for each armor type
            self.armor_checks[slot] = {}
            for armor_type in ['cloth', 'leather', 'studded', 'plate']:
                var = tk.BooleanVar(value=False)
                self.armor_checks[slot][armor_type] = var
                ttk.Checkbutton(slot_frame, text=armor_type.title(), 
                               variable=var).pack(side='left', padx=4)
        
        # Load saved armor defaults
        self._load_armor_defaults()
        
        # Weapon constraints
        weapon_header = ttk.Frame(constraints_frame)
        weapon_header.pack(fill='x', pady=(8,4))
        ttk.Label(weapon_header, text="Weapon Constraints:", 
                 font=('Arial', 10, 'bold')).pack(side='left')
        
        # Create checkbox storage for weapon types
        self.weapon_checks = {}
        
        # Melee vs Direct (radio buttons)
        weapon_style_frame = ttk.Frame(constraints_frame)
        weapon_style_frame.pack(fill='x', pady=2)
        ttk.Label(weapon_style_frame, text="Weapon Style:", width=14).pack(side='left')
        self.weapon_style_var = tk.StringVar(value='')  # '' = Any
        ttk.Radiobutton(weapon_style_frame, text="Melee", 
                       variable=self.weapon_style_var, value='melee').pack(side='left', padx=4)
        ttk.Radiobutton(weapon_style_frame, text="Direct (Caster)", 
                       variable=self.weapon_style_var, value='direct-caster').pack(side='left', padx=4)
        
        # Parry Staff (formerly "Direct (Stat)") with note
        direct_stat_frame = ttk.Frame(weapon_style_frame)
        direct_stat_frame.pack(side='left')
        ttk.Radiobutton(direct_stat_frame, text="Parry Staff",
                       variable=self.weapon_style_var, value='direct-stat').pack(side='left', padx=4)
        ttk.Label(direct_stat_frame, text="[not yet implemented]",
                 foreground='#888', font=('Arial', 8)).pack(side='left', padx=2)
        
        ttk.Radiobutton(weapon_style_frame, text="Any", 
                       variable=self.weapon_style_var, value='').pack(side='left', padx=4)
        
        # Build configuration (checkboxes)
        weapon_config_frame = ttk.Frame(constraints_frame)
        weapon_config_frame.pack(fill='x', pady=2)
        ttk.Label(weapon_config_frame, text="Build Config:", width=14).pack(side='left')
        
        self.build_config_checks = {}
        config_items = [('weapon', 'Weapon'), ('shield', 'Shield'), ('two-handed', 'Two-Handed'),
                        ('claw_1', '1 Claw'), ('claw_2', '2 Claw')]
        for config_type, label in config_items:
            var = tk.BooleanVar(value=False)
            self.build_config_checks[config_type] = var
            ttk.Checkbutton(weapon_config_frame, text=label, variable=var,
                           command=self._update_weapon_config_options).pack(side='left', padx=4)
        
        # Dual-wield sub-option (indent under main checkboxes)
        dualwield_frame = ttk.Frame(constraints_frame)
        dualwield_frame.pack(fill='x', pady=2)
        ttk.Label(dualwield_frame, text="", width=14).pack(side='left')  # Spacer for alignment
        ttk.Label(dualwield_frame, text="↳", foreground='#888').pack(side='left', padx=(0,2))
        self.dual_wield_var = tk.BooleanVar(value=False)
        self.dual_wield_cb = ttk.Checkbutton(dualwield_frame, text="Dual-Wield (for Weapon/Two-Handed)", 
                                              variable=self.dual_wield_var)
        self.dual_wield_cb.pack(side='left')
        
        # Weapon type checkboxes (damage types)
        self.weapon_type_frame = ttk.Frame(constraints_frame)
        self.weapon_type_frame.pack(fill='x', pady=2)
        ttk.Label(self.weapon_type_frame, text="Damage Type:", width=14).pack(side='left')
        
        self.weapon_type_checkbuttons = {}
        for weapon_type in ['slashing', 'thrusting', 'crushing']:
            var = tk.BooleanVar(value=False)
            self.weapon_checks[weapon_type] = var
            cb = ttk.Checkbutton(self.weapon_type_frame, text=weapon_type.title(), 
                                variable=var)
            cb.pack(side='left', padx=4)
            self.weapon_type_checkbuttons[weapon_type] = cb
        
        # Note: claw and offhand removed since they're in build config now
        
        # Load weapon defaults
        self._load_weapon_defaults()
        
        # Realm filter (Only Found In) - placed in the empty space to the right
        # of the spell dropdowns, in the row frame set up alongside spell_block.
        realm_block = ttk.LabelFrame(spell_and_realm_frame, text="Only Found In", padding=8)
        realm_block.pack(side='left', anchor='n', padx=(20, 0))

        self.realm_filters = {}
        realm_options = ['Evil', 'Chaos', 'Good', 'Kaid', 'Crafted', 'Glory Bea', 'Event']
        cols = 2
        for i, realm in enumerate(realm_options):
            var = tk.BooleanVar(value=False)
            self.realm_filters[realm] = var
            ttk.Checkbutton(realm_block, text=realm, variable=var).grid(
                row=i // cols, column=i % cols, sticky='w', padx=4, pady=2)
        
        # Search buttons
        button_frame = ttk.Frame(constraints_frame)
        button_frame.pack(pady=8)

        ttk.Button(button_frame, text="🎯 Find Optimal Build",
                  command=self._find_optimal_build, width=20).pack(side='left', padx=4)
        ttk.Button(button_frame, text="📋 Show All Matches",
                  command=self._show_all_matches, width=20).pack(side='left', padx=4)

        self.generate_multi_builds_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(button_frame, text="🎲 Generate multiple build options",
                       variable=self.generate_multi_builds_var).pack(side='left', padx=(12,4))
        
        # Status
        self.search_status = ttk.Label(t, text="Load a master database file and add desired spells to begin", 
                                      foreground='#666')
        self.search_status.pack(anchor='w', pady=4, padx=4)
        
        # Store loaded data
        self.master_data = []
    
    # ── RESULTS TAB ───────────────────────────────────────────────
    def _build_results_tab(self):
        t = self.tab_results
        
        # Header with display mode and export button
        header_frame = ttk.Frame(t)
        header_frame.pack(fill='x', pady=(0,10))
        
        ttk.Label(header_frame, text="Build Search Results",
                 font=('Arial', 13, 'bold')).pack(side='left')
        
        # Display mode radio buttons (center)
        mode_frame = ttk.Frame(header_frame)
        mode_frame.pack(side='left', padx=40)
        ttk.Label(mode_frame, text="Display:").pack(side='left', padx=(0,8))
        self.results_display_mode = tk.StringVar(value='optimal')
        ttk.Radiobutton(mode_frame, text="🎯 Best Per Slot", 
                       variable=self.results_display_mode, value='optimal',
                       command=self._refresh_results_display).pack(side='left', padx=4)
        ttk.Radiobutton(mode_frame, text="📋 All Matches",
                       variable=self.results_display_mode, value='all',
                       command=self._refresh_results_display).pack(side='left', padx=4)

        # Build variant selector - only meaningful when "Generate multiple build
        # options" produced more than one alternate full build.
        ttk.Label(mode_frame, text="Build Variant:").pack(side='left', padx=(16,4))
        self.build_variant_var = tk.StringVar(value='Build 1')
        self.build_variant_combo = ttk.Combobox(mode_frame, textvariable=self.build_variant_var,
                                                values=['Build 1'], state='disabled', width=12)
        self.build_variant_combo.pack(side='left', padx=4)
        self.build_variant_combo.bind('<<ComboboxSelected>>', self._switch_build_variant)
        
        ttk.Button(header_frame, text="📤 Export Results", 
                  command=self._export_build_results).pack(side='right')
        
        # Results treeview
        results_frame = ttk.LabelFrame(t, text="Suggested Build", padding=10)
        results_frame.pack(fill='both', expand=True)
        
        cols = ('Slot', 'Item', 'Type', 'Spell', 'Level', 'Mob', 'Area', 'Div', 'Alt Options')
        col_widths = {'Slot': 55, 'Item': 200, 'Type': 60, 'Spell': 100, 'Level': 45,
                     'Mob': 120, 'Area': 120, 'Alt Options': 280}
        self.search_results_tv = ttk.Treeview(results_frame, columns=cols,
                                             show='headings', height=20)
        for col in cols:
            if col == 'Div':
                self.search_results_tv.heading(col, text='')
                self.search_results_tv.column(col, width=14, stretch=False, anchor='center')
            else:
                self.search_results_tv.heading(col, text=col)
                self.search_results_tv.column(col, width=col_widths[col])
        
        scroll = ttk.Scrollbar(results_frame, orient='vertical', 
                              command=self.search_results_tv.yview)
        self.search_results_tv.configure(yscrollcommand=scroll.set)
        
        self.search_results_tv.pack(side='left', fill='both', expand=True)
        scroll.pack(side='right', fill='y')
        
        # Filter controls at bottom
        filter_frame = ttk.Frame(t)
        filter_frame.pack(fill='x', pady=(8,0))
        
        # Remove by Area filter
        ttk.Label(filter_frame, text="Remove Area:").pack(side='left', padx=(0,4))
        self.remove_area_var = tk.StringVar(value='')
        ttk.Entry(filter_frame, textvariable=self.remove_area_var, width=20).pack(side='left', padx=4)
        ttk.Button(filter_frame, text="Remove Items", 
                  command=self._remove_items_by_area).pack(side='left', padx=4)
        ttk.Label(filter_frame, text="(e.g. 'Easter 2023', 'Olympics')", 
                 font=('Arial', 8, 'italic'), foreground='#666').pack(side='left', padx=4)
        
        # Store the last search results for toggling
        self.last_optimal_results = []
        self.last_all_results = []
        self.optimal_build = {}
        self.slot_alternates = {}
        self.build_variants = []
    
    def _remove_items_by_area(self):
        """Remove items from results by area name"""
        area_filter = self.remove_area_var.get().strip().lower()
        if not area_filter:
            messagebox.showwarning("No Area", "Please enter an area name to remove")
            return
        
        # Remove from both stored results
        self.last_optimal_results = [
            row for row in self.last_optimal_results 
            if area_filter not in (row[6] or '').lower()  # row[6] is Area column
        ]
        self.last_all_results = [
            row for row in self.last_all_results 
            if area_filter not in (row[6] or '').lower()
        ]
        
        # Refresh display
        self._refresh_results_display()
        
        # Clear the input
        self.remove_area_var.set('')
    
    def _apply_realm_filter(self):
        """Filter results by selected realms - applied during search"""
        # Realm filtering is now integrated into search functions
        # This function is just a placeholder for the checkbox command
        pass
    
    def _update_tier_options(self, category):
        """Restrict the tier dropdown based on the selected spell (e.g. Agility has no tier iii)"""
        spell = self.category_spell_vars[category].get().strip().lower()
        allowed = SPELL_TIER_RESTRICTIONS.get(spell, SPELL_TIERS)
        combo = self.category_tier_combos[category]
        combo['values'] = allowed
        if self.category_tier_vars[category].get() not in allowed:
            self.category_tier_vars[category].set(allowed[0])

    def _add_categorized_spell(self, category):
        """Add the spell selected in a category dropdown (plus tier) to the wanted spells list"""
        spell = self.category_spell_vars[category].get().strip()
        if not spell:
            return

        combined = SPELL_VALUE_OVERRIDES.get(spell.lower(), spell.lower())
        tier = self.category_tier_vars[category].get()
        if tier and tier != '(any)':
            combined = f"{combined}.{tier}"

        self._append_wanted_spell(combined)

    def _add_manual_spell(self):
        """TEMPORARY free-text add - remove once Class Specific/Other1/Other2 dropdowns are populated"""
        spell = self.manual_spell_var.get().strip()
        if not spell:
            return
        self._append_wanted_spell(spell.lower())
        self.manual_spell_var.set('')

    def _append_wanted_spell(self, spell_value):
        """Add a spell string to the wanted list (deduped) and re-render the chip display"""
        if spell_value not in [s.lower() for s in self.wanted_spells_data]:
            self.wanted_spells_data.append(spell_value)
            self._render_spell_chips()

    def _remove_wanted_spell(self, spell_value):
        """Remove one spell chip (called by a chip's own ✕ button)"""
        if spell_value in self.wanted_spells_data:
            self.wanted_spells_data.remove(spell_value)
            self._render_spell_chips()

    def _clear_spell_list(self):
        """Clear all spells from list"""
        self.wanted_spells_data = []
        self._render_spell_chips()

    def _render_spell_chips(self):
        """Redraw the Wanted Spells area as chips that flow horizontally and
        wrap to new lines, instead of one spell per vertical row."""
        text = self.spell_chips_text
        text.config(state='normal')
        text.delete('1.0', tk.END)
        for spell in self.wanted_spells_data:
            chip = ttk.Frame(text, relief='raised', borderwidth=1)
            ttk.Label(chip, text=spell, padding=(4, 1)).pack(side='left')
            remove_lbl = ttk.Label(chip, text='✕', padding=(4, 1),
                                   foreground='#a33', cursor='hand2')
            remove_lbl.pack(side='left')
            remove_lbl.bind('<Button-1>', lambda e, s=spell: self._remove_wanted_spell(s))
            text.window_create(tk.END, window=chip)
            text.insert(tk.END, ' ')
        text.config(state='disabled')

    def _add_required_item(self):
        """Look up a specific item the user typed in the master database and
        force it into the build. Falls back to fuzzy matching (tolerating
        spelling errors) when there's no exact (case-insensitive) match."""
        query = self.required_item_var.get().strip()
        if not query:
            return
        if not self.master_data:
            messagebox.showwarning("No Data", "Load a master database first.")
            return

        exact = [it for it in self.master_data
                if (it.get('Item') or '').strip().lower() == query.lower()]
        if exact:
            self._confirm_and_add_required_item(exact[0])
            return

        all_names = sorted({(it.get('Item') or '').strip()
                            for it in self.master_data if it.get('Item')})
        matches = difflib.get_close_matches(query, all_names, n=5, cutoff=0.5)
        if not matches:
            messagebox.showinfo("No Match Found",
                f"No item found matching '{query}' (even allowing for spelling errors).")
            return

        if len(matches) == 1:
            if messagebox.askyesno("Did You Mean?",
                    f"No exact match for '{query}'.\n\nDid you mean '{matches[0]}'?"):
                item = next(it for it in self.master_data
                           if (it.get('Item') or '').strip() == matches[0])
                self._confirm_and_add_required_item(item)
            return

        dlg = ItemMatchDialog(self, query, matches)
        self.wait_window(dlg)
        if dlg.result:
            item = next(it for it in self.master_data
                       if (it.get('Item') or '').strip() == dlg.result)
            self._confirm_and_add_required_item(item)

    def _confirm_and_add_required_item(self, item):
        """Add a resolved item (exact or fuzzy-confirmed) to the required list"""
        name = (item.get('Item') or '').strip()
        if any((it.get('Item') or '').strip().lower() == name.lower() for it in self.required_items):
            messagebox.showinfo("Already Added", f"'{name}' is already in your required items list.")
            return
        self.required_items.append(item)
        self.required_item_var.set('')
        self._render_required_item_chips()

    def _remove_required_item(self, item):
        """Remove one required-item chip (called by a chip's own ✕ button)"""
        if item in self.required_items:
            self.required_items.remove(item)
            self._render_required_item_chips()

    def _clear_required_items(self):
        """Clear all required items"""
        self.required_items = []
        self._render_required_item_chips()

    def _render_required_item_chips(self):
        """Redraw the Required Items area as removable chips showing item + slot"""
        text = self.required_items_text
        text.config(state='normal')
        text.delete('1.0', tk.END)
        for item in self.required_items:
            label = f"{item.get('Item', '')} ({(item.get('Slot') or '').title()})"
            chip = ttk.Frame(text, relief='raised', borderwidth=1)
            ttk.Label(chip, text=label, padding=(4, 1)).pack(side='left')
            remove_lbl = ttk.Label(chip, text='✕', padding=(4, 1),
                                   foreground='#a33', cursor='hand2')
            remove_lbl.pack(side='left')
            remove_lbl.bind('<Button-1>', lambda e, it=item: self._remove_required_item(it))
            text.window_create(tk.END, window=chip)
            text.insert(tk.END, ' ')
        text.config(state='disabled')

    def _browse_search_master(self):
        """Browse for master database to search"""
        path = filedialog.askopenfilename(
            title='Select Master Database to Search',
            initialdir=self.last_save_dir,
            filetypes=[('Excel', '*.xlsx'), ('All', '*.*')])
        if path:
            self.search_master_path.set(path)
    
    def _use_community_list(self):
        """Use the bundled community equipment list as master database"""
        # Look for the community list in common locations
        possible_paths = [
            # Same directory as the script
            os.path.join(os.path.dirname(__file__), 'Olmran_Community_Eq_and_Stats_List.xlsx'),
            # Current working directory
            'Olmran_Community_Eq_and_Stats_List.xlsx',
            # One level up
            os.path.join(os.path.dirname(__file__), '..', 'Olmran_Community_Eq_and_Stats_List.xlsx'),
        ]
        
        community_file = None
        for path in possible_paths:
            if os.path.exists(path):
                community_file = os.path.abspath(path)
                break
        
        if not community_file:
            messagebox.showerror("File Not Found", 
                "Could not find 'Olmran_Community_Eq_and_Stats_List.xlsx'\n\n"
                "Please ensure the community equipment list is in the same folder as the program.")
            return
        
        # Set the path and auto-load
        self.search_master_path.set(community_file)
        self._load_master_for_search()
    
    def _create_new_search_master(self):
        """Create a new master database file from Build tab"""
        from datetime import datetime
        timestamp = datetime.now().strftime('%Y%m%d')
        default_name = f"Master_Loot_Database_{timestamp}"
        
        path = filedialog.asksaveasfilename(
            title='Create New Master Database',
            initialdir=self.last_save_dir,
            initialfile=default_name + '.xlsx',
            defaultextension='.xlsx',
            filetypes=[('Excel', '*.xlsx'), ('All', '*.*')])
        
        if not path:
            return
        
        try:
            # Create new Excel file with Loot sheet
            wb = openpyxl.Workbook()
            wb.remove(wb.active)  # Remove default sheet
            
            ws = wb.create_sheet("Loot")
            
            # Write header row using loot fields
            from openpyxl.styles import Font, PatternFill, Alignment
            
            headers = []
            max_col = max(f['col'] for f in self.fields['loot'])
            headers = [''] * max_col
            for f in self.fields['loot']:
                headers[f['col'] - 1] = f['label']
            
            ws.append(headers)
            
            # Style header
            hdr_font = Font(name='Arial', size=11, bold=True, color='FFFFFF')
            hdr_fill = PatternFill(fill_type='solid', start_color='4472C4', end_color='4472C4')
            for col_idx, header in enumerate(headers, 1):
                cell = ws.cell(1, col_idx)
                cell.font = hdr_font
                cell.fill = hdr_fill
                cell.alignment = Alignment(horizontal='center', vertical='center')
            
            ws.freeze_panes = 'A2'
            
            # Save
            wb.save(path)
            
            # Set as master file and load it
            self.search_master_path.set(path)
            self.last_save_dir = os.path.dirname(path)
            
            messagebox.showinfo("Master Created", 
                f"✓ New master database created!\n\n{path}\n\n"
                "Click 'Load' to load it for searching.")
                
        except Exception as e:
            messagebox.showerror("Create Error", f"Failed to create master file:\n{str(e)}")
    
    def _load_master_for_search(self):
        """Load master database into memory for searching"""
        path = self.search_master_path.get()
        if not path or not os.path.exists(path):
            messagebox.showwarning("No File", "Please select a master database file first")
            return
        
        try:
            wb = openpyxl.load_workbook(path, read_only=True)
            
            # Check for either 'Loot' or 'Equipment' sheet
            sheet_name = None
            if 'Loot' in wb.sheetnames:
                sheet_name = 'Loot'
            elif 'Equipment' in wb.sheetnames:
                sheet_name = 'Equipment'
            else:
                messagebox.showerror("No Equipment Sheet", 
                    "Selected file has no 'Loot' or 'Equipment' sheet")
                return
            
            ws = wb[sheet_name]
            
            # Read all data
            self.master_data = []
            headers = [cell.value for cell in ws[1]]
            
            for row in ws.iter_rows(min_row=2, values_only=True):
                item_dict = {}
                for idx, header in enumerate(headers):
                    if header and idx < len(row):
                        item_dict[header] = row[idx] or ''
                if item_dict.get('Item'):  # Only add if has item name
                    self.master_data.append(item_dict)
            
            wb.close()
            
            self.search_status.config(
                text=f"✓ Loaded {len(self.master_data)} items from {os.path.basename(path)}")
            messagebox.showinfo("Loaded", 
                f"Master database loaded!\n{len(self.master_data)} items ready to search")
            
        except Exception as e:
            messagebox.showerror("Load Error", f"Failed to load master:\n{str(e)}")
    
    def _clear_armor_defaults(self):
        """Clear saved defaults from config file"""
        self.armor_defaults = {}
        self.weapon_defaults = {}
        self.weapon_style_default = ''
        self.build_config_defaults = {}
        self.dual_wield_default = False
        self._save_config()
        messagebox.showinfo("Defaults Cleared", "Saved armor and weapon defaults have been cleared!")
    
    def _clear_all_armor(self):
        """Clear all armor and weapon checkboxes"""
        # Clear All row
        for armor_type in ['cloth', 'leather', 'studded', 'plate']:
            self.armor_all_checks[armor_type].set(False)
        
        # Clear all individual armor slots
        for slot in ['head', 'cloak', 'body', 'hands', 'legs', 'feet']:
            for armor_type in ['cloth', 'leather', 'studded', 'plate']:
                self.armor_checks[slot][armor_type].set(False)
        
        # Clear weapon damage types
        for weapon_type in ['slashing', 'thrusting', 'crushing']:
            self.weapon_checks[weapon_type].set(False)
        
        # Clear weapon style
        self.weapon_style_var.set('')
        
        # Clear build config checkboxes
        for config_type in ['weapon', 'shield', 'two-handed', 'claw_1', 'claw_2']:
            self.build_config_checks[config_type].set(False)
        
        # Clear dual-wield
        self.dual_wield_var.set(False)
        
        # Clear all level filters
        self.min_level_var.set('')
        self.max_level_var.set('')
        self.specific_level_var.set('')
    
    def _update_all_armor(self, armor_type):
        """When an 'All' checkbox is clicked, update all individual slots to match"""
        all_value = self.armor_all_checks[armor_type].get()
        for slot in ['head', 'cloak', 'body', 'hands', 'legs', 'feet']:
            self.armor_checks[slot][armor_type].set(all_value)
        # Note: After this, individual slots can be changed independently
        # The "All" checkbox doesn't stay synced - it's just a quick-set tool
    
    def _refresh_results_display(self):
        """Toggle between optimal and all matches view"""
        mode = self.results_display_mode.get()
        
        # Clear current display
        self.search_results_tv.delete(*self.search_results_tv.get_children())
        
        # Show appropriate results
        if mode == 'optimal':
            results = self.last_optimal_results
        else:
            results = self.last_all_results
        
        # Display results
        for row in results:
            self.search_results_tv.insert('', 'end', values=row)

        self._autosize_results_columns()

    def _autosize_results_columns(self):
        """Shrink/grow each results column to the minimum width that fits its
        header and currently displayed cell values - no wasted or cramped space."""
        tv = self.search_results_tv
        cols = tv['columns']
        font = tkfont.Font(font=ttk.Style().lookup('Treeview', 'font') or 'TkDefaultFont')
        padding = 24  # cell borders/inset allowance

        for idx, col in enumerate(cols):
            if col == 'Div':
                continue  # fixed-width visual spacer, not content-driven
            max_width = font.measure(tv.heading(col)['text'])
            for iid in tv.get_children():
                value = tv.item(iid)['values'][idx]
                w = font.measure(str(value))
                if w > max_width:
                    max_width = w
            tv.column(col, width=max_width + padding)

    def _build_dict_to_rows(self, build_dict):
        """Turn a {slot: item} build mapping into result-table rows, listing
        each slot's other tied candidates (from self.optimal_build / self.slot_alternates)
        as its Alt Options rather than the item actually shown in this variant."""
        slot_order = ['head', 'jewel_1', 'jewel_2', 'cloak', 'body', 'hands', 'legs', 'feet',
                     'weapon', 'shield', 'claw_1', 'claw_2']
        rows = []
        for slot in slot_order:
            if slot not in build_dict:
                continue
            item = build_dict[slot]
            display_slot = 'jewel' if slot.startswith('jewel') else 'claw' if slot.startswith('claw') else slot
            candidates = [self.optimal_build.get(slot)] + self.slot_alternates.get(slot, [])
            alt_names = [f"{c.get('Item', '')} ({c.get('Spell', '')})" for c in candidates if c and c is not item]
            alt_text = ', '.join(alt_names) if alt_names else '(none)'
            rows.append((
                display_slot.title(),
                item.get('Item', ''),
                item.get('Type', ''),
                item.get('Spell', ''),
                item.get('Level', ''),
                item.get('Mob', ''),
                item.get('Area', ''),
                '████████',
                alt_text
            ))
        return rows

    def _switch_build_variant(self, event=None):
        """Show a different generated build variant in the results table"""
        label = self.build_variant_var.get()
        try:
            idx = int(label.split()[-1]) - 1
        except (ValueError, IndexError):
            idx = 0
        if not (0 <= idx < len(self.build_variants)):
            return

        self.last_optimal_results = self._build_dict_to_rows(self.build_variants[idx])
        self.results_display_mode.set('optimal')
        self._refresh_results_display()

    def _update_level_fields(self):
        """Enable/disable level fields based on which one is being used"""
        has_min = bool(self.min_level_var.get().strip())
        has_max = bool(self.max_level_var.get().strip())
        has_specific = bool(self.specific_level_var.get().strip())
        
        # If specific level is set, disable min/max
        if has_specific:
            self.min_level_entry.config(state='disabled')
            self.max_level_entry.config(state='disabled')
        # If min or max is set, disable specific
        elif has_min or has_max:
            self.specific_level_entry.config(state='disabled')
        # If nothing is set, enable all
        else:
            self.min_level_entry.config(state='normal')
            self.max_level_entry.config(state='normal')
            self.specific_level_entry.config(state='normal')
    
    def _update_weapon_config_options(self):
        """Update UI based on build config selections"""
        # Currently no special logic needed - all damage types are always available
        pass
    
    def _save_armor_defaults(self):
        """Save current armor and weapon selections as defaults"""
        self.armor_defaults = {}
        for slot in ['head', 'cloak', 'body', 'hands', 'legs', 'feet']:
            self.armor_defaults[slot] = {}
            for armor_type in ['cloth', 'leather', 'studded', 'plate']:
                self.armor_defaults[slot][armor_type] = self.armor_checks[slot][armor_type].get()
        
        # Save weapon defaults
        self.weapon_defaults = {}
        for weapon_type in ['slashing', 'thrusting', 'crushing']:
            self.weapon_defaults[weapon_type] = self.weapon_checks[weapon_type].get()
        
        # Save weapon style default
        self.weapon_style_default = self.weapon_style_var.get()
        
        # Save build config defaults (checkboxes)
        self.build_config_defaults = {}
        for config_type in ['weapon', 'shield', 'two-handed', 'claw_1', 'claw_2']:
            self.build_config_defaults[config_type] = self.build_config_checks[config_type].get()
        
        # Save dual-wield default
        self.dual_wield_default = self.dual_wield_var.get()
        
        self._save_config()
        messagebox.showinfo("Defaults Saved", "Armor and weapon preferences saved as defaults!")
    
    def _load_armor_defaults(self):
        """Load saved armor defaults into checkboxes"""
        if hasattr(self, 'armor_defaults') and self.armor_defaults:
            for slot in ['head', 'cloak', 'body', 'hands', 'legs', 'feet']:
                if slot in self.armor_defaults:
                    for armor_type in ['cloth', 'leather', 'studded', 'plate']:
                        if armor_type in self.armor_defaults[slot]:
                            self.armor_checks[slot][armor_type].set(
                                self.armor_defaults[slot][armor_type])
    
    def _load_weapon_defaults(self):
        """Load saved weapon defaults into checkboxes"""
        if hasattr(self, 'weapon_defaults') and self.weapon_defaults:
            for weapon_type in ['slashing', 'thrusting', 'crushing']:
                if weapon_type in self.weapon_defaults:
                    self.weapon_checks[weapon_type].set(self.weapon_defaults[weapon_type])
        
        # Load weapon style default
        if hasattr(self, 'weapon_style_default'):
            self.weapon_style_var.set(self.weapon_style_default)
        
        # Load build config defaults (checkboxes)
        if hasattr(self, 'build_config_defaults') and self.build_config_defaults:
            for config_type in ['weapon', 'shield', 'two-handed', 'claw_1', 'claw_2']:
                if config_type in self.build_config_defaults:
                    self.build_config_checks[config_type].set(self.build_config_defaults[config_type])
        
        # Load dual-wield default
        if hasattr(self, 'dual_wield_default'):
            self.dual_wield_var.set(self.dual_wield_default)
    
    def _export_build_results(self):
        """Export build search results to Excel"""
        # Check if there are results to export
        if not self.search_results_tv.get_children():
            messagebox.showwarning("No Results", "No build results to export. Run a search first.")
            return
        
        # Ask for save location
        from datetime import datetime
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        default_name = f"Build_Results_{timestamp}.xlsx"
        
        path = filedialog.asksaveasfilename(
            title='Export Build Results',
            initialdir=self.last_save_dir,
            initialfile=default_name,
            defaultextension='.xlsx',
            filetypes=[('Excel', '*.xlsx'), ('All', '*.*')])
        
        if not path:
            return
        
        try:
            # Create workbook
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Build Results"
            
            # Write headers (skip the visual "Div" spacer column - index 7, not real data)
            from openpyxl.styles import Font, PatternFill, Alignment
            headers = ['Slot', 'Item', 'Type', 'Spell', 'Level', 'Mob', 'Area', 'Alt Options']
            ws.append(headers)

            # Style header row
            hdr_font = Font(name='Arial', size=11, bold=True, color='FFFFFF')
            hdr_fill = PatternFill(fill_type='solid', start_color='4472C4', end_color='4472C4')
            for col_idx, header in enumerate(headers, 1):
                cell = ws.cell(1, col_idx)
                cell.font = hdr_font
                cell.fill = hdr_fill
                cell.alignment = Alignment(horizontal='center', vertical='center')

            # Write data from treeview
            for item_id in self.search_results_tv.get_children():
                values = list(self.search_results_tv.item(item_id)['values'])
                del values[7]  # drop the Div spacer column
                ws.append(values)
            
            # Auto-size columns
            for col in ws.columns:
                max_length = 0
                column = col[0].column_letter
                for cell in col:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 50)
                ws.column_dimensions[column].width = adjusted_width
            
            # Freeze header row
            ws.freeze_panes = 'A2'
            
            # Save
            wb.save(path)
            self.last_save_dir = os.path.dirname(path)
            
            messagebox.showinfo("Exported", 
                f"Build results exported successfully!\n\n{len(self.search_results_tv.get_children())} items exported to:\n{path}")
            
        except Exception as e:
            messagebox.showerror("Export Error", f"Failed to export results:\n{str(e)}")

    def _assign_required_items(self, build, covered_bases, wanted_bases):
        """Seed the build dict with user-required items, one per available
        slot (jewel/claw items fill whichever of their two slots is still
        open). Returns (crafted_count, used_claw_items) so the rest of the
        search's Crafted-cap and no-duplicate-claw rules stay consistent."""
        crafted_count = 0
        used_claw_items = []

        for item in self.required_items:
            item_slot = (item.get('Slot') or '').strip().lower()
            if item_slot == 'jewel':
                target = 'jewel_1' if 'jewel_1' not in build else ('jewel_2' if 'jewel_2' not in build else None)
            elif item_slot == 'claw':
                target = 'claw_1' if 'claw_1' not in build else ('claw_2' if 'claw_2' not in build else None)
            else:
                target = item_slot if item_slot and item_slot not in build else None

            if not target:
                continue  # no free matching slot left for this item

            build[target] = item

            item_spell = (item.get('Spell') or '').lower()
            for base in wanted_bases:
                if base in item_spell:
                    covered_bases.add(base)

            if 'crafted' in (item.get('Realm') or '').strip().lower():
                crafted_count += 1
            if item_slot == 'claw':
                used_claw_items.append(item)

        return crafted_count, used_claw_items

    def _find_optimal_build(self):
        """Find optimal item combination that covers ALL wanted spells"""
        if not self.master_data:
            messagebox.showwarning("No Data", "Please load a master database first")
            return
        
        # Get wanted spells
        wanted_spells = list(self.wanted_spells_data)
        if not wanted_spells and not self.required_items:
            messagebox.showwarning("No Spells", "Please add at least one wanted spell or required item")
            return

        # Clear previous results
        self.search_results_tv.delete(*self.search_results_tv.get_children())
        
        # Get level constraints
        min_level = None
        max_level = None
        specific_level = None
        
        # Check for specific level first (takes precedence)
        specific_level_str = self.specific_level_var.get().strip()
        if specific_level_str:
            try:
                specific_level = int(specific_level_str)
            except ValueError:
                messagebox.showwarning("Invalid Level", "Specific level must be a number")
                return
        else:
            # If no specific level, check min/max
            min_level_str = self.min_level_var.get().strip()
            if min_level_str:
                try:
                    min_level = int(min_level_str)
                except ValueError:
                    messagebox.showwarning("Invalid Level", "Minimum level must be a number")
                    return
            
            max_level_str = self.max_level_var.get().strip()
            if max_level_str:
                try:
                    max_level = int(max_level_str)
                except ValueError:
                    messagebox.showwarning("Invalid Level", "Maximum level must be a number")
                    return
            
            # Validate min <= max
            if min_level is not None and max_level is not None and min_level > max_level:
                messagebox.showwarning("Invalid Range", "Minimum level cannot be greater than maximum level")
                return
        
        # Get armor type constraints from checkboxes
        armor_constraints = {}
        for slot in ['head', 'cloak', 'body', 'hands', 'legs', 'feet']:
            # Get checked armor types for this slot
            checked_types = []
            for armor_type in ['cloth', 'leather', 'studded', 'plate']:
                if self.armor_checks[slot][armor_type].get():
                    checked_types.append(armor_type)
            armor_constraints[slot] = checked_types  # List of allowed types (empty = any)
        
        # Get weapon constraints
        weapon_style = self.weapon_style_var.get()  # 'melee', 'direct', or ''
        
        # Build config checkboxes
        wants_weapon = self.build_config_checks['weapon'].get()
        wants_shield = self.build_config_checks['shield'].get()
        wants_two_handed = self.build_config_checks['two-handed'].get()
        wants_claw_1 = self.build_config_checks['claw_1'].get()
        wants_claw_2 = self.build_config_checks['claw_2'].get()
        wants_dual_wield = self.dual_wield_var.get()
        
        # Damage type constraints
        weapon_type_constraints = []
        for weapon_type in ['slashing', 'thrusting', 'crushing']:
            if self.weapon_checks[weapon_type].get():
                weapon_type_constraints.append(weapon_type)
        
        # Filter items by constraints and group by slot
        items_by_slot = {}
        all_slots = ['head', 'jewel', 'jewel', 'cloak', 'body', 'hands', 'legs', 'feet', 'weapon', 'shield']
        
        for item in self.master_data:
            item_slot = (item.get('Slot') or '').lower()
            item_spell = (item.get('Spell') or '').lower()
            item_type = (item.get('Type') or '').lower()
            item_realm = (item.get('Realm') or '').strip()
            
            # Apply realm filter if any selected
            selected_realms = [realm for realm, var in self.realm_filters.items()
                              if var.get()]
            if selected_realms:
                realm_match = False
                for selected in selected_realms:
                    if selected.lower() in item_realm.lower():
                        realm_match = True
                        break
                if not realm_match:
                    continue
            
            
            # Skip items without spells
            if not item_spell:
                continue
            
            # Check if item's spell matches any wanted spell (partial match)
            has_wanted_spell = False
            for wanted in wanted_spells:
                if wanted in item_spell:
                    has_wanted_spell = True
                    break
            
            if not has_wanted_spell:
                continue
            
            # Apply level constraints
            if specific_level is not None or min_level is not None or max_level is not None:
                item_level_str = item.get('Level', '')
                if item_level_str:
                    try:
                        item_level = int(item_level_str)
                        
                        # Check specific level
                        if specific_level is not None:
                            if item_level != specific_level:
                                continue
                        # Check min/max range
                        else:
                            if min_level is not None and item_level < min_level:
                                continue
                            if max_level is not None and item_level > max_level:
                                continue
                    except ValueError:
                        # Skip items with invalid level data
                        continue
                else:
                    # Skip items without level data when filtering by level
                    continue
            
            # Apply armor type constraints (checkboxes)
            if item_slot in armor_constraints:
                allowed_types = armor_constraints[item_slot]
                # If any checkboxes are checked, item must match one of them
                if allowed_types:
                    type_match = False
                    for allowed_type in allowed_types:
                        if allowed_type.lower() in item_type:
                            type_match = True
                            break
                    if not type_match:
                        continue
            
            # Apply weapon/shield/claw constraints
            if item_slot == 'weapon' or item_slot == 'shield' or item_slot == 'claw':
                # Weapon style constraint (melee vs direct-caster vs direct-stat)
                if weapon_style == 'melee':
                    # Melee weapons do NOT have 'direct' in type
                    if 'direct' in item_type:
                        continue
                elif weapon_style == 'direct-caster':
                    # Direct (Caster) weapons HAVE 'direct' in type
                    # TODO: Will need to exclude stat skills once implemented
                    if 'direct' not in item_type:
                        continue
                elif weapon_style == 'direct-stat':
                    # Parry Staff - NOT YET IMPLEMENTED
                    # Will filter for stat skills when defined
                    # For now, skip all items
                    continue
                
                # Build configuration checkboxes
                # If no checkboxes are selected, accept everything
                has_any_config = wants_weapon or wants_shield or wants_two_handed or wants_claw_1 or wants_claw_2

                if has_any_config:
                    slot_accepted = False

                    if item_slot == 'weapon':
                        # Regular 1h weapons
                        if wants_weapon and '1h' in item_type:
                            slot_accepted = True
                        # Two-handed weapons
                        if wants_two_handed and '2h' in item_type:
                            slot_accepted = True
                        # Dual-wield: need 1h weapons for both hands
                        if wants_dual_wield and '1h' in item_type:
                            slot_accepted = True

                    elif item_slot == 'shield':
                        # Shields
                        if wants_shield:
                            slot_accepted = True

                    elif item_slot == 'claw':
                        # Claws - 1 Claw fills one claw slot, 2 Claw fills both
                        if wants_claw_1 or wants_claw_2:
                            slot_accepted = True

                    if not slot_accepted:
                        continue

                # Damage type constraints (for weapons/claws only, not shields)
                if item_slot in ('weapon', 'claw') and weapon_type_constraints:
                    type_match = False
                    for weapon_type in weapon_type_constraints:
                        if weapon_type.lower() in item_type:
                            type_match = True
                            break
                    if not type_match:
                        continue
            
            # Add to slot group
            if item_slot not in items_by_slot:
                items_by_slot[item_slot] = []
            items_by_slot[item_slot].append(item)
        
        # OPTIMAL BUILD ALGORITHM: Greedy approach to cover maximum spells
        # Wanted spells are grouped by base name (stripping .i/.ii/.iii) so that
        # requesting the same spell at two tiers is one requirement, not two -
        # a single item satisfying either tier fills it, instead of hunting for
        # a second item elsewhere in the build.
        wanted_bases = {}
        for wanted in wanted_spells:
            wanted_bases.setdefault(_spell_base(wanted), []).append(wanted)

        build = {}
        covered_bases = set()
        self.slot_alternates = {}

        # Force any Required Items directly into the build first - they aren't
        # subject to the other filters (armor/weapon/level/realm), since the
        # user explicitly asked for them. The rest of the build is calculated
        # around them: their spells count as covered and their slots are skipped
        # below, same as anything the greedy search would have picked itself.
        crafted_count, used_claw_items = self._assign_required_items(build, covered_bases, wanted_bases)

        # Slots to fill (including 2 jewel slots). Claw slots are only attempted
        # when 1 Claw / 2 Claw is checked - 2 Claw fills both claw slots (dual-wield).
        slots_to_fill = ['head', 'jewel_1', 'jewel_2', 'cloak', 'body', 'hands', 'legs', 'feet', 'weapon', 'shield']
        if wants_claw_2:
            slots_to_fill += ['claw_1', 'claw_2']
        elif wants_claw_1:
            slots_to_fill += ['claw_1']

        for slot in slots_to_fill:
            if slot in build:
                continue  # already filled by a Required Item

            # For jewel/claw slots, look in the shared 'jewel'/'claw' item groups
            lookup_slot = 'jewel' if slot.startswith('jewel') else 'claw' if slot.startswith('claw') else slot

            if lookup_slot not in items_by_slot:
                continue

            # Score each item by how many NEW base spells it covers; break ties
            # by preferring the item that matches the highest requested tier
            # ("best fit") so a plain-tier item doesn't edge out a better match.
            # Claw slots are exempt from the "already covered" skip: dual-wielding
            # needs a physical item in both hands even if the second one is
            # spell-redundant, unlike unique slots (head, cloak, etc).
            allow_redundant = slot.startswith('claw')
            best_item = None
            best_matched_bases = []
            best_key = (0, -1, -1)
            tied_items = []

            for item in items_by_slot[lookup_slot]:
                # A claw already equipped in the other hand can't be picked again.
                if allow_redundant and any(item is used for used in used_claw_items):
                    continue

                # Cap on Crafted-realm items across the whole build - once the
                # limit is reached, Crafted items are no longer eligible picks.
                item_realm = (item.get('Realm') or '').strip().lower()
                if 'crafted' in item_realm and crafted_count >= MAX_CRAFTED_ITEMS:
                    continue

                item_spell = (item.get('Spell') or '').lower()
                # The item's own tier, read off its actual spell text - not which
                # wanted variant matched, since "dexterity.i" is itself a substring
                # of "dexterity.iii" and would misreport a tier-iii item as tier-i.
                item_tier = _spell_tier_rank(item_spell)

                # Level tie-break: prefer the item closest to Max Level, falling
                # back to progressively lower levels when nothing sits at the max.
                try:
                    item_level_num = int(item.get('Level') or 0)
                except (ValueError, TypeError):
                    item_level_num = 0

                matched_bases = []
                for base, variants in wanted_bases.items():
                    if base in covered_bases and not allow_redundant:
                        continue
                    if any(v in item_spell for v in variants):
                        matched_bases.append(base)

                if not matched_bases:
                    continue

                key = (len(matched_bases), item_level_num, item_tier)
                if key > best_key:
                    best_key = key
                    best_item = item
                    best_matched_bases = matched_bases
                    tied_items = [item]
                elif key == best_key:
                    tied_items.append(item)

            if best_item:
                build[slot] = best_item
                covered_bases.update(best_matched_bases)
                if 'crafted' in (best_item.get('Realm') or '').strip().lower():
                    crafted_count += 1
                if allow_redundant:
                    used_claw_items.append(best_item)
                # Other items that tied for "best fit" in this slot - equally
                # valid alternate picks for a future build-set selector.
                alternates = [it for it in tied_items if it is not best_item]
                if alternates:
                    self.slot_alternates[slot] = alternates
        
        # Store the primary optimal build, then (if requested) generate additional
        # full build variants by swapping one tied alternate in at a time - lets
        # testers compare a few equally-good builds instead of just one.
        self.optimal_build = dict(build)
        self.build_variants = [dict(build)]
        if self.generate_multi_builds_var.get():
            for slot, alternates in self.slot_alternates.items():
                for alt in alternates:
                    if len(self.build_variants) >= MAX_BUILD_VARIANTS:
                        break
                    variant = dict(build)
                    variant[slot] = alt
                    self.build_variants.append(variant)
                if len(self.build_variants) >= MAX_BUILD_VARIANTS:
                    break

        variant_labels = [f"Build {i + 1}" for i in range(len(self.build_variants))]
        self.build_variant_combo['values'] = variant_labels
        self.build_variant_var.set(variant_labels[0])
        self.build_variant_combo.config(state='readonly' if len(variant_labels) > 1 else 'disabled')

        # Store and display the primary build (Build 1)
        self.last_optimal_results = self._build_dict_to_rows(self.build_variants[0])
        for row in self.last_optimal_results:
            self.search_results_tv.insert('', 'end', values=row)
        self._autosize_results_columns()

        # Show coverage (counted per distinct spell, not per tier requested)
        coverage = len(covered_bases)
        total = len(wanted_bases)
        status = (f"Optimal build covers {coverage}/{total} wanted spells | "
                   f"Spells covered: {', '.join(sorted(covered_bases))}")
        if self.slot_alternates:
            status += f" | {len(self.slot_alternates)} slot(s) have equally-good alternate items"
        if len(self.build_variants) > 1:
            status += f" | {len(self.build_variants)} build variants available"
        self.search_status.config(text=status)

        # Set display mode to optimal and switch to Results tab
        self.results_display_mode.set('optimal')
        self.notebook.select(self.tab_results)
    
    def _show_all_matches(self):
        """Show ALL items matching criteria (not optimized for spell coverage)"""
        if not self.master_data:
            messagebox.showwarning("No Data", "Please load a master database first")
            return

        # Get wanted spells
        wanted_spells = list(self.wanted_spells_data)
        if not wanted_spells:
            messagebox.showwarning("No Spells", "Please add at least one wanted spell")
            return

        # Clear previous results
        self.search_results_tv.delete(*self.search_results_tv.get_children())

        # Build variants only apply to "Find Optimal Build" results
        self.build_variants = [{}]
        self.build_variant_combo['values'] = ['Build 1']
        self.build_variant_var.set('Build 1')
        self.build_variant_combo.config(state='disabled')
        
        # Get level constraints
        min_level = None
        max_level = None
        specific_level = None
        
        # Check for specific level first (takes precedence)
        specific_level_str = self.specific_level_var.get().strip()
        if specific_level_str:
            try:
                specific_level = int(specific_level_str)
            except ValueError:
                messagebox.showwarning("Invalid Level", "Specific level must be a number")
                return
        else:
            # If no specific level, check min/max
            min_level_str = self.min_level_var.get().strip()
            if min_level_str:
                try:
                    min_level = int(min_level_str)
                except ValueError:
                    messagebox.showwarning("Invalid Level", "Minimum level must be a number")
                    return
            
            max_level_str = self.max_level_var.get().strip()
            if max_level_str:
                try:
                    max_level = int(max_level_str)
                except ValueError:
                    messagebox.showwarning("Invalid Level", "Maximum level must be a number")
                    return
            
            # Validate min <= max
            if min_level is not None and max_level is not None and min_level > max_level:
                messagebox.showwarning("Invalid Range", "Minimum level cannot be greater than maximum level")
                return
        
        # Get armor type constraints from checkboxes
        armor_constraints = {}
        for slot in ['head', 'cloak', 'body', 'hands', 'legs', 'feet']:
            # Get checked armor types for this slot
            checked_types = []
            for armor_type in ['cloth', 'leather', 'studded', 'plate']:
                if self.armor_checks[slot][armor_type].get():
                    checked_types.append(armor_type)
            armor_constraints[slot] = checked_types  # List of allowed types (empty = any)
        
        # Get weapon constraints
        weapon_style = self.weapon_style_var.get()  # 'melee', 'direct', or ''
        
        # Build config checkboxes
        wants_weapon = self.build_config_checks['weapon'].get()
        wants_shield = self.build_config_checks['shield'].get()
        wants_two_handed = self.build_config_checks['two-handed'].get()
        wants_claw_1 = self.build_config_checks['claw_1'].get()
        wants_claw_2 = self.build_config_checks['claw_2'].get()
        wants_dual_wield = self.dual_wield_var.get()
        
        # Damage type constraints
        weapon_type_constraints = []
        for weapon_type in ['slashing', 'thrusting', 'crushing']:
            if self.weapon_checks[weapon_type].get():
                weapon_type_constraints.append(weapon_type)
        
        # Filter items by constraints and group by slot
        items_by_slot = {}
        all_slots = ['head', 'jewel', 'jewel', 'cloak', 'body', 'hands', 'legs', 'feet', 'weapon', 'shield']
        
        for item in self.master_data:
            item_slot = (item.get('Slot') or '').lower()
            item_spell = (item.get('Spell') or '').lower()
            item_type = (item.get('Type') or '').lower()
            
            # Skip items without spells
            if not item_spell:
                continue
            
            # Check if item's spell matches any wanted spell (partial match)
            has_wanted_spell = False
            for wanted in wanted_spells:
                if wanted in item_spell:
                    has_wanted_spell = True
                    break
            
            item_realm = (item.get('Realm') or '').strip()
            
            # Apply realm filter if any selected
            selected_realms = [realm for realm, var in self.realm_filters.items()
                              if var.get()]
            if selected_realms:
                realm_match = False
                for selected in selected_realms:
                    if selected.lower() in item_realm.lower():
                        realm_match = True
                        break
                if not realm_match:
                    continue
            
            if not has_wanted_spell:
                continue
            
            # Apply level constraints
            if specific_level is not None or min_level is not None or max_level is not None:
                item_level_str = item.get('Level', '')
                if item_level_str:
                    try:
                        item_level = int(item_level_str)
                        
                        # Check specific level
                        if specific_level is not None:
                            if item_level != specific_level:
                                continue
                        # Check min/max range
                        else:
                            if min_level is not None and item_level < min_level:
                                continue
                            if max_level is not None and item_level > max_level:
                                continue
                    except ValueError:
                        # Skip items with invalid level data
                        continue
                else:
                    # Skip items without level data when filtering by level
                    continue
            
            # Apply armor type constraints (checkboxes)
            if item_slot in armor_constraints:
                allowed_types = armor_constraints[item_slot]
                # If any checkboxes are checked, item must match one of them
                if allowed_types:
                    type_match = False
                    for allowed_type in allowed_types:
                        if allowed_type.lower() in item_type:
                            type_match = True
                            break
                    if not type_match:
                        continue
            
            # Apply weapon/shield/claw constraints
            if item_slot == 'weapon' or item_slot == 'shield' or item_slot == 'claw':
                # Weapon style constraint (melee vs direct-caster vs direct-stat)
                if weapon_style == 'melee':
                    # Melee weapons do NOT have 'direct' in type
                    if 'direct' in item_type:
                        continue
                elif weapon_style == 'direct-caster':
                    # Direct (Caster) weapons HAVE 'direct' in type
                    # TODO: Will need to exclude stat skills once implemented
                    if 'direct' not in item_type:
                        continue
                elif weapon_style == 'direct-stat':
                    # Parry Staff - NOT YET IMPLEMENTED
                    # Will filter for stat skills when defined
                    # For now, skip all items
                    continue
                
                # Build configuration checkboxes
                # If no checkboxes are selected, accept everything
                has_any_config = wants_weapon or wants_shield or wants_two_handed or wants_claw_1 or wants_claw_2

                if has_any_config:
                    slot_accepted = False

                    if item_slot == 'weapon':
                        # Regular 1h weapons
                        if wants_weapon and '1h' in item_type:
                            slot_accepted = True
                        # Two-handed weapons
                        if wants_two_handed and '2h' in item_type:
                            slot_accepted = True
                        # Dual-wield: need 1h weapons for both hands
                        if wants_dual_wield and '1h' in item_type:
                            slot_accepted = True

                    elif item_slot == 'shield':
                        # Shields
                        if wants_shield:
                            slot_accepted = True

                    elif item_slot == 'claw':
                        # Claws - 1 Claw fills one claw slot, 2 Claw fills both
                        if wants_claw_1 or wants_claw_2:
                            slot_accepted = True

                    if not slot_accepted:
                        continue

                # Damage type constraints (for weapons/claws only, not shields)
                if item_slot in ('weapon', 'claw') and weapon_type_constraints:
                    type_match = False
                    for weapon_type in weapon_type_constraints:
                        if weapon_type.lower() in item_type:
                            type_match = True
                            break
                    if not type_match:
                        continue
            
            # Add to slot group
            if item_slot not in items_by_slot:
                items_by_slot[item_slot] = []
            items_by_slot[item_slot].append(item)
        
        # Display ALL matching items (not just one per slot)
        # Group by slot for better organization
        all_matches = []
        
        for slot_key, items in items_by_slot.items():
            for item in items:
                all_matches.append((slot_key, item))
        
        # Sort by slot order for display
        slot_priority = {'head': 1, 'jewel': 2, 'cloak': 3, 'body': 4, 'hands': 5, 
                        'legs': 6, 'feet': 7, 'weapon': 8, 'shield': 9, 'claw': 10}
        all_matches.sort(key=lambda x: slot_priority.get(x[0], 99))
        
        # Track which spells are covered
        covered_spells = set()
        
        # Store and display all matching items
        self.last_all_results = []
        for slot, item in all_matches:
            item_spell = (item.get('Spell') or '').lower()
            
            # Check if this item provides any wanted spells
            for wanted in wanted_spells:
                if wanted in item_spell:
                    covered_spells.add(wanted)
            
            row = (
                slot.title(),
                item.get('Item', ''),
                item.get('Type', ''),
                item.get('Spell', ''),
                item.get('Level', ''),
                item.get('Mob', ''),
                item.get('Area', ''),
                '████████',
                ''
            )
            self.last_all_results.append(row)
            self.search_results_tv.insert('', 'end', values=row)
        self._autosize_results_columns()

        # Show coverage
        coverage = len(covered_spells)
        total = len(wanted_spells)
        self.search_status.config(
            text=f"All matches show {coverage}/{total} wanted spells | Spells covered: {', '.join(sorted(covered_spells))}")
        
        # Set display mode to all and switch to Results tab
        self.results_display_mode.set('all')
        self.notebook.select(self.tab_results)



def main():
    app = App()
    app.mainloop()


if __name__ == '__main__':
    main()
