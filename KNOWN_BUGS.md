# Known Bugs

## Item Counter: bogus "areas" still showing up in Zone charts

**Status:** Open - needs a real log sample to diagnose further.

**Symptom:** The Item Counter's Zone breakdown (Overall "All" tree, and
each of Daily/Weekly/Monthly/Yearly) shows entries that aren't real
zones - e.g. `Demon 74`, `Demon 75`, `Sin 60`, `Sin 61`, `Sin 68`,
`Skeleton 60` - mixed in among real zone names like `Cerulean Lakes`,
`City of Zhak-Tor`, `Dark Forest`, etc.

**Where it comes from:** `_parse_item_events` in `gaming_log_parser.py`
tracks the "current area" via `LootParser.LOC_RE`, which matches any
line starting with `[...]` and treats the bracket contents as a zone
name (sticky until the next `[...]` line). Something in the user's logs
produces a bracket line that matches this pattern but isn't a real zone
transition - most likely a class/summon/pet status display of some
kind, based on the shape of the bogus entries (a short name + a
level-like number).

**Attempted fixes (none confirmed working per user report):**
1. Matched the bracket text against `BLOCKED_AREAS` (`{'class'}`) - the
   set used to exclude master-database items whose own `Area` column is
   literally "class". Wrong target: that's a master-data item field,
   unrelated to log-parsed zone names.
2. Matched the bracket text (with a trailing " <number>" stripped, exact
   match) against the `Mob` field of `self.blocked_area_items` - the
   Class Items shown in the Build tab's own "Class Items" typeahead
   (BARD, BURNING HAND, COVEN, etc.). User reports this still doesn't
   catch the bogus entries.
3. Structural check: treat any bracket text containing 3+ consecutive
   spaces before a trailing number as not a real zone (see
   `blocked_area_gap_re` in `_parse_item_events`). User reports this
   still doesn't catch it either.

**Next step:** Get an actual raw excerpt of the log line(s) that produce
one of these bogus entries (e.g. paste the exact line(s) around where
`Demon 74` or `Sin 60` shows up) so the real bracket text/spacing/shape
can be seen directly, instead of guessing at the pattern.

**Relevant code:** `_parse_item_events`, `is_blocked_area`/
`blocked_area_gap_re` (`gaming_log_parser.py`), `LootParser.LOC_RE`.
