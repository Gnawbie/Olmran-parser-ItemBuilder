# Contributing to Olmran Item Builder

Thanks for your interest in contributing! This is a small, actively-maintained
solo project, but bug reports, feature requests, and pull requests are welcome.

## Reporting bugs

Open a [GitHub issue](https://github.com/Gnawbie/Olmran-parser-ItemBuilder/issues/new)
and include:
- What you did and what you expected to happen
- What actually happened (screenshots help a lot for UI issues)
- Your Windows version and whether you're running `OlmranItemBuilder.exe` or
  `gaming_log_parser.py` from source

## Suggesting features

Open an issue describing the use case - what build/search/parsing scenario
you're trying to solve, not just the specific UI you have in mind. That
context makes it much easier to design a fix that fits the rest of the app.

## Submitting changes

1. Fork the repo and create a branch off `main`
2. Make your change to `gaming_log_parser.py`
3. Test it by running `python gaming_log_parser.py` (requires `openpyxl` and
   `pillow` - see INSTALL_INSTRUCTIONS.txt)
4. Open a pull request describing what changed and why

Please keep pull requests focused - one fix or feature per PR is much easier
to review than a bundle of unrelated changes.

## Code style

There's no formal linter configured. Match the existing style in
`gaming_log_parser.py`: minimal comments (only where the *why* isn't obvious
from the code), descriptive variable names, and no unrelated refactoring
mixed into a bug-fix or feature PR.

## Code of Conduct

This project follows a [Code of Conduct](CODE_OF_CONDUCT.md). By
participating, you're expected to uphold it.
