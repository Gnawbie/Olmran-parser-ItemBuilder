# Arch Linux packaging

`PKGBUILD` wraps the prebuilt Linux binary that `.github/workflows/build-linux.yml`
attaches to each GitHub Release (`OlmranItemBuilder-linux`), rather than compiling
from source - this project has no `setup.py`/`pyproject.toml` of its own, so
reusing the same PyInstaller onefile binary every other Linux download uses is
simpler and more reliable than a separate from-source build here.

## Building locally

```bash
makepkg -si
```

## Publishing a new version

1. Bump `pkgver` in `PKGBUILD` to match the new release's version (no leading `v`).
2. Reset `pkgrel=1`.
3. Run `updpkgsums` (from the `pacman-contrib` package) to fill in the real
   `sha256sums` once the release (and its `OlmranItemBuilder-linux` asset from
   the Linux build workflow) actually exists - they're left as `SKIP` in this
   repo since a not-yet-tagged version's asset doesn't exist yet.
4. If publishing to the AUR, copy the updated `PKGBUILD` (and `.desktop`/icon
   source) into the AUR git repo and push as usual.
