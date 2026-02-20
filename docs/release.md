# Release Process

This repository uses a tag-driven release flow.

## Standard release steps

1. Ensure `main` is clean and pushed.
2. Bump version in:
   - `pyproject.toml`
   - `src/Mudabbir/__init__.py`
3. Commit and push the version bump to `main`.
4. Create and push an annotated tag:
   - `git tag -a vX.Y.Z -m "Mudabbir vX.Y.Z"`
   - `git push origin refs/tags/vX.Y.Z`
5. Verify GitHub Actions:
   - `Build Desktop Launcher`
   - `Publish to PyPI`
6. Verify GitHub Release has 6 assets and PyPI contains both wheel and sdist.

## Important: immutable tag names

If GitHub returns:

- `Cannot create ref due to creations being restricted`
- or `tag_name was used by an immutable release`

then that exact tag name cannot be reused.

Use the next available version tag (for example, `v0.4.10` instead of a blocked `v0.4.4`) and keep package version/tag aligned.
