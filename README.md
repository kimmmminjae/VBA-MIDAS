# VBA-MIDAS

Excel VBA module for the MIDAS CIVIL NX / GEN NX Open API.

## Download

Latest release: **[CivilVBA.zip](https://github.com/kimmmminjae/VBA-MIDAS/releases/latest/download/CivilVBA.zip)**
(`CivilVBA.bas` + `JsonConverter.bas`). All versions: [Releases](https://github.com/kimmmminjae/VBA-MIDAS/releases).

## Files

- `src/CivilVBA.bas` - main module
- `src/JsonConverter.bas` - JSON parser ([VBA-JSON](https://github.com/VBA-tools/VBA-JSON), MIT), required by `CivilVBA.bas`

## Usage

1. In the Excel VBA editor, import `CivilVBA.bas` and `JsonConverter.bas`.
2. Add a reference to *Microsoft Scripting Runtime*.
3. Call model functions (`Node`, `Beam`, `Material`, ...) and finish with `ModelCreate`
   (or `RunAnalysis` / `SaveFile`, which send the stored model first).

## Publishing a release

1. Update `src/CivilVBA.bas` and write the notes for the new version in `RELEASE_NOTES.md`.
2. Merge to `main`, then tag and push:

   ```
   git tag v1.2
   git push origin v1.2
   ```

   A tag ending in `-beta` (e.g. `v1.2-beta`) gets "Beta" in the release title.
3. The `Release` workflow builds `CivilVBA.zip` and publishes the release.
   It is always a full release, so the `releases/latest/download/...` links point at it.
