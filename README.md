# Scripts

This repository contains Adobe Illustrator ExtendScript automation.

## Files

- [PackByTypeColor.jsx](/D:/Github/Scripts/PackByTypeColor.jsx)
  Packs Illustrator artwork into 90 x 90 cm boxes by color and type.

## PackByTypeColor

Current menu flow:

1. `Pack for lasercut` or `Pack for print`
2. `Pack current selection` or `Pack from folder`
3. `Pack only` or `Pack + draw box + label`

Current source behavior:

- `Pack for print -> Pack from folder` imports items, selects them, and stops.
- `Pack for print -> Pack current selection` does the actual print packing run.
- Lasercut supports both current selection and folder import as direct packing sources.

Current packing behavior:

- Box size is 90 cm with 1 cm box padding.
- Objects are grouped by matched color and type before packing.
- Type is stroke-based: `#0000FF = 5mm`, otherwise `3mm`.
- Small items fill bottom rows first, then larger items pack above them.
- Items with the same packed `cw` and `ch` stay adjacent inside each group.
- Clipped artwork sizes from clipping mask bounds; non-clipped containers size from recursive child bounds.
- Drawn boxes use a black stroke and no fill.

Current lasercut post-pack styling:

- Print mode preserves fill and stroke.
- Standalone top-level red `#FF0000` path or compound path turns red stroke off and keeps fill.
- Whole `3mm` grouped items with all descendant path-like leaves red-stroked turn red stroke off and keep fill.
- Mixed `3mm` groups can style direct children differently:
  - first rendered all-red subgroup gets red stroke off
  - first rendered red leaf gets only that leaf's red stroke off
  - other direct siblings use the normal fill-off path
- Normal fill-off only applies to leaves that have both usable fill and usable stroke.

## Notes for edits

- This repo targets old Illustrator ExtendScript, not modern browser JavaScript.
- Keep changes minimal and patch the existing script instead of rewriting it.
- Be explicit about bounds behavior when changing placement logic.
