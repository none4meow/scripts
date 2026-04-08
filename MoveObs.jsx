/**
 * Illustrator ExtendScript: Pack selected objects into 90x90cm boxes (rectangle-grid, FAST)
 *
 * Menu:
 *  Option 1: Pack ONLY (do NOT draw box + label)
 *  Option 2: Pack + Draw box + label
 *
 * Settings:
 *  - BOX = 90cm
 *  - BOX_PAD = 1cm
 *  - gap between objects = 0.3cm  (implemented as per-side OBJ_PAD = 0.15cm)
 *  - CELL = 0.15cm (tight grid)
 *  - New box when full; max 6 boxes per row; box gap = 10cm
 *  - Label is 10cm below box (only in Option 2)
 *  - Sort small first, but GROUP items by area within +3 cm^2 buckets (stable within group)
 *
 * Notes:
 *  - Bounds for sizing AND moving: union(visibleBounds, geometricBounds) for consistency (reduces random extra gaps)
 */

(function () {
  if (app.documents.length === 0) { alert("Open a document first."); return; }
  var doc = app.activeDocument;
  var sel = doc.selection;
  if (!sel || sel.length === 0) { alert("Select the objects (groups are OK) you want to pack, then run again."); return; }

  // ----------------------------
  // Menu (ScriptUI)
  // ----------------------------
  var doDraw = null;
  try {
    var w = new Window("dialog", "Packing Options");
    w.alignChildren = "fill";

    var b1 = w.add("button", undefined, "Option 1: Pack only");
    var b2 = w.add("button", undefined, "Option 2: Pack + draw box + label");
    var cancel = w.add("button", undefined, "Cancel");

    b1.onClick = function(){ doDraw = false; w.close(1); };
    b2.onClick = function(){ doDraw = true;  w.close(1); };
    cancel.onClick = function(){ w.close(0); };

    if (w.show() !== 1 || doDraw === null) return;
  } catch (eUI) {
    // If ScriptUI fails for any reason, fallback to confirm
    doDraw = confirm("Draw box + label? (OK = Option 2, Cancel = Option 1)");
  }

  // ----------------------------
  // Constants (cm)
  // ----------------------------
  var BOX = 90;
  var BOX_PAD = 1;

  var OBJ_GAP = 0.3;
  var OBJ_PAD = OBJ_GAP / 2; // 0.15cm per side

  var CELL = 0.3; // tighter grid (1mm)

  var BOX_GAP = 10;
  var MAX_PER_ROW = 6;

  var AREA_TOL = 3; // cm^2 grouping tolerance (+3 cm^2)

  var USE_CM = BOX - 2 * BOX_PAD; // 88
  var GW = Math.floor(USE_CM / CELL);
  var GH = Math.floor(USE_CM / CELL);
  var USE_EFF_CM = GW * CELL;

  // ----------------------------
  // Unit conversion
  // ----------------------------
  var PT_PER_CM = 72 / 2.54;
  function cmToPt(cm) { return cm * PT_PER_CM; }
  function ptToCm(pt) { return pt / PT_PER_CM; }

  // ----------------------------
  // Bounds helpers
  // ----------------------------
  // bounds: [L, T, R, B]
  function getUnionBounds(item) {
    var gb = null, vb = null;
    try { gb = item.geometricBounds; } catch (e1) {}
    try { vb = item.visibleBounds; } catch (e2) {}

    if (!gb && !vb) return null;
    if (!gb) return vb;
    if (!vb) return gb;

    return [
      Math.min(gb[0], vb[0]),
      Math.max(gb[1], vb[1]),
      Math.max(gb[2], vb[2]),
      Math.min(gb[3], vb[3])
    ];
  }

  function boundsWidthPt(b) { return b[2] - b[0]; }
  function boundsHeightPt(b) { return b[1] - b[3]; }

  // Move bounds MUST match sizing bounds for consistent spacing
  function getMoveBounds(item) { return getUnionBounds(item); }

  // ----------------------------
  // Grid helpers
  // ----------------------------
  function initGrid(w, h) {
    var g = new Array(h);
    for (var y = 0; y < h; y++) {
      var row = new Array(w);
      for (var x = 0; x < w; x++) row[x] = 0;
      g[y] = row;
    }
    return g;
  }

  function regionEmpty(grid, x, y, cw, ch) {
    for (var yy = y; yy < y + ch; yy++) {
      var row = grid[yy];
      for (var xx = x; xx < x + cw; xx++) {
        if (row[xx] === 1) return false;
      }
    }
    return true;
  }

  function occupy(grid, x, y, cw, ch) {
    for (var yy = y; yy < y + ch; yy++) {
      var row = grid[yy];
      for (var xx = x; xx < x + cw; xx++) row[xx] = 1;
    }
  }

  function findSpot(grid, cw, ch) {
    // bottom -> right, then up
    var maxX = GW - cw;
    var maxY = GH - ch;

    for (var y = 0; y <= maxY; y++) {
      for (var x = 0; x <= maxX; x++) {
        if (regionEmpty(grid, x, y, cw, ch)) return { x: x, y: y };
      }
    }
    return null;
  }

  // ----------------------------
  // Drawing helpers (optional)
  // ----------------------------
  function createBoxAt(boxLeftPt, boxTopPt) {
    var rect = doc.pathItems.rectangle(boxTopPt, boxLeftPt, cmToPt(BOX), cmToPt(BOX));
    rect.stroked = true;
    rect.filled = false;
    return rect;
  }

  function createCenteredLabel(boxIndex1Based, boxLeftPt, boxTopPt) {
    var labelY = boxTopPt - cmToPt(BOX + 10);
    var label = doc.textFrames.pointText([boxLeftPt, labelY]);
    label.contents = "Box " + boxIndex1Based;

    try {
      var b = label.visibleBounds;
      var textCenterX = (b[0] + b[2]) / 2;
      var boxCenterX = boxLeftPt + cmToPt(BOX) / 2;
      label.translate(boxCenterX - textCenterX, 0);
    } catch (e) {}
    return label;
  }

  // ----------------------------
  // Move item so its (union) bottom-left lands on target
  // ----------------------------
  function moveItemBottomLeftTo(item, targetBLxPt, targetBLyPt) {
    var b = getMoveBounds(item);
    if (!b) return false;

    item.translate(targetBLxPt - b[0], targetBLyPt - b[3]);

    // corrective pass for rounding
    var b2 = getMoveBounds(item);
    if (b2) item.translate(targetBLxPt - b2[0], targetBLyPt - b2[3]);

    return true;
  }

  // ----------------------------
  // Collect items
  // ----------------------------
  function isPackable(item) {
    try {
      if (item.locked) return false;
      if (item.hidden) return false;
    } catch (e) {}
    return getUnionBounds(item) !== null;
  }

  var items = [];
  for (var i = 0; i < sel.length; i++) {
    var it = sel[i];
    if (!isPackable(it)) continue;

    var ub = getUnionBounds(it);
    var wCm = ptToCm(boundsWidthPt(ub));
    var hCm = ptToCm(boundsHeightPt(ub));

    var wPad = wCm + 2 * OBJ_PAD;
    var hPad = hCm + 2 * OBJ_PAD;

    var cw = Math.ceil(wPad / CELL);
    var ch = Math.ceil(hPad / CELL);

    items.push({
      item: it,
      area: wCm * hCm,
      cw: cw,
      ch: ch
    });
  }

  if (items.length === 0) { alert("No packable items found in selection."); return; }

  // ----------------------------
  // Sort small first + group by area within +3 cm^2
  // ----------------------------
  items.sort(function (a, b) { return a.area - b.area; });

  var grouped = [];
  var idx = 0;
  while (idx < items.length) {
    var startArea = items[idx].area;
    grouped.push(items[idx]);
    idx++;
    while (idx < items.length && items[idx].area <= startArea + AREA_TOL) {
      grouped.push(items[idx]);
      idx++;
    }
  }
  items = grouped;

  // ----------------------------
  // Artboard anchoring + box layout
  // ----------------------------
  var abIndex = doc.artboards.getActiveArtboardIndex();
  var abRect = doc.artboards[abIndex].artboardRect; // [L, T, R, B]
  var AB_L = abRect[0];
  var AB_T = abRect[1];

  function boxTopLeftForIndex(boxIndex) {
    var col = boxIndex % MAX_PER_ROW;
    var row = Math.floor(boxIndex / MAX_PER_ROW);
    return {
      left: AB_L + cmToPt(col * (BOX + BOX_GAP)),
      top:  AB_T - cmToPt(row * (BOX + BOX_GAP))
    };
  }

  function usableOriginForBox(boxLeftPt, boxTopPt) {
    var usableLeft = boxLeftPt + cmToPt(BOX_PAD);
    var usableBottom = boxTopPt - cmToPt(BOX_PAD + USE_EFF_CM);
    return { left: usableLeft, bottom: usableBottom };
  }

  // ----------------------------
  // Pack across boxes
  // ----------------------------
  var boxIndex = 0;
  var placedCount = 0;
  var unplaced = [];

  var pos0 = boxTopLeftForIndex(boxIndex);
  if (doDraw) {
    createBoxAt(pos0.left, pos0.top);
    createCenteredLabel(boxIndex + 1, pos0.left, pos0.top);
  }
  var grid = initGrid(GW, GH);

  for (var n = 0; n < items.length; n++) {
    var obj = items[n];

    if (obj.cw > GW || obj.ch > GH) {
      unplaced.push(obj.item);
      continue;
    }

    var spot = findSpot(grid, obj.cw, obj.ch);

    if (!spot) {
      boxIndex++;
      var pos = boxTopLeftForIndex(boxIndex);
      if (doDraw) {
        createBoxAt(pos.left, pos.top);
        createCenteredLabel(boxIndex + 1, pos.left, pos.top);
      }
      grid = initGrid(GW, GH);
      spot = findSpot(grid, obj.cw, obj.ch);
    }

    if (!spot) {
      unplaced.push(obj.item);
      continue;
    }

    occupy(grid, spot.x, spot.y, obj.cw, obj.ch);

    var posBox = boxTopLeftForIndex(boxIndex);
    var usable = usableOriginForBox(posBox.left, posBox.top);

    var targetX = usable.left + cmToPt(spot.x * CELL + OBJ_PAD);
    var targetY = usable.bottom + cmToPt(spot.y * CELL + OBJ_PAD);

    if (moveItemBottomLeftTo(obj.item, targetX, targetY)) placedCount++;
    else unplaced.push(obj.item);
  }

  // ----------------------------
  // Report
  // ----------------------------
  var msg = "";
  msg += "Packing finished.\n\n";
  msg += "Mode: " + (doDraw ? "Option 2 (draw box + label)" : "Option 1 (pack only)") + "\n";
  msg += "Placed: " + placedCount + " / " + items.length + "\n";
  msg += "Boxes used: " + (boxIndex + 1) + "\n";
  msg += "CELL: " + CELL + "cm\n";
  msg += "BOX_PAD: " + BOX_PAD + "cm\n";
  msg += "Gap between objects: " + OBJ_GAP + "cm (per-side pad " + OBJ_PAD + "cm)\n";
  msg += "Area grouping tolerance: +" + AREA_TOL + " cm^2\n";
  msg += "Usable (grid-covered): " + USE_EFF_CM + "cm\n";
  if (unplaced.length > 0) msg += "\nUnplaced items: " + unplaced.length + "\n";

  alert(msg);
})();