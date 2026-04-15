/**
 * Illustrator ExtendScript: Pack selected objects into 90x90cm boxes (bottom-strip small rows + dense upper packing)
 *
 * Menu:
 *  Choose workflow: Pack for lasercut or Pack for print
 *  Choose source: current selection or folder import under E:\DON GO\DON GO 2604
 *  Folder import chooser lists direct child folders newest modified first, then their subfolders
 *  Option 1: Pack ONLY (do NOT draw box + label)
 *  Option 2: Pack + Draw box + label
 *  Pack for print + folder import selects imported items and stops so you can choose what to pack
 *
 * Settings:
 *  - BOX = 90cm
 *  - BOX_PAD = 1cm
 *  - gap between objects = 0.3cm  (implemented as per-side OBJ_PAD = 0.15cm)
 *  - CELL = 0.15cm (tight grid)
 *  - New box when full; max 6 boxes per row; box gap = 10cm
 *  - Label is 10cm below box (only in Option 2)
 *  - Option 1 (Pack only) packs all eligible items as one combined pool
 *  - Option 2 groups objects by matched color + type before packing
 *  - Type is based on stroke color (#0000FF = 5mm, otherwise 3mm)
 *  - Option 2 uses one color/type group per box; group overflow continues to the next box
 *  - Small items are packed into bottom rows first; larger items are then packed above them
 *
 * Notes:
 *  - Bounds for sizing AND moving: clipped groups use the clipping mask path geometricBounds; other items use union(visibleBounds, geometricBounds)
 */

(function () {
  if (app.documents.length === 0) {
    alert("Open a target document first.");
    return;
  }
  var doc = app.activeDocument;
  var WORKFLOW_LASERCUT = "lasercut";
  var WORKFLOW_PRINT = "print";
  var SOURCE_SELECTION = "selection";
  var SOURCE_FOLDER = "folder";
  var workflowMode = null;
  var sourceMode = null;

  // ----------------------------
  // Workflow mode (ScriptUI)
  // ----------------------------
  try {
    var workflowWindow = new Window("dialog", "Packing Workflow");
    workflowWindow.alignChildren = "fill";

    var workflowLasercut = workflowWindow.add(
      "button",
      undefined,
      "Pack for lasercut",
    );
    var workflowPrint = workflowWindow.add(
      "button",
      undefined,
      "Pack for print",
    );
    var workflowCancel = workflowWindow.add("button", undefined, "Cancel");

    workflowLasercut.onClick = function () {
      workflowMode = WORKFLOW_LASERCUT;
      workflowWindow.close(1);
    };
    workflowPrint.onClick = function () {
      workflowMode = WORKFLOW_PRINT;
      workflowWindow.close(1);
    };
    workflowCancel.onClick = function () {
      workflowWindow.close(0);
    };

    if (workflowWindow.show() !== 1 || workflowMode === null) return;
  } catch (eWorkflowUI) {
    var workflowFallbackChoice = prompt(
      "Packing Workflow:\n1 = Pack for lasercut\n2 = Pack for print\n\nEnter 1 or 2.",
      "",
    );
    if (workflowFallbackChoice === null) return;
    workflowFallbackChoice = workflowFallbackChoice.replace(/^\s+|\s+$/g, "");
    if (workflowFallbackChoice === "1") workflowMode = WORKFLOW_LASERCUT;
    else if (workflowFallbackChoice === "2") workflowMode = WORKFLOW_PRINT;
    else return;
  }

  // ----------------------------
  // Source mode (ScriptUI)
  // ----------------------------
  try {
    var sourceWindow = new Window("dialog", "Packing Source");
    sourceWindow.alignChildren = "fill";

    var sourceSelection = sourceWindow.add(
      "button",
      undefined,
      "Pack current selection",
    );
    var sourceFolder = sourceWindow.add(
      "button",
      undefined,
      "Pack from folder",
    );
    var sourceCancel = sourceWindow.add("button", undefined, "Cancel");

    sourceSelection.onClick = function () {
      sourceMode = SOURCE_SELECTION;
      sourceWindow.close(1);
    };
    sourceFolder.onClick = function () {
      sourceMode = SOURCE_FOLDER;
      sourceWindow.close(1);
    };
    sourceCancel.onClick = function () {
      sourceWindow.close(0);
    };

    if (sourceWindow.show() !== 1 || sourceMode === null) return;
  } catch (eSourceUI) {
    var sourceFallbackChoice = prompt(
      "Packing Source:\n1 = Pack current selection\n2 = Pack from folder\n\nEnter 1 or 2.",
      "",
    );
    if (sourceFallbackChoice === null) return;
    sourceFallbackChoice = sourceFallbackChoice.replace(/^\s+|\s+$/g, "");
    if (sourceFallbackChoice === "1") sourceMode = SOURCE_SELECTION;
    else if (sourceFallbackChoice === "2") sourceMode = SOURCE_FOLDER;
    else return;
  }

  function chooseDrawMode() {
    var drawChoice = null;

    try {
      var w = new Window("dialog", "Packing Options");
      w.alignChildren = "fill";

      var b1 = w.add("button", undefined, "Option 1: Pack only");
      var b2 = w.add("button", undefined, "Option 2: Pack + draw box + label");
      var cancel = w.add("button", undefined, "Cancel");

      b1.onClick = function () {
        drawChoice = false;
        w.close(1);
      };
      b2.onClick = function () {
        drawChoice = true;
        w.close(1);
      };
      cancel.onClick = function () {
        w.close(0);
      };

      if (w.show() !== 1 || drawChoice === null) return null;
    } catch (eUI) {
      var fallbackChoice = prompt(
        "Packing Options:\n1 = Pack only\n2 = Pack + draw box + label\n\nEnter 1 or 2.",
        "",
      );
      if (fallbackChoice === null) return null;
      fallbackChoice = fallbackChoice.replace(/^\s+|\s+$/g, "");
      if (fallbackChoice === "1") drawChoice = false;
      else if (fallbackChoice === "2") drawChoice = true;
      else return null;
    }

    return drawChoice;
  }

  var doDraw = null;

  // ----------------------------
  // Constants (cm)
  // ----------------------------
  var BOX = 90;
  var BOX_PAD = 1;

  var OBJ_GAP = 0.3;
  var OBJ_PAD = OBJ_GAP / 2; // 0.15cm per side

  var CELL = 0.3; // tighter grid (1mm)

  var BOX_COL_GAP = 10;
  var BOX_ROW_GAP = 30;
  var LABEL_OFFSET = 10;
  var LABEL_FONT_NAME = "Fraunces";
  var LABEL_FONT_SIZE = 220;
  var MAIN_FOLDER_PATH = "E:/DON GO/DON GO 2604";
  var SMALL_BUCKET_MAX_SIDE_CM = 18;
  var COLOR_TOLERANCE = 12;
  var TYPE_5MM = "5mm";
  var TYPE_3MM = "3mm";
  var TYPE_5MM_STROKE_HEX = "#0000FF";
  var RED_STROKE_OFF_HEX = "#FF0000";
  var MAX_PER_ROW = 5;
  var SMALL_BUCKET_MAX_CELLS = Math.ceil(SMALL_BUCKET_MAX_SIDE_CM / CELL);

  var TYPE_SORT_ORDER = {};
  TYPE_SORT_ORDER[TYPE_3MM] = 0;
  TYPE_SORT_ORDER[TYPE_5MM] = 1;

  var COLOR_PALETTE = [
    { label: "Mint", name: "2551 - Mint", hex: "#ABDDBA" },
    { label: "Light Sage", name: "B17 - Light Sage", hex: "#D4E5CA" },
    { label: "Lemon Green", name: "B22 - Lemon Green", hex: "#DFDF8E" },
    { label: "Olive Green", name: "B38 - Olive Green", hex: "#959F66" },
    { label: "White", name: "Trang - White", hex: "#FFFFFF" },
    { label: "Baby Blue", name: "B07 - Baby Blue", hex: "#A4EAE4" },
    { label: "Teal", name: "B14 - Teal", hex: "#2CB1AE" },
    { label: "Blue", name: "B66 - Blue", hex: "#126489" },
    { label: "Pastel Pink", name: "B47 - Pastel Pink", hex: "#EDD2CE" },
    { label: "Blush Pink", name: "B62 - Blush Pink", hex: "#F4A5AC" },
    { label: "Lilac", name: "B86 - Lilac", hex: "#C992D3" },
    { label: "Retro Purple", name: "2552 - Retro Purple", hex: "#8085AF" },
    { label: "Yellow", name: "B03 - Yellow", hex: "#FEE36E" },
    { label: "Gold", name: "Nhu vang - Gold", hex: "#EEB453" },
    { label: "Orange", name: "B55 - Orange", hex: "#FF8A3D" },
    { label: "Red", name: "Do tuoi - Red", hex: "#F2340F" },
    { label: "Silver", name: "Nhu bac - Silver", hex: "#E2DED9" },
    { label: "Grey", name: "B11 - Grey", hex: "#C6C0BD" },
    { label: "Black", name: "Den - Black", hex: "#000000" },
    { label: "Brown", name: "Nau 10 - Brown", hex: "#5A4029" },
    { label: "Do man", name: "Do man - B61", hex: "#A0180A" },
    { label: "Classic Gray", name: "37 - Classic Gray", hex: "#BDBABC" },
    { label: "Ebony", name: "Ebony", hex: "#535458" },
    { label: "Early America", name: "23 - Early America", hex: "#9B6E4F" },
    { label: "Jacobean", name: "19 - Jacobean", hex: "#876E56" },
    { label: "Koson", name: "Ko son", hex: "#E1D6C6" },
    { label: "Vecny", name: "Vecny - Honey Maple", hex: "#EACCA7" },
    { label: "Golden Oak", name: "15 - Golden Oak", hex: "#CE925C" },
    { label: "Candlelite", name: "07 - Candlelite", hex: "#B1592C" },
    { label: "Nau nhat", name: "Nau nhat", hex: "#896F56" },
    { label: "Mica", name: "Mica", hex: "#e2e2e2" },
  ];

  var USE_CM = BOX - 2 * BOX_PAD; // 88
  var GW = Math.floor(USE_CM / CELL);
  var GH = Math.floor(USE_CM / CELL);
  var USE_EFF_CM = GW * CELL;

  // ----------------------------
  // Unit conversion
  // ----------------------------
  var PT_PER_CM = 72 / 2.54;
  function cmToPt(cm) {
    return cm * PT_PER_CM;
  }
  function ptToCm(pt) {
    return pt / PT_PER_CM;
  }

  // ----------------------------
  // Color helpers
  // ----------------------------
  function clampByte(value) {
    value = Math.round(value);
    if (value < 0) return 0;
    if (value > 255) return 255;
    return value;
  }

  function byteToHex(value) {
    var hex = clampByte(value).toString(16).toUpperCase();
    return hex.length < 2 ? "0" + hex : hex;
  }

  function normalizeHex(hex) {
    if (!hex) return null;

    var text = String(hex)
      .replace(/^\s+|\s+$/g, "")
      .toUpperCase();
    if (!text) return null;
    if (text.charAt(0) !== "#") text = "#" + text;

    if (text.length === 4) {
      text =
        "#" +
        text.charAt(1) +
        text.charAt(1) +
        text.charAt(2) +
        text.charAt(2) +
        text.charAt(3) +
        text.charAt(3);
    }

    if (!/^#[0-9A-F]{6}$/.test(text)) return null;
    return text;
  }

  function hexToRgb(hex) {
    var normalized = normalizeHex(hex);
    if (!normalized) return null;

    return {
      r: parseInt(normalized.substr(1, 2), 16),
      g: parseInt(normalized.substr(3, 2), 16),
      b: parseInt(normalized.substr(5, 2), 16),
    };
  }

  function rgbToHex(r, g, b) {
    return "#" + byteToHex(r) + byteToHex(g) + byteToHex(b);
  }

  function makeRgbColorFromHex(hex) {
    var rgb = hexToRgb(hex);
    if (!rgb) return null;

    var color = new RGBColor();
    color.red = rgb.r;
    color.green = rgb.g;
    color.blue = rgb.b;
    return color;
  }

  function cmykToHex(c, m, y, k) {
    var cc = Math.max(0, Math.min(100, c)) / 100;
    var mm = Math.max(0, Math.min(100, m)) / 100;
    var yy = Math.max(0, Math.min(100, y)) / 100;
    var kk = Math.max(0, Math.min(100, k)) / 100;

    var red = 255 * (1 - cc) * (1 - kk);
    var green = 255 * (1 - mm) * (1 - kk);
    var blue = 255 * (1 - yy) * (1 - kk);

    return rgbToHex(red, green, blue);
  }

  function grayToHex(gray) {
    // Imported grayscale artwork from Illustrator exposes darker values with higher gray percentages.
    var value = 255 * (1 - Math.max(0, Math.min(100, gray)) / 100);
    return rgbToHex(value, value, value);
  }

  function makeNoColor() {
    try {
      return new NoColor();
    } catch (eNoColor) {}
    return null;
  }

  function colorToHex(color) {
    if (!color) return null;

    var typename = "";
    try {
      typename = color.typename;
    } catch (eType) {}
    if (!typename) return null;

    if (typename === "NoColor") return null;
    if (typename === "PatternColor") return null;
    if (typename === "GradientColor") return null;

    if (typename === "RGBColor") {
      return rgbToHex(color.red, color.green, color.blue);
    }

    if (typename === "CMYKColor") {
      return cmykToHex(color.cyan, color.magenta, color.yellow, color.black);
    }

    if (typename === "GrayColor") {
      return grayToHex(color.gray);
    }

    if (typename === "SpotColor") {
      try {
        return colorToHex(color.spot.color);
      } catch (eSpot) {}
      return null;
    }

    return null;
  }

  function getItemTypename(item) {
    try {
      return item.typename;
    } catch (eType) {}
    return "";
  }

  function getDocumentSelectionItems(doc) {
    var items = [];
    var selection = null;

    try {
      selection = doc.selection;
    } catch (eSelection) {}

    if (!selection || !selection.length) return items;

    for (var i = 0; i < selection.length; i++) {
      items.push(selection[i]);
    }

    return items;
  }

  function getDirectItemFillHex(item, ignoreState) {
    if (!item) return null;

    if (!ignoreState) {
      try {
        if (item.filled === false) return null;
      } catch (eFilled) {}
    }

    try {
      return colorToHex(item.fillColor);
    } catch (eFill) {}

    return null;
  }

  function getDirectItemStrokeHex(item, ignoreState) {
    if (!item) return null;

    if (!ignoreState) {
      try {
        if (item.stroked === false) return null;
      } catch (eStroked) {}
    }

    try {
      return colorToHex(item.strokeColor);
    } catch (eStroke) {}

    return null;
  }

  function getCompoundChildStyleSource(item) {
    var childPaths = [];

    try {
      for (var i = 0; i < item.pathItems.length; i++) {
        childPaths.push(item.pathItems[i]);
      }
    } catch (eCompoundChildren) {}

    if (childPaths.length === 0) return item;

    childPaths.reverse();

    var backMostChild = childPaths[0];
    for (var j = 0; j < childPaths.length; j++) {
      var child = childPaths[j];
      if (getItemFillHex(child) !== null || getItemStrokeHex(child) !== null) {
        return child;
      }
    }

    return backMostChild;
  }

  function getItemFillHex(item) {
    var source = item;
    var typename = getItemTypename(source);

    if (typename === "TextFrame") {
      try {
        return colorToHex(source.textRange.characterAttributes.fillColor);
      } catch (eTextFill) {}
      return null;
    }

    if (typename === "CompoundPathItem") {
      try {
        if (source.filled !== false) {
          var compoundFillHex = colorToHex(source.fillColor);
          if (compoundFillHex !== null) return compoundFillHex;
        }
      } catch (eCompoundFillParent) {}
      source = getCompoundChildStyleSource(source);
      typename = getItemTypename(source);
    }

    try {
      if (source.filled === false) return null;
    } catch (eFilled) {}

    try {
      return colorToHex(source.fillColor);
    } catch (eFill) {}
    return null;
  }

  function getItemStrokeHex(item) {
    var source = item;
    var typename = getItemTypename(source);

    if (typename === "TextFrame") {
      try {
        return colorToHex(source.textRange.characterAttributes.strokeColor);
      } catch (eTextStroke) {}
      return null;
    }

    if (typename === "CompoundPathItem") {
      try {
        if (source.stroked !== false) {
          var compoundStrokeHex = colorToHex(source.strokeColor);
          if (compoundStrokeHex !== null) return compoundStrokeHex;
        }
      } catch (eCompoundStrokeParent) {}
      source = getCompoundChildStyleSource(source);
      typename = getItemTypename(source);
    }

    try {
      if (source.stroked === false) return null;
    } catch (eStroked) {}

    try {
      return colorToHex(source.strokeColor);
    } catch (eStroke) {}
    return null;
  }

  function isDrawableLeaf(item) {
    var typename = getItemTypename(item);
    return (
      typename === "PathItem" ||
      typename === "CompoundPathItem" ||
      typename === "TextFrame" ||
      typename === "RasterItem" ||
      typename === "PlacedItem" ||
      typename === "SymbolItem" ||
      typename === "MeshItem" ||
      typename === "PluginItem"
    );
  }

  function hasUsableRepresentativeColor(item) {
    return getItemFillHex(item) !== null || getItemStrokeHex(item) !== null;
  }

  function hasUsableLeafAppearance(item) {
    if (!item) return false;
    if (isClippingItem(item)) return false;
    if (!isDrawableLeaf(item)) return false;

    var typename = getItemTypename(item);
    if (
      typename === "RasterItem" ||
      typename === "PlacedItem" ||
      typename === "SymbolItem" ||
      typename === "MeshItem" ||
      typename === "PluginItem"
    ) {
      return true;
    }

    return hasUsableRepresentativeColor(item);
  }

  function hasVisiblePackAppearance(item) {
    if (!item) return false;

    try {
      if (item.hidden || item.locked) return false;
    } catch (eState) {}

    if (isContainerForRepresentativeLookup(item)) {
      var children = getRenderedChildrenInDirectOrder(item);
      for (var i = 0; i < children.length; i++) {
        if (hasVisiblePackAppearance(children[i])) return true;
      }
      return false;
    }

    return hasUsableLeafAppearance(item);
  }

  function isClippingItem(item) {
    try {
      return item.clipping === true;
    } catch (eClipping) {}
    return false;
  }

  function getItemZOrderPosition(item) {
    try {
      return item.zOrderPosition;
    } catch (eZOrder) {}
    return 0;
  }

  function compareZOrderPaths(pathA, pathB) {
    var maxLen = Math.max(pathA.length, pathB.length);
    for (var i = 0; i < maxLen; i++) {
      var a = i < pathA.length ? pathA[i] : -1;
      var b = i < pathB.length ? pathB[i] : -1;
      if (a !== b) return b - a;
    }
    return 0;
  }

  function isContainerForRepresentativeLookup(item) {
    var typename = getItemTypename(item);
    if (typename === "CompoundPathItem") return false;

    try {
      return item.pageItems && item.pageItems.length > 0;
    } catch (ePageItems) {}

    return false;
  }

  function getRenderedChildrenInDirectOrder(item) {
    var children = [];

    try {
      for (var i = 0; i < item.pageItems.length; i++) {
        children.push(item.pageItems[i]);
      }
    } catch (eChildren) {}

    children.reverse();

    return children;
  }

  function getFirstRenderedRepresentativeInfo(item, requireColor) {
    if (!item) return null;

    try {
      if (item.hidden) return null;
    } catch (eHidden) {}

    try {
      if (item.locked) return null;
    } catch (eLocked) {}

    if (isContainerForRepresentativeLookup(item)) {
      var children = getRenderedChildrenInDirectOrder(item);
      for (var i = 0; i < children.length; i++) {
        var childInfo = getFirstRenderedRepresentativeInfo(
          children[i],
          requireColor,
        );
        if (childInfo) return childInfo;
      }
      return null;
    }

    if (!isDrawableLeaf(item)) return null;

    var fillHex = getItemFillHex(item);
    var strokeHex = getItemStrokeHex(item);
    if (requireColor && fillHex === null && strokeHex === null) return null;

    return {
      item: item,
      fillHex: fillHex,
      strokeHex: strokeHex,
    };
  }

  function collectDrawableCandidates(item, zPath, out) {
    if (!item) return null;

    try {
      if (item.hidden) return;
    } catch (eHidden) {}

    var children = null;
    try {
      children = item.pageItems;
    } catch (eChildren) {}
    if (children && children.length) {
      for (var i = 0; i < children.length; i++) {
        var childPath = zPath.slice(0);
        childPath.push(getItemZOrderPosition(children[i]));
        collectDrawableCandidates(children[i], childPath, out);
      }
      return;
    }

    if (isDrawableLeaf(item) && !isClippingItem(item)) {
      out.push({
        item: item,
        zPath: zPath,
        hasColor: hasUsableRepresentativeColor(item),
      });
    }
  }

  function isPathLikeCandidate(item) {
    var typename = getItemTypename(item);
    return typename === "PathItem" || typename === "CompoundPathItem";
  }

  function getCandidateColorInfo(candidateItem) {
    var fillHex = getItemFillHex(candidateItem);
    if (fillHex !== null) {
      return {
        item: candidateItem,
        fillHex: fillHex,
        strokeHex: getItemStrokeHex(candidateItem),
      };
    }

    var strokeHex = getItemStrokeHex(candidateItem);
    if (strokeHex !== null) {
      return {
        item: candidateItem,
        fillHex: null,
        strokeHex: strokeHex,
      };
    }

    return null;
  }

  function getPathPriorityColorInfo(candidates) {
    var pathCandidates = [];
    var fallbackCandidates = [];

    for (var i = 0; i < candidates.length; i++) {
      if (isPathLikeCandidate(candidates[i].item))
        pathCandidates.push(candidates[i]);
      else fallbackCandidates.push(candidates[i]);
    }

    for (var pathIndex = 0; pathIndex < pathCandidates.length; pathIndex++) {
      var pathInfo = getCandidateColorInfo(pathCandidates[pathIndex].item);
      if (pathInfo) return pathInfo;
    }

    for (
      var fallbackIndex = 0;
      fallbackIndex < fallbackCandidates.length;
      fallbackIndex++
    ) {
      var fallbackInfo = getCandidateColorInfo(
        fallbackCandidates[fallbackIndex].item,
      );
      if (fallbackInfo) return fallbackInfo;
    }

    return {
      item: candidates[0].item,
      fillHex: null,
      strokeHex: null,
    };
  }

  function findBackMostDrawable(item) {
    var info = getFirstRenderedRepresentativeInfo(item, false);
    return info ? info.item : null;
  }

  function getBackMostDrawableColorInfo(item) {
    return getFirstRenderedRepresentativeInfo(item, true);
  }

  function clearCompoundPathFill(item, noColor) {
    if (!item) return;

    try {
      item.filled = false;
    } catch (eCompoundFilledFalse) {}
    try {
      if (noColor) item.fillColor = noColor;
    } catch (eCompoundNoFillColor) {}

    var childPaths = [];
    try {
      for (var i = 0; i < item.pathItems.length; i++) {
        childPaths.push(item.pathItems[i]);
      }
    } catch (eCompoundChildren) {}

    for (var childIndex = 0; childIndex < childPaths.length; childIndex++) {
      var child = childPaths[childIndex];
      if (getItemFillHex(child) === null || getItemStrokeHex(child) === null)
        continue;

      try {
        child.filled = false;
      } catch (eChildFilledFalse) {}
      try {
        if (noColor) child.fillColor = noColor;
      } catch (eChildNoFillColor) {}
    }
  }

  function isRedStrokeOffEligibleItem(item) {
    var typename = getItemTypename(item);
    return typename === "PathItem" || typename === "CompoundPathItem";
  }

  function getDirectComparableStrokeHex(item) {
    var strokeHex = getDirectItemStrokeHex(item, false);
    if (strokeHex === null) strokeHex = getDirectItemStrokeHex(item, true);
    return normalizeHex(strokeHex);
  }

  function getComparableRedStrokeHex(item) {
    if (!item) return null;

    if (getItemTypename(item) === "CompoundPathItem") {
      return normalizeHex(getItemStrokeHex(item));
    }

    return getDirectComparableStrokeHex(item);
  }

  function clearCompoundPathRedStroke(item, noColor) {
    if (!item) return;

    // Illustrator may store visible compound stroke on the parent or child paths.
    try {
      item.stroked = false;
    } catch (eCompoundStrokedFalse) {}
    try {
      if (noColor) item.strokeColor = noColor;
    } catch (eCompoundNoStrokeColor) {}

    var childPaths = [];
    try {
      for (var i = 0; i < item.pathItems.length; i++) {
        childPaths.push(item.pathItems[i]);
      }
    } catch (eCompoundStrokeChildren) {}

    for (var childIndex = 0; childIndex < childPaths.length; childIndex++) {
      var child = childPaths[childIndex];
      try {
        child.stroked = false;
      } catch (eChildStrokedFalse) {}
      try {
        if (noColor) child.strokeColor = noColor;
      } catch (eChildNoStrokeColor) {}
    }
  }

  function isStandaloneRedStrokeTarget(item, groupType) {
    if (!item) return false;
    if (groupType !== TYPE_3MM) return false;
    if (getItemTypename(item.parent) !== "Layer") return false;
    if (!isRedStrokeOffEligibleItem(item)) return false;

    return getComparableRedStrokeHex(item) === RED_STROKE_OFF_HEX;
  }

  function collectGroupPathStrokeStatus(item, status) {
    if (!item) return;

    try {
      if (item.hidden || item.locked) return;
    } catch (eState) {}

    if (isContainerForRepresentativeLookup(item)) {
      var children = [];
      try {
        for (var i = 0; i < item.pageItems.length; i++) {
          children.push(item.pageItems[i]);
        }
      } catch (eChildren) {}

      for (var childIndex = 0; childIndex < children.length; childIndex++) {
        collectGroupPathStrokeStatus(children[childIndex], status);
      }
      return;
    }

    if (!isRedStrokeOffEligibleItem(item)) return;

    status.hasLeaf = true;
    if (getComparableRedStrokeHex(item) !== RED_STROKE_OFF_HEX) {
      status.allRed = false;
    }
  }

  function isAllRedStrokeSubtree(item) {
    if (!item) return false;
    if (!isContainerForRepresentativeLookup(item)) return false;

    var status = { hasLeaf: false, allRed: true };
    collectGroupPathStrokeStatus(item, status);
    return status.hasLeaf && status.allRed;
  }

  function isAllGroupPathLayersRedStroke(item, groupType) {
    if (!item) return false;
    if (groupType !== TYPE_3MM) return false;
    return isAllRedStrokeSubtree(item);
  }

  function getFirstRenderedDirectChildSubtree(item) {
    if (!item) return null;
    if (!isContainerForRepresentativeLookup(item)) return null;

    var children = getRenderedChildrenInDirectOrder(item);
    for (var i = 0; i < children.length; i++) {
      var child = children[i];
      var childInfo = getFirstRenderedRepresentativeInfo(child, false);
      if (childInfo) return child;
    }

    return null;
  }

  function isFirstRenderedSubgroupAllRedStroke(item, groupType) {
    if (!item) return false;
    if (groupType !== TYPE_3MM) return false;
    if (!isContainerForRepresentativeLookup(item)) return false;

    var firstSubtree = getFirstRenderedDirectChildSubtree(item);
    if (!firstSubtree) return false;
    return isAllRedStrokeSubtree(firstSubtree);
  }

  function isFirstRenderedDirectChildRedLeaf(item, groupType) {
    if (!item) return false;
    if (groupType !== TYPE_3MM) return false;
    if (!isContainerForRepresentativeLookup(item)) return false;

    var firstChild = getFirstRenderedDirectChildSubtree(item);
    if (!firstChild) return false;
    if (isContainerForRepresentativeLookup(firstChild)) return false;
    if (!isRedStrokeOffEligibleItem(firstChild)) return false;

    return getComparableRedStrokeHex(firstChild) === RED_STROKE_OFF_HEX;
  }

  function hasRedStrokeTarget(item, groupType) {
    if (!item) return false;
    return (
      isStandaloneRedStrokeTarget(item, groupType) ||
      isAllGroupPathLayersRedStroke(item, groupType)
    );
  }

  function applyMixedGroupLasercutStyling(item, groupType) {
    if (!item) return false;
    if (groupType !== TYPE_3MM) return false;
    if (!isContainerForRepresentativeLookup(item)) return false;

    var firstChild = getFirstRenderedDirectChildSubtree(item);
    var useFirstRenderedRedSubgroup =
      isFirstRenderedSubgroupAllRedStroke(item, groupType);
    var useFirstRenderedRedLeaf = isFirstRenderedDirectChildRedLeaf(
      item,
      groupType,
    );
    if (!useFirstRenderedRedSubgroup && !useFirstRenderedRedLeaf) return false;
    if (!firstChild) return false;

    var children = getRenderedChildrenInDirectOrder(item);
    var styled = false;

    for (var i = 0; i < children.length; i++) {
      var child = children[i];
      var childInfo = getFirstRenderedRepresentativeInfo(child, false);
      if (!childInfo) continue;

      if (child === firstChild) {
        removeRedStrokeFromPackedItem(child, groupType);
      } else {
        removeFillFromPackedItem(child);
      }
      styled = true;
    }

    return styled;
  }

  function removeRedStrokeFromPackedItem(item, groupType) {
    if (!item) return;
    if (groupType !== TYPE_3MM) return;

    try {
      if (item.hidden || item.locked) return;
    } catch (eState) {}

    if (isContainerForRepresentativeLookup(item)) {
      var children = [];
      try {
        for (var i = 0; i < item.pageItems.length; i++) {
          children.push(item.pageItems[i]);
        }
      } catch (eChildren) {}

      for (var childIndex = 0; childIndex < children.length; childIndex++) {
        removeRedStrokeFromPackedItem(children[childIndex], groupType);
      }
      return;
    }

    var typename = getItemTypename(item);
    if (!isRedStrokeOffEligibleItem(item)) return;

    var noColor = makeNoColor();

    if (typename === "CompoundPathItem") {
      clearCompoundPathRedStroke(item, noColor);
      return;
    }

    if (getDirectComparableStrokeHex(item) !== RED_STROKE_OFF_HEX) return;

    try {
      item.stroked = false;
    } catch (eStrokedFalse) {}
    try {
      if (noColor) item.strokeColor = noColor;
    } catch (eNoStrokeColor) {}
  }

  function turnOffFillIfFilledAndStroked(item) {
    if (!item) return;

    var fillHex = getItemFillHex(item);
    var strokeHex = getItemStrokeHex(item);
    if (fillHex === null || strokeHex === null) return;

    var typename = getItemTypename(item);
    var noColor = makeNoColor();

    if (typename === "TextFrame") {
      try {
        if (noColor) item.textRange.characterAttributes.fillColor = noColor;
      } catch (eTextNoFill) {}
      return;
    }

    if (typename === "CompoundPathItem") {
      clearCompoundPathFill(item, noColor);
      return;
    }

    try {
      item.filled = false;
    } catch (eFilledFalse) {}
    try {
      if (noColor) item.fillColor = noColor;
    } catch (eNoFillColor) {}
  }

  function removeFillFromPackedItem(item) {
    if (!item) return;

    try {
      if (item.hidden || item.locked) return;
    } catch (eState) {}

    if (isContainerForRepresentativeLookup(item)) {
      var children = [];
      try {
        for (var i = 0; i < item.pageItems.length; i++) {
          children.push(item.pageItems[i]);
        }
      } catch (eChildren) {}

      for (var childIndex = 0; childIndex < children.length; childIndex++) {
        removeFillFromPackedItem(children[childIndex]);
      }
      return;
    }

    if (!isDrawableLeaf(item)) return;
    turnOffFillIfFilledAndStroked(item);
  }

  function hasUsablePackAppearance(item) {
    return hasVisiblePackAppearance(item);
  }

  function colorsMatchWithinTolerance(rgbA, rgbB, tolerance) {
    return (
      Math.abs(rgbA.r - rgbB.r) <= tolerance &&
      Math.abs(rgbA.g - rgbB.g) <= tolerance &&
      Math.abs(rgbA.b - rgbB.b) <= tolerance
    );
  }

  function matchPaletteColor(rawHex) {
    var rawRgb = hexToRgb(rawHex);
    if (!rawRgb) return null;

    var bestMatch = null;
    var bestDistance = Number.MAX_VALUE;

    for (var i = 0; i < COLOR_PALETTE.length; i++) {
      if (
        colorsMatchWithinTolerance(
          rawRgb,
          COLOR_PALETTE[i].rgb,
          COLOR_TOLERANCE,
        )
      ) {
        var paletteRgb = COLOR_PALETTE[i].rgb;
        var dr = rawRgb.r - paletteRgb.r;
        var dg = rawRgb.g - paletteRgb.g;
        var db = rawRgb.b - paletteRgb.b;
        var distance = dr * dr + dg * dg + db * db;

        if (distance < bestDistance) {
          bestDistance = distance;
          bestMatch = COLOR_PALETTE[i];
        }
      }
    }

    return bestMatch;
  }

  function getTypeNameForStrokeHex(strokeHex) {
    return normalizeHex(strokeHex) === TYPE_5MM_STROKE_HEX
      ? TYPE_5MM
      : TYPE_3MM;
  }

  function compareText(a, b) {
    if (a < b) return -1;
    if (a > b) return 1;
    return 0;
  }

  function getItemGroupInfo(item) {
    var colorInfo = getBackMostDrawableColorInfo(item);
    var representative = colorInfo
      ? colorInfo.item
      : findBackMostDrawable(item);
    if (!representative) representative = item;

    var fillHex = colorInfo
      ? colorInfo.fillHex
      : getItemFillHex(representative);
    var strokeHex = colorInfo
      ? colorInfo.strokeHex
      : getItemStrokeHex(representative);
    var rawHex = fillHex || strokeHex;
    var matched = rawHex ? matchPaletteColor(rawHex) : null;

    var colorName = "Unknown";
    var colorHex = "#000000";
    var colorSortIndex = COLOR_PALETTE.length + 1;
    var colorSortKey = "ZZZZUNKNOWN";

    if (matched) {
      colorName = matched.name || matched.label;
      colorHex = matched.hex;
      colorSortIndex = matched.sortIndex;
      colorSortKey = matched.label;
    } else if (rawHex) {
      rawHex = normalizeHex(rawHex);
      colorName = rawHex;
      colorHex = rawHex;
      colorSortIndex = COLOR_PALETTE.length;
      colorSortKey = rawHex;
    }

    var groupType = getTypeNameForStrokeHex(strokeHex);

    return {
      type: groupType,
      colorName: colorName,
      colorHex: colorHex,
      key: groupType + "|" + colorName,
      labelText: groupType + " - " + colorName,
      typeSortIndex: TYPE_SORT_ORDER[groupType],
      colorSortIndex: colorSortIndex,
      colorSortKey: colorSortKey,
    };
  }

  function compareCollectedItems(a, b) {
    if (a.groupTypeSortIndex !== b.groupTypeSortIndex) {
      return a.groupTypeSortIndex - b.groupTypeSortIndex;
    }

    if (a.groupColorSortIndex !== b.groupColorSortIndex) {
      return a.groupColorSortIndex - b.groupColorSortIndex;
    }

    var textCompare = compareText(a.groupColorSortKey, b.groupColorSortKey);
    if (textCompare !== 0) return textCompare;

    textCompare = compareText(a.groupKey, b.groupKey);
    if (textCompare !== 0) return textCompare;

    return a.collectIndex - b.collectIndex;
  }

  function compareByCellSize(a, b) {
    if (a.cw !== b.cw) return a.cw - b.cw;
    if (a.ch !== b.ch) return a.ch - b.ch;
    return 0;
  }

  function compareSmallBucketItems(a, b) {
    var cellCompare = compareByCellSize(a, b);
    if (cellCompare !== 0) return cellCompare;

    if (a.heightCells !== b.heightCells) return a.heightCells - b.heightCells;
    if (a.area !== b.area) return a.area - b.area;
    return a.collectIndex - b.collectIndex;
  }

  function compareDenseGroupItems(a, b) {
    var cellCompare = compareByCellSize(a, b);
    if (cellCompare !== 0) return cellCompare;

    if (a.longestSideCells !== b.longestSideCells) {
      return b.longestSideCells - a.longestSideCells;
    }
    if (a.area !== b.area) return b.area - a.area;
    if (a.heightCells !== b.heightCells) return b.heightCells - a.heightCells;
    return a.collectIndex - b.collectIndex;
  }

  function normalizeFsPathText(pathText) {
    if (!pathText) return "";

    var text = String(pathText).replace(/\//g, "\\");
    while (text.length > 3 && text.charAt(text.length - 1) === "\\") {
      text = text.substring(0, text.length - 1);
    }
    return text.toLowerCase();
  }

  function getDocumentFsPath(docRef) {
    try {
      return normalizeFsPathText(docRef.fullName.fsName);
    } catch (eDocPath) {}
    return "";
  }

  function isFolderInsideMain(chosenFolder, mainFolder) {
    var chosenPath = normalizeFsPathText(chosenFolder.fsName);
    var mainPath = normalizeFsPathText(mainFolder.fsName);
    if (!chosenPath || !mainPath) return false;
    return chosenPath !== mainPath && chosenPath.indexOf(mainPath + "\\") === 0;
  }

  function getFolderModifiedTime(folder) {
    var modified = null;
    try {
      modified = folder.modified;
    } catch (eModified) {}
    if (!modified) return 0;

    try {
      var time = modified.getTime();
      return isNaN(time) ? 0 : time;
    } catch (eGetTime) {}
    return 0;
  }

  function pad2(value) {
    value = String(value);
    return value.length < 2 ? "0" + value : value;
  }

  function formatFolderModifiedDate(folder) {
    var modified = null;
    try {
      modified = folder.modified;
    } catch (eModified) {}
    if (!modified) return "Unknown";

    try {
      return (
        modified.getFullYear() +
        "-" +
        pad2(modified.getMonth() + 1) +
        "-" +
        pad2(modified.getDate()) +
        " " +
        pad2(modified.getHours()) +
        ":" +
        pad2(modified.getMinutes())
      );
    } catch (eFormatModified) {}
    return "Unknown";
  }

  function getReadableFolderName(folder) {
    var text = "";
    try {
      text = String(folder.name);
    } catch (eFolderName) {}
    if (!text) return "";

    try {
      if (/%[0-9A-Fa-f]{2}/.test(text)) return decodeURIComponent(text);
    } catch (eDecodeFolderName) {}

    return text.replace(/%20/gi, " ");
  }

  function getDirectChildFolders(mainFolder) {
    var childFolders = mainFolder.getFiles(function (entry) {
      return entry instanceof Folder;
    });

    var folderEntries = [];
    for (var i = 0; i < childFolders.length; i++) {
      folderEntries.push({
        folder: childFolders[i],
        modifiedTime: getFolderModifiedTime(childFolders[i]),
        modifiedText: formatFolderModifiedDate(childFolders[i]),
      });
    }

    folderEntries.sort(function (a, b) {
      if (a.modifiedTime !== b.modifiedTime)
        return b.modifiedTime - a.modifiedTime;
      return compareText(
        String(a.folder.name).toLowerCase(),
        String(b.folder.name).toLowerCase(),
      );
    });

    return folderEntries;
  }

  function chooseFolderFromEntries(dialogTitle, infoText, folderEntries) {
    var chosenFolder = null;

    try {
      var picker = new Window("dialog", dialogTitle);
      picker.orientation = "column";
      picker.alignChildren = "fill";

      var info = picker.add("statictext", undefined, infoText, {
        multiline: true,
      });
      info.alignment = "fill";

      var list = picker.add("listbox", undefined, [], {
        multiselect: false,
      });
      list.preferredSize = [720, 320];

      for (var i = 0; i < folderEntries.length; i++) {
        var entry = folderEntries[i];
        var item = list.add(
          "item",
          getReadableFolderName(entry.folder) +
            "    [" +
            entry.modifiedText +
            "]",
        );
        item.folderRef = entry.folder;
      }

      if (list.items.length > 0) list.selection = 0;

      var buttons = picker.add("group");
      buttons.alignment = "right";
      var ok = buttons.add("button", undefined, "OK");
      var cancel = buttons.add("button", undefined, "Cancel");

      ok.onClick = function () {
        if (!list.selection) {
          alert("Choose a folder.");
          return;
        }
        chosenFolder = list.selection.folderRef;
        picker.close(1);
      };
      cancel.onClick = function () {
        picker.close(0);
      };
      list.onDoubleClick = function () {
        if (!list.selection) return;
        chosenFolder = list.selection.folderRef;
        picker.close(1);
      };

      if (picker.show() !== 1 || !chosenFolder) return null;
    } catch (ePickerUI) {
      alert("Folder chooser UI failed.");
      return null;
    }

    return chosenFolder;
  }

  function chooseSourceFolderUnderMain(mainFolder) {
    if (!mainFolder.exists) {
      alert("Main folder not found:\n" + MAIN_FOLDER_PATH);
      return null;
    }

    var topLevelEntries = getDirectChildFolders(mainFolder);
    if (topLevelEntries.length === 0) {
      alert("No child folders found under:\n" + mainFolder.fsName);
      return null;
    }

    var parentFolder = chooseFolderFromEntries(
      "Choose Source Folder",
      "Folders under:\n" + mainFolder.fsName,
      topLevelEntries,
    );
    if (!parentFolder) return null;

    if (!isFolderInsideMain(parentFolder, mainFolder)) {
      alert("Choose a folder inside:\n" + mainFolder.fsName);
      return null;
    }

    var childEntries = getDirectChildFolders(parentFolder);
    if (childEntries.length === 0) {
      alert("No child folders found under:\n" + parentFolder.fsName);
      return null;
    }

    var chosenFolder = chooseFolderFromEntries(
      "Choose Source Subfolder",
      "Subfolders under:\n" + parentFolder.fsName,
      childEntries,
    );
    if (!chosenFolder) return null;

    if (!isFolderInsideMain(chosenFolder, mainFolder)) {
      alert("Choose a folder inside:\n" + mainFolder.fsName);
      return null;
    }

    return chosenFolder;
  }

  function getDirectAiFiles(folder) {
    var aiFiles = folder.getFiles(function (entry) {
      return entry instanceof File && /\.ai$/i.test(entry.name);
    });

    aiFiles.sort(function (a, b) {
      return compareText(
        String(a.name).toLowerCase(),
        String(b.name).toLowerCase(),
      );
    });
    return aiFiles;
  }

  function findOpenDocumentByFile(fileRef) {
    var targetPath = normalizeFsPathText(fileRef.fsName);

    for (var i = 0; i < app.documents.length; i++) {
      if (getDocumentFsPath(app.documents[i]) === targetPath)
        return app.documents[i];
    }
    return null;
  }

  function isTopLevelImportItem(item) {
    return getItemTypename(item.parent) === "Layer";
  }

  function collectPackableSourceItems(sourceDoc) {
    var collected = [];

    for (var i = 0; i < sourceDoc.pageItems.length; i++) {
      var sourceItem = sourceDoc.pageItems[i];
      if (!isTopLevelImportItem(sourceItem)) continue;
      if (!isPackable(sourceItem)) continue;
      collected.push(sourceItem);
    }

    return collected;
  }

  function importItemsFromFolder(targetDoc, sourceFolder) {
    var result = {
      items: [],
      processedFileCount: 0,
      skippedFileCount: 0,
      importedCount: 0,
      aiFileCount: 0,
      sourceFolderPath: sourceFolder.fsName,
    };

    var aiFiles = getDirectAiFiles(sourceFolder);
    result.aiFileCount = aiFiles.length;

    var targetLayer = targetDoc.activeLayer;
    var targetDocPath = getDocumentFsPath(targetDoc);

    for (var i = 0; i < aiFiles.length; i++) {
      var aiFile = aiFiles[i];
      if (
        targetDocPath &&
        normalizeFsPathText(aiFile.fsName) === targetDocPath
      ) {
        result.skippedFileCount++;
        continue;
      }

      var sourceDoc = null;
      var openedHere = false;
      var importedFromFile = 0;

      try {
        sourceDoc = findOpenDocumentByFile(aiFile);
        if (!sourceDoc) {
          sourceDoc = app.open(aiFile);
          openedHere = true;
        }

        var sourceItems = collectPackableSourceItems(sourceDoc);
        if (sourceItems.length === 0) {
          result.skippedFileCount++;
          continue;
        }

        for (var itemIndex = 0; itemIndex < sourceItems.length; itemIndex++) {
          try {
            var dup = sourceItems[itemIndex].duplicate(
              targetLayer,
              ElementPlacement.PLACEATEND,
            );
            if (dup && isPackable(dup)) {
              result.items.push(dup);
              importedFromFile++;
            }
          } catch (eDuplicate) {}
        }

        if (importedFromFile > 0) {
          result.processedFileCount++;
          result.importedCount += importedFromFile;
        } else {
          result.skippedFileCount++;
        }
      } catch (eImportFile) {
        result.skippedFileCount++;
      } finally {
        if (sourceDoc && openedHere) {
          try {
            sourceDoc.close(SaveOptions.DONOTSAVECHANGES);
          } catch (eCloseSource) {}
        }
      }
    }

    return result;
  }

  for (
    var paletteIndex = 0;
    paletteIndex < COLOR_PALETTE.length;
    paletteIndex++
  ) {
    COLOR_PALETTE[paletteIndex].hex = normalizeHex(
      COLOR_PALETTE[paletteIndex].hex,
    );
    COLOR_PALETTE[paletteIndex].rgb = hexToRgb(COLOR_PALETTE[paletteIndex].hex);
    COLOR_PALETTE[paletteIndex].sortIndex = paletteIndex;
  }

  // ----------------------------
  // Bounds helpers
  // ----------------------------
  // bounds: [L, T, R, B]
  function getUnionBounds(item) {
    var gb = null,
      vb = null;
    try {
      gb = item.geometricBounds;
    } catch (e1) {}
    try {
      vb = item.visibleBounds;
    } catch (e2) {}

    if (!gb && !vb) return null;
    if (!gb) return vb;
    if (!vb) return gb;

    return [
      Math.min(gb[0], vb[0]),
      Math.max(gb[1], vb[1]),
      Math.max(gb[2], vb[2]),
      Math.min(gb[3], vb[3]),
    ];
  }

  function isCompoundClippingItem(item) {
    if (getItemTypename(item) !== "CompoundPathItem") return false;

    try {
      for (var i = 0; i < item.pathItems.length; i++) {
        if (item.pathItems[i].clipping === true) return true;
      }
    } catch (eCompoundClip) {}

    return false;
  }

  function getClippingMaskBounds(item) {
    var isClipped = false;
    try {
      isClipped = item.clipped === true;
    } catch (eClipped) {}
    if (!isClipped) return null;

    var children = null;
    try {
      children = item.pageItems;
    } catch (eChildren) {}
    if (!children || !children.length) return null;

    for (var i = 0; i < children.length; i++) {
      var child = children[i];
      if (!isClippingItem(child) && !isCompoundClippingItem(child)) continue;

      try {
        return child.geometricBounds;
      } catch (eMaskBounds) {}
    }

    return null;
  }

  function unionBoundsArrays(boundsA, boundsB) {
    if (!boundsA) return boundsB;
    if (!boundsB) return boundsA;

    return [
      Math.min(boundsA[0], boundsB[0]),
      Math.max(boundsA[1], boundsB[1]),
      Math.max(boundsA[2], boundsB[2]),
      Math.min(boundsA[3], boundsB[3]),
    ];
  }

  function getPackingBounds(item) {
    var maskBounds = getClippingMaskBounds(item);
    if (maskBounds) return maskBounds;

    var children = null;
    try {
      children = item.pageItems;
    } catch (eChildren) {}

    if (children && children.length) {
      var mergedBounds = null;

      for (var i = 0; i < children.length; i++) {
        var child = children[i];

        try {
          if (child.hidden || child.locked) continue;
        } catch (eChildState) {}

        mergedBounds = unionBoundsArrays(mergedBounds, getPackingBounds(child));
      }

      if (mergedBounds) return mergedBounds;
    }

    return getUnionBounds(item);
  }

  function boundsWidthPt(b) {
    return b[2] - b[0];
  }
  function boundsHeightPt(b) {
    return b[1] - b[3];
  }

  // Move bounds MUST match sizing bounds for consistent spacing
  function getMoveBounds(item) {
    return getPackingBounds(item);
  }

  // ----------------------------
  // Free-rectangle packer helpers
  // ----------------------------
  function initFreeRects(startY) {
    if (startY === undefined || startY === null) startY = 0;
    if (startY >= GH) return [];
    return [{ x: 0, y: startY, w: GW, h: GH - startY }];
  }

  function comparePlacementScores(a, b) {
    if (a.leftoverShort !== b.leftoverShort) {
      return a.leftoverShort - b.leftoverShort;
    }

    if (a.leftoverLong !== b.leftoverLong) {
      return a.leftoverLong - b.leftoverLong;
    }

    if (a.y !== b.y) {
      return a.y - b.y;
    }

    return a.x - b.x;
  }

  function findBestRectPlacement(freeRects, cw, ch) {
    var best = null;

    for (var i = 0; i < freeRects.length; i++) {
      var rect = freeRects[i];
      if (cw > rect.w || ch > rect.h) continue;

      var leftoverW = rect.w - cw;
      var leftoverH = rect.h - ch;
      var candidate = {
        x: rect.x,
        y: rect.y,
        leftoverShort: Math.min(leftoverW, leftoverH),
        leftoverLong: Math.max(leftoverW, leftoverH),
        freeRect: rect,
        placedRect: { x: rect.x, y: rect.y, w: cw, h: ch },
      };

      if (!best || comparePlacementScores(candidate, best) < 0) {
        best = candidate;
      }
    }

    return best;
  }

  function findNextFittingRegularIndex(regularItems, startIndex, freeRects) {
    for (var i = startIndex; i < regularItems.length; i++) {
      var obj = regularItems[i];
      if (findBestRectPlacement(freeRects, obj.cw, obj.ch)) return i;
    }
    return -1;
  }

  function rectContainsRect(a, b) {
    return (
      a.x <= b.x &&
      a.y <= b.y &&
      a.x + a.w >= b.x + b.w &&
      a.y + a.h >= b.y + b.h
    );
  }

  function pruneContainedFreeRects(freeRects) {
    var pruned = [];

    for (var i = 0; i < freeRects.length; i++) {
      var rect = freeRects[i];
      var contained = false;

      for (var j = 0; j < freeRects.length; j++) {
        if (i === j) continue;
        if (rectContainsRect(freeRects[j], rect)) {
          contained = true;
          break;
        }
      }

      if (!contained) pruned.push(rect);
    }

    return pruned;
  }

  function splitFreeRects(freeRects, chosenRect, placedRect) {
    var nextFreeRects = [];

    for (var i = 0; i < freeRects.length; i++) {
      if (freeRects[i] !== chosenRect) nextFreeRects.push(freeRects[i]);
    }

    var rightW = chosenRect.w - placedRect.w;
    if (rightW > 0) {
      nextFreeRects.push({
        x: chosenRect.x + placedRect.w,
        y: chosenRect.y,
        w: rightW,
        h: placedRect.h,
      });
    }

    var topH = chosenRect.h - placedRect.h;
    if (topH > 0) {
      nextFreeRects.push({
        x: chosenRect.x,
        y: chosenRect.y + placedRect.h,
        w: chosenRect.w,
        h: topH,
      });
    }

    return pruneContainedFreeRects(nextFreeRects);
  }

  function packSmallItemsInRows(smallItems, startIndex) {
    var placements = [];
    var rowX = 0;
    var rowY = 0;
    var rowHeight = 0;
    var index = startIndex;

    while (index < smallItems.length) {
      var obj = smallItems[index];

      if (rowHeight === 0) {
        if (rowY + obj.ch > GH) break;
        placements.push({ item: obj, x: rowX, y: rowY });
        rowX += obj.cw;
        rowHeight = obj.ch;
        index++;
        continue;
      }

      if (rowX + obj.cw <= GW) {
        placements.push({ item: obj, x: rowX, y: rowY });
        rowX += obj.cw;
        if (obj.ch > rowHeight) rowHeight = obj.ch;
        index++;
        continue;
      }

      rowY += rowHeight;
      rowX = 0;
      rowHeight = 0;
    }

    return {
      placements: placements,
      nextIndex: index,
      usedHeight: rowHeight > 0 ? rowY + rowHeight : rowY,
    };
  }

  // ----------------------------
  // Drawing helpers (optional)
  // ----------------------------
  function createBoxAt(boxLeftPt, boxTopPt) {
    var rect = doc.pathItems.rectangle(
      boxTopPt,
      boxLeftPt,
      cmToPt(BOX),
      cmToPt(BOX),
    );
    rect.stroked = true;
    try {
      var boxStrokeColor = makeRgbColorFromHex("#000000");
      if (boxStrokeColor) rect.strokeColor = boxStrokeColor;
    } catch (eBoxStroke) {}
    rect.filled = false;
    return rect;
  }

  function getBestMatchingFont(fontName) {
    var font = null;
    try {
      font = app.textFonts.getByName(fontName);
    } catch (eExact) {}
    if (font) return font;

    var needle = String(fontName).toLowerCase();
    var bestFont = null;
    var bestScore = -1;

    for (var i = 0; i < app.textFonts.length; i++) {
      var candidate = app.textFonts[i];
      var name = "";
      var family = "";
      var style = "";

      try {
        name = String(candidate.name).toLowerCase();
      } catch (eName) {}
      try {
        family = String(candidate.family).toLowerCase();
      } catch (eFamily) {}
      try {
        style = String(candidate.style).toLowerCase();
      } catch (eStyle) {}

      if (name.indexOf(needle) === -1 && family.indexOf(needle) === -1)
        continue;

      var score = 0;
      if (name === needle) score += 100;
      if (family === needle) score += 80;
      if (name.indexOf(needle) !== -1) score += 20;
      if (family.indexOf(needle) !== -1) score += 20;
      if (style === "regular" || style === "roman" || style === "normal")
        score += 10;

      if (score > bestScore) {
        bestScore = score;
        bestFont = candidate;
      }
    }

    return bestFont;
  }

  function createCenteredLabel(labelText, labelColorHex, boxLeftPt, boxTopPt) {
    var labelY = boxTopPt - cmToPt(BOX + LABEL_OFFSET);
    var label = doc.textFrames.pointText([boxLeftPt, labelY]);
    label.contents = labelText;

    try {
      var labelRange = label.textRange;
      var labelFont = getBestMatchingFont(LABEL_FONT_NAME);
      if (labelFont) labelRange.characterAttributes.textFont = labelFont;
      labelRange.characterAttributes.size = LABEL_FONT_SIZE;
      var labelFillColor = makeRgbColorFromHex(labelColorHex);
      if (labelFillColor)
        labelRange.characterAttributes.fillColor = labelFillColor;
    } catch (eTextStyle) {}

    try {
      var b = label.visibleBounds;
      var textCenterX = (b[0] + b[2]) / 2;
      var boxCenterX = boxLeftPt + cmToPt(BOX) / 2;
      label.translate(boxCenterX - textCenterX, 0);

      var b2 = label.visibleBounds;
      var textCenterX2 = (b2[0] + b2[2]) / 2;
      label.translate(boxCenterX - textCenterX2, 0);
    } catch (e) {}
    return label;
  }

  // ----------------------------
  // Move item so its (union) bottom-left lands on target
  // ----------------------------
  function moveItemBottomLeftTo(item, targetBLxPt, targetBLyPt) {
    var b = getMoveBounds(item);
    if (!b) return false;

    try {
      item.translate(targetBLxPt - b[0], targetBLyPt - b[3]);
    } catch (eMove1) {
      return false;
    }

    // corrective pass for rounding
    var b2 = getMoveBounds(item);
    if (b2) {
      try {
        item.translate(targetBLxPt - b2[0], targetBLyPt - b2[3]);
      } catch (eMove2) {
        return false;
      }
    }

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

  var workflowModeLabel =
    workflowMode === WORKFLOW_PRINT ? "Pack for print" : "Pack for lasercut";
  var sourceModeLabel =
    sourceMode === SOURCE_FOLDER ? "Folder import" : "Current selection";
  var keepFillAfterPacking = workflowMode === WORKFLOW_PRINT;
  var processedFileCount = 0;
  var skippedFileCount = 0;
  var importedCount = 0;
  var aiFileCount = 0;
  var sourceFolderPath = "";
  var inputItems = null;

  if (sourceMode === SOURCE_SELECTION) {
    // Freeze the current selection so temporary probe selection changes
    // cannot leak extra items into the packing input set.
    inputItems = getDocumentSelectionItems(doc);
    if (!inputItems || inputItems.length === 0) {
      alert(
        "Select the objects (groups are OK) you want to pack, then run again.",
      );
      return;
    }
  } else {
    var mainFolder = new Folder(MAIN_FOLDER_PATH);
    var chosenFolder = chooseSourceFolderUnderMain(mainFolder);
    if (!chosenFolder) return;

    var importResult = importItemsFromFolder(doc, chosenFolder);
    processedFileCount = importResult.processedFileCount;
    skippedFileCount = importResult.skippedFileCount;
    importedCount = importResult.importedCount;
    aiFileCount = importResult.aiFileCount;
    sourceFolderPath = importResult.sourceFolderPath;
    inputItems = importResult.items;

    if (!inputItems || inputItems.length === 0) {
      alert(
        "No packable items were imported.\n\n" +
          "Folder: " +
          sourceFolderPath +
          "\n" +
          ".ai files found: " +
          aiFileCount +
          "\n" +
          "Files processed: " +
          processedFileCount +
          "\n" +
          "Files skipped: " +
          skippedFileCount,
      );
      return;
    }

    if (workflowMode === WORKFLOW_PRINT) {
      try {
        doc.selection = null;
      } catch (eClearSelection) {}

      for (
        var importedIndex = 0;
        importedIndex < importResult.items.length;
        importedIndex++
      ) {
        try {
          importResult.items[importedIndex].selected = true;
        } catch (eSelectImported) {}
      }

      alert(
        "Imported items are selected.\n\n" +
          "Choose what you want to pack, then rerun:\n" +
          "Pack for print -> Pack current selection\n\n" +
          "Folder: " +
          sourceFolderPath +
          "\n" +
          "Imported objects: " +
          importedCount +
          "\n" +
          "Files processed: " +
          processedFileCount +
          "\n" +
          "Files skipped: " +
          skippedFileCount,
      );
      return;
    }
  }

  doDraw = chooseDrawMode();
  if (doDraw === null) return;

  var items = [];
  for (var i = 0; i < inputItems.length; i++) {
    var it = inputItems[i];
    if (!isPackable(it)) continue;
    if (!hasUsablePackAppearance(it)) continue;

    var pb = getPackingBounds(it);
    var wCm = ptToCm(boundsWidthPt(pb));
    var hCm = ptToCm(boundsHeightPt(pb));

    var wPad = wCm + 2 * OBJ_PAD;
    var hPad = hCm + 2 * OBJ_PAD;

    var cw = Math.ceil(wPad / CELL);
    var ch = Math.ceil(hPad / CELL);
    var groupInfo = getItemGroupInfo(it);

    items.push({
      item: it,
      collectIndex: i,
      area: wCm * hCm,
      cw: cw,
      ch: ch,
      longestSideCells: Math.max(cw, ch),
      heightCells: ch,
      isSmallBucket: Math.max(cw, ch) <= SMALL_BUCKET_MAX_CELLS,
      groupType: groupInfo.type,
      groupColorName: groupInfo.colorName,
      groupColorHex: groupInfo.colorHex,
      groupKey: groupInfo.key,
      groupLabelText: groupInfo.labelText,
      groupTypeSortIndex: groupInfo.typeSortIndex,
      groupColorSortIndex: groupInfo.colorSortIndex,
      groupColorSortKey: groupInfo.colorSortKey,
    });
  }

  if (items.length === 0) {
    if (sourceMode === SOURCE_FOLDER) {
      alert(
        "No packable imported items found.\n\n" +
          "Folder: " +
          sourceFolderPath +
          "\n" +
          "Imported objects: " +
          importedCount +
          "\n" +
          "Files processed: " +
          processedFileCount +
          "\n" +
          "Files skipped: " +
          skippedFileCount,
      );
    } else {
      alert("No packable items found in selection.");
    }
    return;
  }

  var itemGroups = [];
  if (doDraw) {
    // ----------------------------
    // Option 2: sort by color + type, then prepare bottom-strip small rows inside each group
    // ----------------------------
    items.sort(compareCollectedItems);

    var currentGroup = null;
    for (var itemIndex = 0; itemIndex < items.length; itemIndex++) {
      var groupedItem = items[itemIndex];
      if (!currentGroup || currentGroup.key !== groupedItem.groupKey) {
        currentGroup = {
          key: groupedItem.groupKey,
          labelText: groupedItem.groupLabelText,
          labelColorHex: groupedItem.groupColorHex,
          items: [],
          smallItems: [],
          regularItems: [],
        };
        itemGroups.push(currentGroup);
      }
      currentGroup.items.push(groupedItem);
    }
  } else {
    // ----------------------------
    // Option 1: pack everything as one combined pool
    // ----------------------------
    itemGroups.push({
      key: "__PACK_ONLY__",
      labelText: "",
      labelColorHex: "#000000",
      items: items.slice(0),
      smallItems: [],
      regularItems: [],
    });
  }

  var preUnplaced = [];
  for (
    var groupSortIndex = 0;
    groupSortIndex < itemGroups.length;
    groupSortIndex++
  ) {
    var groupedItems = itemGroups[groupSortIndex].items;
    var smallBucketItems = [];
    var regularItems = [];

    for (
      var groupedItemIndex = 0;
      groupedItemIndex < groupedItems.length;
      groupedItemIndex++
    ) {
      var groupedObj = groupedItems[groupedItemIndex];
      if (groupedObj.cw > GW || groupedObj.ch > GH) {
        preUnplaced.push(groupedObj.item);
        continue;
      }

      if (groupedObj.isSmallBucket) smallBucketItems.push(groupedObj);
      else regularItems.push(groupedObj);
    }

    smallBucketItems.sort(compareSmallBucketItems);
    regularItems.sort(compareDenseGroupItems);
    itemGroups[groupSortIndex].smallItems = smallBucketItems;
    itemGroups[groupSortIndex].regularItems = regularItems;
  }

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
      left: AB_L + cmToPt(col * (BOX + BOX_COL_GAP)),
      top: AB_T - cmToPt(row * (BOX + BOX_ROW_GAP)),
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
  var boxIndex = -1;
  var placedCount = 0;
  var boxesUsed = 0;
  var unplaced = preUnplaced.slice(0);
  var currentBoxPos = null;
  var currentBoxHasPlacement = false;
  var currentBoxDrawn = false;
  var currentBoxLabelText = "";
  var currentBoxLabelColorHex = "#000000";
  var freeRects = null;

  function clearCurrentBox() {
    currentBoxPos = null;
    currentBoxHasPlacement = false;
    currentBoxDrawn = false;
    currentBoxLabelText = "";
    currentBoxLabelColorHex = "#000000";
    freeRects = null;
  }

  function startNewBox(labelText, labelColorHex) {
    boxIndex++;
    currentBoxPos = boxTopLeftForIndex(boxIndex);
    currentBoxHasPlacement = false;
    currentBoxDrawn = false;
    currentBoxLabelText = labelText;
    currentBoxLabelColorHex = labelColorHex;
    freeRects = null;
  }

  function drawCurrentBoxIfNeeded() {
    if (!doDraw || currentBoxDrawn || !currentBoxPos) return;
    createBoxAt(currentBoxPos.left, currentBoxPos.top);
    createCenteredLabel(
      currentBoxLabelText,
      currentBoxLabelColorHex,
      currentBoxPos.left,
      currentBoxPos.top,
    );
    currentBoxDrawn = true;
  }

  function placeItemAtCells(obj, cellX, cellY) {
    var usable = usableOriginForBox(currentBoxPos.left, currentBoxPos.top);
    var targetX = usable.left + cmToPt(cellX * CELL + OBJ_PAD);
    var targetY = usable.bottom + cmToPt(cellY * CELL + OBJ_PAD);
    var keepFillForWholeItemRedStrokeTarget = hasRedStrokeTarget(
      obj.item,
      obj.groupType,
    );

    if (!moveItemBottomLeftTo(obj.item, targetX, targetY)) return false;
    if (!keepFillAfterPacking) {
      if (keepFillForWholeItemRedStrokeTarget) {
        try {
          removeRedStrokeFromPackedItem(obj.item, obj.groupType);
        } catch (eRemoveStroke) {}
      } else if (applyMixedGroupLasercutStyling(obj.item, obj.groupType)) {
      } else {
        try {
          removeFillFromPackedItem(obj.item);
        } catch (eRemoveFill) {}
      }
    }

    if (!currentBoxHasPlacement) {
      currentBoxHasPlacement = true;
      boxesUsed++;
    }
    drawCurrentBoxIfNeeded();
    placedCount++;
    return true;
  }

  for (var groupIndex = 0; groupIndex < itemGroups.length; groupIndex++) {
    var group = itemGroups[groupIndex];
    clearCurrentBox();

    var smallIndex = 0;
    var regularIndex = 0;

    while (
      smallIndex < group.smallItems.length ||
      regularIndex < group.regularItems.length
    ) {
      startNewBox(group.labelText, group.labelColorHex);

      var smallPack = packSmallItemsInRows(group.smallItems, smallIndex);
      for (
        var smallPlacementIndex = 0;
        smallPlacementIndex < smallPack.placements.length;
        smallPlacementIndex++
      ) {
        var smallPlacement = smallPack.placements[smallPlacementIndex];
        if (
          !placeItemAtCells(
            smallPlacement.item,
            smallPlacement.x,
            smallPlacement.y,
          )
        ) {
          unplaced.push(smallPlacement.item.item);
        }
      }
      smallIndex = smallPack.nextIndex;

      freeRects = initFreeRects(smallPack.usedHeight);

      while (regularIndex < group.regularItems.length) {
        var regularItemIndex = regularIndex;
        var obj = group.regularItems[regularItemIndex];
        var spot = findBestRectPlacement(freeRects, obj.cw, obj.ch);
        if (!spot) {
          regularItemIndex = findNextFittingRegularIndex(
            group.regularItems,
            regularIndex + 1,
            freeRects,
          );
          if (regularItemIndex < 0) break;
          obj = group.regularItems[regularItemIndex];
          spot = findBestRectPlacement(freeRects, obj.cw, obj.ch);
          if (!spot) break;
        }

        if (placeItemAtCells(obj, spot.x, spot.y)) {
          freeRects = splitFreeRects(freeRects, spot.freeRect, spot.placedRect);
        } else {
          unplaced.push(obj.item);
        }

        if (regularItemIndex === regularIndex) {
          regularIndex++;
        } else {
          group.regularItems.splice(regularItemIndex, 1);
        }
      }

      clearCurrentBox();
    }
  }

  // ----------------------------
  // Report
  // ----------------------------
  var msg = "";
  msg += "Packing finished.\n\n";
  msg += "Workflow: " + workflowModeLabel + "\n";
  msg += "Source: " + sourceModeLabel + "\n";
  if (sourceMode === SOURCE_FOLDER) {
    msg += "Folder: " + sourceFolderPath + "\n";
    msg += ".ai files found: " + aiFileCount + "\n";
    msg += "Files processed: " + processedFileCount + "\n";
    msg += "Files skipped: " + skippedFileCount + "\n";
    msg += "Imported objects: " + importedCount + "\n";
  }
  msg +=
    "Mode: " +
    (doDraw ? "Option 2 (draw box + label)" : "Option 1 (pack only)") +
    "\n";
  msg += "Placed: " + placedCount + " / " + items.length + "\n";
  msg += "Boxes used: " + boxesUsed + "\n";
  msg += "CELL: " + CELL + "cm\n";
  msg += "BOX_PAD: " + BOX_PAD + "cm\n";
  msg +=
    "Gap between objects: " + OBJ_GAP + "cm (per-side pad " + OBJ_PAD + "cm)\n";
  if (doDraw) {
    msg +=
      "Grouping: type + color (per-channel RGB tolerance " +
      COLOR_TOLERANCE +
      ")\n";
  } else {
    msg += "Grouping: single pool (Pack only ignores type + color)\n";
  }
  msg += "Usable (grid-covered): " + USE_EFF_CM + "cm\n";
  if (unplaced.length > 0) msg += "\nUnplaced items: " + unplaced.length + "\n";

  alert(msg);
})();
