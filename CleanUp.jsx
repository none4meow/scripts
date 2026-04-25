/**
 * Illustrator ExtendScript: remove empty groups, empty compound paths, and stray points.
 *
 * Runs on the active document. Locked layers and locked items are skipped.
 */
(function () {
  if (app.documents.length === 0) {
    alert("Open a document first.");
    return;
  }

  var doc = app.activeDocument;
  var stats = {
    groups: 0,
    compoundPaths: 0,
    strayPoints: 0,
    skippedLocked: 0
  };

  // Locked artwork cannot be removed safely, so skip it and report the count.
  function isLocked(item) {
    try {
      if (item.locked) return true;
    } catch (eLocked) {}
    return false;
  }

  // Locked layers block changes to all artwork inside them.
  function isLayerEditable(layer) {
    try {
      return !layer.locked;
    } catch (eLayer) {
      return false;
    }
  }

  // Remove an Illustrator item and update the matching cleanup counter.
  function removeItem(item, counterName) {
    try {
      if (isLocked(item)) {
        stats.skippedLocked++;
        return false;
      }
      item.remove();
      stats[counterName]++;
      return true;
    } catch (eRemove) {
      stats.skippedLocked++;
      return false;
    }
  }

  // Illustrator's Clean Up treats one-point paths as stray points.
  // Guides and clipping paths are left alone.
  function isStrayPoint(pathItem) {
    try {
      if (pathItem.guides || pathItem.clipping) return false;
      return pathItem.pathPoints.length <= 1;
    } catch (eStray) {
      return false;
    }
  }

  // Copy direct children into a normal array before removal.
  // Illustrator collections are live, so deleting while looping them can skip items.
  function getDirectPageItems(container) {
    var items = [];
    var i;
    var item;

    try {
      for (i = 0; i < container.pageItems.length; i++) {
        item = container.pageItems[i];
        if (item.parent === container) {
          items[items.length] = item;
        }
      }
    } catch (eItems) {}

    return items;
  }

  // Sub-layers are cleaned before artwork on the current layer.
  function getDirectLayers(layer) {
    var layers = [];
    var i;

    try {
      for (i = 0; i < layer.layers.length; i++) {
        if (layer.layers[i].parent === layer) {
          layers[layers.length] = layer.layers[i];
        }
      }
    } catch (eLayers) {}

    return layers;
  }

  // Clean stray path children first; if none remain, remove the empty compound path.
  function cleanCompoundPath(compoundPath) {
    var i;
    var pathItem;

    try {
      for (i = compoundPath.pathItems.length - 1; i >= 0; i--) {
        pathItem = compoundPath.pathItems[i];
        if (isStrayPoint(pathItem)) {
          removeItem(pathItem, "strayPoints");
        }
      }
    } catch (eCompoundChildren) {}

    try {
      if (compoundPath.pathItems.length === 0) {
        removeItem(compoundPath, "compoundPaths");
      }
    } catch (eCompoundEmpty) {}
  }

  function cleanPageItem(item) {
    var children;
    var i;

    if (isLocked(item)) {
      stats.skippedLocked++;
      return;
    }

    if (item.typename === "GroupItem") {
      // Clean deepest children first so parent groups can become empty afterward.
      children = getDirectPageItems(item);
      for (i = children.length - 1; i >= 0; i--) {
        cleanPageItem(children[i]);
      }

      try {
        if (item.pageItems.length === 0) {
          removeItem(item, "groups");
        }
      } catch (eGroupEmpty) {}
      return;
    }

    if (item.typename === "CompoundPathItem") {
      cleanCompoundPath(item);
      return;
    }

    if (item.typename === "PathItem" && isStrayPoint(item)) {
      removeItem(item, "strayPoints");
    }
  }

  // Work from the back of each layer/group to keep stacking order traversal stable.
  function cleanLayer(layer) {
    var layers;
    var items;
    var i;

    if (!isLayerEditable(layer)) {
      stats.skippedLocked++;
      return;
    }

    layers = getDirectLayers(layer);
    for (i = layers.length - 1; i >= 0; i--) {
      cleanLayer(layers[i]);
    }

    items = getDirectPageItems(layer);
    for (i = items.length - 1; i >= 0; i--) {
      cleanPageItem(items[i]);
    }
  }

  // Start with top-level layers, then recurse into each layer's contents.
  function cleanDocument(documentRef) {
    var i;

    for (i = documentRef.layers.length - 1; i >= 0; i--) {
      cleanLayer(documentRef.layers[i]);
    }
  }

  cleanDocument(doc);
  app.redraw();

  alert(
    "Clean Up complete.\n\n" +
      "Removed empty groups: " +
      stats.groups +
      "\n" +
      "Removed empty compound paths: " +
      stats.compoundPaths +
      "\n" +
      "Removed stray points: " +
      stats.strayPoints +
      "\n\n" +
      "Skipped locked items/layers: " +
      stats.skippedLocked
  );
})();
