// Frede Knudsen, frediverden.nu
// This script is used to created an outline and a cut contour path from the current selection.
// I use this to create stickers with a bleed, as the outline is offset double the amount of the cut contour.

function main() {
  /* --- INITIALIZATION --- */
  var doc = app.activeDocument;
  var sel = doc.selection;

  // If selection is empty, abort mission!
  if (!sel.length) {
    alert("Nothing selected!");
    return;
  }

  // Disable alerts
  app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS;

  /* --- VARIABLES --- */
  // Path offset in points
  var OFFSET = 2;
  // Outline fill color
  var OUTLINE_COLOR = new CMYKColor();
  OUTLINE_COLOR.cyan = 0;
  OUTLINE_COLOR.magenta = 0;
  OUTLINE_COLOR.yellow = 0;
  OUTLINE_COLOR.black = 100;
  // Cut contour stroke color
  var CUT_CONTOUR_COLOR = new CMYKColor();
  CUT_CONTOUR_COLOR.cyan = 0;
  CUT_CONTOUR_COLOR.magenta = 100;
  CUT_CONTOUR_COLOR.yellow = 0;
  CUT_CONTOUR_COLOR.black = 0;
  // Layer names
  var CUT_CONTOUR_NAME = "CUT_CONTOUR";
  var OUTLINE_NAME = "Outline";

  /* --- LAYER CREATION --- */
  // Get or create needed layers
  var cutContourLayer = findOrCreateLayer(CUT_CONTOUR_NAME);
  var outlineLayer = findOrCreateLayer(OUTLINE_NAME);

  /* --- COLLECT AND DUPLICATE SELECTED PATH ITEMS --- */
  // Create a new array for collecting path items from selection. We don't want any groups.
  var itemArray = new Array();
  // Collect all the path items into the array
  for (var i = 0; i < sel.length; i++) {
    collectPathItems(sel[i], itemArray);
  }
  // Duplicate all the collected path items to both CUT_CONTOUR layer, and Outline layer
  for (var i = 0; i < itemArray.length; i++) {
    itemArray[i].duplicate(cutContourLayer, ElementPlacement.PLACEATEND);
    itemArray[i].duplicate(outlineLayer, ElementPlacement.PLACEATEND);
  }
  // Crop any compoundPaths in the array
  doc.selection = null;
  for (var i = 0; i < cutContourLayer.pluginItems.length; i++) {
    cutContourLayer.pluginItems[i].selected = true;
  }
  for (var i = 0; i < cutContourLayer.compoundPathItems.length; i++) {
    cutContourLayer.compoundPathItems[i].selected = true;
  }
  app.executeMenuCommand("Live Pathfinder Crop");
  app.executeMenuCommand("expandStyle");
  app.executeMenuCommand("ungroup");
  app.executeMenuCommand("ungroup");
  app.executeMenuCommand("ungroup");
  app.executeMenuCommand("ungroup");
  app.executeMenuCommand("ungroup");

  /* --- OFFSET, MERGE AND STYLE CUT CONTOUR LAYER --- */
  // Offset Path effect string for Cut Contour layer.
  var xmlstringOffsetPath =
    '<LiveEffect name="Adobe Offset Path"><Dict data="R mlim 4 R ofst value I jntp 0 "/></LiveEffect>'; // 0 = round, 1=bevel, 2=miter
  xmlstringOffsetPath = xmlstringOffsetPath.replace("value", OFFSET);
  // Unselect everything
  doc.selection = null;
  // Select cut contour layer PathItems and apply offset
  for (var i = 0; i < cutContourLayer.pathItems.length; i++) {
    cutContourLayer.pathItems[i].selected = true;
    cutContourLayer.pathItems[i].filled = false;
    cutContourLayer.pathItems[i].stroked = true;
    cutContourLayer.pathItems[i].strokeWidth = 0.25;
    cutContourLayer.pathItems[i].strokeColor = CUT_CONTOUR_COLOR;
    cutContourLayer.pathItems[i].applyEffect(xmlstringOffsetPath);
  }
  // Redraw here to make sure executed commands work
  app.redraw();
  // Merge the currently selected Cut Contour layer paths
  mergeSelectedPaths();

  /* --- OFFSET, MERGE AND STYLE OUTLINE LAYER --- */
  // Offset Path effect string for Outline layer. Offset value should be doubled, for the bleed
  xmlstringOffsetPath =
    '<LiveEffect name="Adobe Offset Path"><Dict data="R mlim 4 R ofst value I jntp 0 "/></LiveEffect>'; // 0 = round, 1=bevel, 2=miter
  xmlstringOffsetPath = xmlstringOffsetPath.replace("value", OFFSET * 2);
  // Unselect everything
  doc.selection = null;
  // Select outline layer PathItems and apply offset
  for (var i = 0; i < outlineLayer.pathItems.length; i++) {
    outlineLayer.pathItems[i].selected = true;
    outlineLayer.pathItems[i].stroked = false;
    outlineLayer.pathItems[i].filled = true;
    outlineLayer.pathItems[i].fillColor = OUTLINE_COLOR;
    outlineLayer.pathItems[i].applyEffect(xmlstringOffsetPath);
  }
  // Redraw here to make sure executed commands work
  app.redraw();
  // Merge the currently selected Cut Contour layer paths
  mergeSelectedPaths();

  /* --- FINALIZATION --- */
  // Move the outline layer below everything
  outlineLayer.move(doc, ElementPlacement.PLACEATEND);
  // Reenable alerts
  app.userInteractionLevel = UserInteractionLevel.DISPLAYALERTS;
}

function mergeSelectedPaths() {
  // This will always merge the paths in the current selection.
  app.executeMenuCommand("group");
  app.executeMenuCommand("Live Pathfinder Add");
  app.executeMenuCommand("expandStyle");
  app.executeMenuCommand("ungroup");
}

function collectPathItems(item, itemArray) {
  if (item.typename == "GroupItem") {
    // If item is a GroupItem
    for (var i = 0; i < item.pathItems.length; i++) {
      // Push each pathItem to array
      itemArray.push(item.pathItems[i]);
    }
    for (var i = 0; i < item.compoundPathItems.length; i++) {
      // Push each pathItem to array
      itemArray.push(item.compoundPathItems[i]);
    }
    for (var i = 0; i < item.pluginItems.length; i++) {
      // Push each pathItem to array
      itemArray.push(item.pluginItems[i]);
    }
    for (var i = 0; i < item.groupItems.length; i++) {
      // For each nested GroupItem inside GroupItem, run collectPathItems recursively
      collectPathItems(item.groupItems[i], itemArray);
    }
  } else if (
    item.typename == "PathItem" ||
    item.typename == "CompoundPathItem" ||
    item.typename == "PluginItem"
  ) {
    itemArray.push(item);
  }
}

function findOrCreateLayer(layerName) {
  try {
    var myLayer = app.activeDocument.layers.getByName(layerName);
  } catch (err) {
    myLayer = app.activeDocument.layers.add();
    myLayer.name = layerName;
  } finally {
    return myLayer;
  }
}

main();
