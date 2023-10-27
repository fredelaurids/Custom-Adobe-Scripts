function main() {
  // 1. Create new layers
  // 2. Duplicate all items to layers
  // 3. Offset and merge items in new layers
  // 5. Set colors of items in new layers
  // 5. Move outline layer to bottom

  // Get active document and selection
  var doc = app.activeDocument;
  var sel = doc.selection;

  // Check if anything is selected
  if (!sel.length) {
    alert("Nothing selected. Aborting.");
    return;
  }

  // Disable alerts
  var originalInteractionLevel = app.userInteractionLevel;
  app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS;

  /* --- VARIABLES --- */
  // Path offset in points
  var OFFSET_VALUE = prompt("Outline thickness in points", "2");
  // Outline fill color
  var OUTLINE_COLOR = new CMYKColor();
  OUTLINE_COLOR.cyan = 0;
  OUTLINE_COLOR.magenta = 0;
  OUTLINE_COLOR.yellow = 0;
  OUTLINE_COLOR.black = 100;
  // Cut contour stroke color
  // var CUT_CONTOUR_COLOR = new CMYKColor();
  // CUT_CONTOUR_COLOR.cyan = 0;
  // CUT_CONTOUR_COLOR.magenta = 100;
  // CUT_CONTOUR_COLOR.yellow = 0;
  // CUT_CONTOUR_COLOR.black = 0;

  var CUT_CONTOUR_COLOR = findOrCreateColor("CutContour");

  // Layer names
  var CUT_CONTOUR_NAME = prompt("Cut contour layer name", "CutContour");
  var OUTLINE_NAME = prompt("Outline layer name", "Outline");

  // Create new layers
  var cutContourLayer = findOrCreateLayer(CUT_CONTOUR_NAME);
  var outlineLayer = findOrCreateLayer(OUTLINE_NAME);

  // Duplicate selection into cutContourLayer and outlineLayer
  for (var i = 0; i < sel.length; i++) {
    sel[i].duplicate(cutContourLayer, ElementPlacement.PLACEATEND);
    sel[i].duplicate(outlineLayer, ElementPlacement.PLACEATEND);
  }

  // Check if anything has a clipping mask
  cropClippedGroups(cutContourLayer);
  cropClippedGroups(outlineLayer);

  // Merge all the stuff in the cutContourLayer
  offsetAndMerge(cutContourLayer, OFFSET_VALUE);
  offsetAndMerge(outlineLayer, OFFSET_VALUE * 2);

  // Set color of cut contour layer
  doc.selection = null;
  cutContourLayer.hasSelectedArtwork = true;
  // Set cut contour stroke and fill
  for (var i = 0; i < doc.selection.length; i++) {
    if (doc.selection[i].typename == "PathItem") {
      doc.selection[i].filled = false;
      doc.selection[i].stroked = true;
      doc.selection[i].strokeWidth = 0.25;
      doc.selection[i].strokeColor = CUT_CONTOUR_COLOR;
    } else {
      doc.selection[i].pathItems[0].filled = false;
      doc.selection[i].pathItems[0].stroked = true;
      doc.selection[i].pathItems[0].strokeWidth = 0.25;
      doc.selection[i].pathItems[0].strokeColor = CUT_CONTOUR_COLOR;
    }
  }

  // Set color of outline layer
  doc.selection = null;
  outlineLayer.hasSelectedArtwork = true;
  // Set cut contour stroke and fill
  for (var i = 0; i < doc.selection.length; i++) {
    if (doc.selection[i].typename == "PathItem") {
      doc.selection[i].filled = true;
      doc.selection[i].stroked = false;
      doc.selection[i].fillColor = OUTLINE_COLOR;
    } else {
      doc.selection[i].pathItems[0].filled = true;
      doc.selection[i].pathItems[0].stroked = false;
      doc.selection[i].pathItems[0].fillColor = OUTLINE_COLOR;
    }
  }

  // Move the outline layer below everything
  outlineLayer.move(doc, ElementPlacement.PLACEATEND);

  // Redraw here to make sure executed commands work
  app.redraw();

  // Reenable alerts
  app.userInteractionLevel = originalInteractionLevel;
}

function offsetAndMerge(layer, offsetValue) {
  // Unselect everything
  app.activeDocument.selection = null;

  // Select cutContourLayer
  layer.hasSelectedArtwork = true;

  // Group the new selection
  app.executeMenuCommand("group");

  // Create an Offset Path effect
  var xmlstringOffsetPath =
    '<LiveEffect name="Adobe Offset Path"><Dict data="R mlim 4 R ofst value I jntp 0 "/></LiveEffect>'; // 0 = round, 1=bevel, 2=miter
  xmlstringOffsetPath = xmlstringOffsetPath.replace("value", offsetValue);

  // Apply the effect to the selected group
  layer.groupItems[0].applyEffect(xmlstringOffsetPath);

  // Add the paths together
  app.executeMenuCommand("Live Pathfinder Add");

  // Expand the Appereance effects/styles
  app.executeMenuCommand("expandStyle");

  // Ungroup
  app.executeMenuCommand("ungroup");
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

function findOrCreateColor(colorName) {
  try {
    var newSpotColor = app.activeDocument.swatches.getByName(colorName);
    newSpotColor.remove();
  } catch(err) {
    // whatever
  } finally {
    var newColor = new CMYKColor();
    newColor.cyan = 0;
    newColor.magenta = 100;
    newColor.yellow = 0;
    newColor.black = 0;
  
    var newSpot = app.activeDocument.spots.add();
    newSpot.name = colorName;
    newSpot.colorType = ColorModel.SPOT;
    newSpot.color = newColor;
  
    newSpotColor = new SpotColor();
    newSpotColor.spot = newSpot;
    newSpotColor.tint = 100;

    return newSpotColor;
  }
}

function cropClippedGroups(layer) {
  app.activeDocument.selection = null;

  var clippedGroups = new Array();

  // Find all groups with a clipping mask
  for (var i = 0; i < layer.groupItems.length; i++) {
    collectClippedGroups(layer.groupItems[0], clippedGroups);
  }

  // Add clipped groups to selection
  for (var i = 0; i < clippedGroups.length; i++) {
    clippedGroups[i].selected = true;
  }

  // Crop the group
  app.executeMenuCommand("Live Pathfinder Crop");
  app.executeMenuCommand("expandStyle");
  app.executeMenuCommand("ungroup");
}

function collectClippedGroups(item, itemArray) {
  if (item.typename == "GroupItem") {
    // If item is a GroupItem and has a clipping mask
    for (var i = 0; i < item.groupItems.length; i++) {
      // For each nested GroupItem inside GroupItem, run collectPathItems recursively
      collectClippedGroups(item.groupItems[i], itemArray);
    }
    if (item.clipped) {
      itemArray.push(item);
    }
  }
}

main();
